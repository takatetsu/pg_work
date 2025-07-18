# データベース仕様書 - 勤務表管理システム

## 1. データベース移行概要

### 1.1 移行方針

- **移行元**: Microsoft SQL Server
- **移行先**: PostgreSQL 15
- **移行戦略**: 
  - テーブル構造は基本的に維持（正規化改善を含む）
  - SQL Server固有機能はPostgreSQL互換機能で代替
  - データ型は適切にマッピング
  - 文字コード: Shift-JIS/Windows-31J → UTF-8

### 1.2 主要な変更点

1. **データ型の最適化**: char型からvarchar型への変更
2. **タイムスタンプ管理**: timestamp型からtimestamptzへ
3. **インデックス戦略**: 検索性能を考慮した複合インデックス追加
4. **制約の強化**: 外部キー制約の追加
5. **パーティショニング**: 大量データテーブルの年月パーティション化

## 2. SQL ServerからPostgreSQLへのデータ型マッピング

### 2.1 基本データ型マッピング表

| SQL Server | PostgreSQL | 変換ルール | 備考 |
|------------|------------|-----------|------|
| char(n) | varchar(n) | 末尾空白除去 | 固定長→可変長で効率化 |
| varchar(n) | varchar(n) | そのまま | |
| int | integer | そのまま | |
| int identity | serial/bigserial | シーケンス使用 | 自動採番 |
| float(53) | double precision | そのまま | |
| timestamp | timestamptz | タイムゾーン付加 | JSTとして扱う |
| datetime | timestamptz | タイムゾーン付加 | |
| bit | boolean | 0→false, 1→true | |
| nvarchar(n) | varchar(n) | そのまま | UTF-8で統一 |
| text | text | そのまま | |
| decimal(p,s) | numeric(p,s) | そのまま | |

### 2.2 特殊な変換が必要なケース

```sql
-- SQL Server: timestamp型（自動更新バイナリ）
-- PostgreSQL: 更新日時として実装
CREATE OR REPLACE FUNCTION update_modified_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ language 'plpgsql';

CREATE TRIGGER update_timestamp BEFORE UPDATE ON table_name
    FOR EACH ROW EXECUTE FUNCTION update_modified_column();
```

## 3. テーブル設計書

### 3.1 スキーマ構成

```sql
-- スキーマ作成
CREATE SCHEMA attendance;
CREATE SCHEMA master;
CREATE SCHEMA archive;

-- 権限設定
GRANT USAGE ON SCHEMA attendance TO app_user;
GRANT USAGE ON SCHEMA master TO app_user;
GRANT SELECT ON SCHEMA archive TO app_user;
```

### 3.2 マスタ系テーブル

#### 3.2.1 職員マスタ（master.staff）

```sql
CREATE TABLE master.staff (
    id SERIAL PRIMARY KEY,
    personal_code VARCHAR(5) UNIQUE NOT NULL,
    staff_name VARCHAR(30) NOT NULL,
    staff_name_kana VARCHAR(60),  -- 新規追加：カナ
    email VARCHAR(100),            -- 新規追加：メールアドレス
    org_code VARCHAR(6) NOT NULL,
    grade_code VARCHAR(3),
    is_operator BOOLEAN DEFAULT false,
    is_input BOOLEAN DEFAULT true,
    is_charge BOOLEAN DEFAULT false,
    is_superior BOOLEAN DEFAULT false,
    is_deduction BOOLEAN DEFAULT false,
    is_enable BOOLEAN DEFAULT true,
    password_hash VARCHAR(255),    -- bcryptハッシュ用に拡張
    last_login_at TIMESTAMPTZ,     -- 新規追加：最終ログイン
    processed_ymb VARCHAR(6),
    holiday_type CHAR(1) DEFAULT '1',
    is_union_executive BOOLEAN DEFAULT false,
    grant_date VARCHAR(4),
    work_shift SMALLINT DEFAULT 0,
    old_work_shift SMALLINT,
    old_work_shift_last_ymb VARCHAR(6) DEFAULT '000000',
    base_am_workmin INTEGER,
    base_pm_workmin INTEGER,
    hire_date DATE,                -- 新規追加：入社日
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT chk_work_shift CHECK (work_shift IN (0,1,2,3,9)),
    CONSTRAINT chk_holiday_type CHECK (holiday_type IN ('1','2'))
);

-- インデックス
CREATE INDEX idx_staff_org_code ON master.staff(org_code);
CREATE INDEX idx_staff_is_enable ON master.staff(is_enable);
CREATE INDEX idx_staff_email ON master.staff(email);
```

#### 3.2.2 組織マスタ（master.organization）

```sql
CREATE TABLE master.organization (
    id SERIAL PRIMARY KEY,
    org_code VARCHAR(6) UNIQUE NOT NULL,
    org_name VARCHAR(100) NOT NULL,
    is_active BOOLEAN DEFAULT true,
    display_order INTEGER,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
);

-- 注：現行システムでは組織は並列構造のため、階層管理は実装しない
-- 各組織は独立しており、権限は組織ごとに個別設定する

CREATE INDEX idx_org_code ON master.organization(org_code);
CREATE INDEX idx_org_active ON master.organization(is_active);
```

#### 3.2.3 権限管理テーブル（master.org_permission）

```sql
-- orgtblを正規化して作成
CREATE TABLE master.org_permission (
    id SERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    manage_class SMALLINT NOT NULL,
    org_code VARCHAR(6) NOT NULL,
    is_active BOOLEAN DEFAULT true,
    valid_from DATE DEFAULT CURRENT_DATE,
    valid_to DATE,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code),
    CONSTRAINT fk_org FOREIGN KEY (org_code) 
        REFERENCES master.organization(org_code),
    CONSTRAINT chk_manage_class CHECK (manage_class IN (0,1,2)),
    CONSTRAINT chk_valid_period CHECK (valid_to IS NULL OR valid_to >= valid_from)
);

CREATE INDEX idx_org_permission_personal ON master.org_permission(personal_code);
CREATE INDEX idx_org_permission_org ON master.org_permission(org_code);
```

#### 3.2.4 祝日マスタ（master.holiday）

```sql
CREATE TABLE master.holiday (
    id SERIAL PRIMARY KEY,
    holiday_date DATE UNIQUE NOT NULL,
    holiday_type CHAR(1) NOT NULL DEFAULT '1',
    holiday_name VARCHAR(100) NOT NULL,  -- memoから名称変更
    is_national BOOLEAN DEFAULT true,    -- 新規：国民の祝日フラグ
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT chk_holiday_type CHECK (holiday_type IN ('1','2'))
);

CREATE INDEX idx_holiday_date ON master.holiday(holiday_date);
CREATE INDEX idx_holiday_year ON master.holiday(EXTRACT(YEAR FROM holiday_date));
```

### 3.3 勤務データ系テーブル

#### 3.3.1 勤務記録テーブル（attendance.work_record）

```sql
CREATE TABLE attendance.work_record (
    id BIGSERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    working_date DATE NOT NULL,
    -- 勤務区分
    morning_work CHAR(1),
    afternoon_work CHAR(1),
    morning_holiday CHAR(1),
    afternoon_holiday CHAR(1),
    summons CHAR(1),
    -- 勤務時間
    work_begin TIME,
    work_end TIME,
    break_begin1 TIME,
    break_end1 TIME,
    break_begin2 TIME,
    break_end2 TIME,
    -- 時間外勤務
    overtime_begin TIME,
    overtime_end TIME,
    rest_begin TIME,
    rest_end TIME,
    -- 計算済み時間（分単位）
    work_minutes INTEGER DEFAULT 0,
    overtime_minutes INTEGER DEFAULT 0,
    late_night_minutes INTEGER DEFAULT 0,
    holiday_minutes INTEGER DEFAULT 0,
    holiday_overtime_minutes INTEGER DEFAULT 0,
    holiday_late_minutes INTEGER DEFAULT 0,
    week_overtime_minutes INTEGER DEFAULT 0,
    -- 申請時間
    request_time_minutes INTEGER,
    request_time_begin TIME,
    request_time_end TIME,
    vacation_time_minutes INTEGER,
    vacation_time_begin TIME,
    vacation_time_end TIME,
    -- その他
    night_duty SMALLINT DEFAULT 0,
    day_duty SMALLINT DEFAULT 0,
    operator CHAR(1),
    memo VARCHAR(100),
    memo_code VARCHAR(2),
    -- ステータス
    is_approval BOOLEAN DEFAULT false,
    is_error BOOLEAN DEFAULT false,
    error_code VARCHAR(10),
    error_message VARCHAR(200),
    -- メタデータ
    created_by VARCHAR(5),
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_by VARCHAR(5),
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT uk_work_record UNIQUE (personal_code, working_date),
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code)
) PARTITION BY RANGE (working_date);

-- 月次パーティション作成（例：2024年1月）
CREATE TABLE attendance.work_record_y2024m01 
    PARTITION OF attendance.work_record
    FOR VALUES FROM ('2024-01-01') TO ('2024-02-01');

-- インデックス（各パーティションに自動適用）
CREATE INDEX idx_work_record_personal_date 
    ON attendance.work_record(personal_code, working_date);
CREATE INDEX idx_work_record_approval 
    ON attendance.work_record(is_approval);
CREATE INDEX idx_work_record_error 
    ON attendance.work_record(is_error) WHERE is_error = true;
```

#### 3.3.2 月次集計テーブル（attendance.monthly_summary）

```sql
CREATE TABLE attendance.monthly_summary (
    id SERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    year_month VARCHAR(6) NOT NULL,
    -- 日数集計
    work_days NUMERIC(4,1) DEFAULT 0,
    holiday_work_days NUMERIC(4,1) DEFAULT 0,
    absence_days NUMERIC(4,1) DEFAULT 0,
    paid_leave_days NUMERIC(4,1) DEFAULT 0,
    preserve_leave_days NUMERIC(4,1) DEFAULT 0,
    special_leave_days NUMERIC(4,1) DEFAULT 0,
    compensatory_days NUMERIC(4,1) DEFAULT 0,
    actual_work_days NUMERIC(4,1) DEFAULT 0,
    -- 時間集計（分単位で保存）
    total_work_minutes INTEGER DEFAULT 0,
    overtime_minutes INTEGER DEFAULT 0,
    late_night_minutes INTEGER DEFAULT 0,
    holiday_work_minutes INTEGER DEFAULT 0,
    holiday_overtime_minutes INTEGER DEFAULT 0,
    week_overtime_minutes INTEGER DEFAULT 0,
    -- 手当関連
    short_work_count INTEGER DEFAULT 0,
    night_duty_a INTEGER DEFAULT 0,
    night_duty_b INTEGER DEFAULT 0,
    night_duty_c INTEGER DEFAULT 0,
    night_duty_d INTEGER DEFAULT 0,
    day_duty_count INTEGER DEFAULT 0,
    -- 休暇残
    vacation_balance NUMERIC(4,1),
    compensatory_balance NUMERIC(4,1),
    -- 基準時間
    base_work_minutes INTEGER,
    current_work_minutes INTEGER,
    -- ステータス
    is_closed BOOLEAN DEFAULT false,
    closed_at TIMESTAMPTZ,
    closed_by VARCHAR(5),
    -- メタデータ
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT uk_monthly_summary UNIQUE (personal_code, year_month),
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code)
);

CREATE INDEX idx_monthly_summary_year_month ON attendance.monthly_summary(year_month);
CREATE INDEX idx_monthly_summary_closed ON attendance.monthly_summary(is_closed);
```

#### 3.3.3 タイムカードテーブル（attendance.timecard）

```sql
CREATE TABLE attendance.timecard (
    id BIGSERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    punch_datetime TIMESTAMPTZ NOT NULL,
    punch_type SMALLINT NOT NULL, -- 1:出社, 2:外出, 3:戻り, 4:退社
    device_id VARCHAR(50),
    ip_address INET,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code),
    CONSTRAINT chk_punch_type CHECK (punch_type IN (1,2,3,4))
) PARTITION BY RANGE (punch_datetime);

-- 月次パーティション
CREATE TABLE attendance.timecard_y2024m01 
    PARTITION OF attendance.timecard
    FOR VALUES FROM ('2024-01-01') TO ('2024-02-01');

CREATE INDEX idx_timecard_personal_date 
    ON attendance.timecard(personal_code, punch_datetime);
```

### 3.4 休暇管理系テーブル

#### 3.4.1 休暇付与テーブル（attendance.leave_grant）

```sql
CREATE TABLE attendance.leave_grant (
    id SERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    leave_type SMALLINT NOT NULL, -- 1:有給, 2:代休, 3:特別, 4:保存
    grant_date DATE NOT NULL,
    expire_date DATE,
    granted_days NUMERIC(4,1) NOT NULL,
    used_days NUMERIC(4,1) DEFAULT 0,
    balance_days NUMERIC(4,1) NOT NULL,
    grant_reason VARCHAR(100),
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code),
    CONSTRAINT chk_leave_type CHECK (leave_type IN (1,2,3,4)),
    CONSTRAINT chk_days CHECK (granted_days >= 0 AND used_days >= 0 AND balance_days >= 0)
);

CREATE INDEX idx_leave_grant_personal ON attendance.leave_grant(personal_code);
CREATE INDEX idx_leave_grant_expire ON attendance.leave_grant(expire_date);
```

#### 3.4.2 休暇使用履歴テーブル（attendance.leave_usage）

```sql
CREATE TABLE attendance.leave_usage (
    id BIGSERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    leave_grant_id INTEGER NOT NULL,
    usage_date DATE NOT NULL,
    usage_type SMALLINT NOT NULL, -- 1:全日, 2:午前半休, 3:午後半休, 4:時間休
    usage_days NUMERIC(3,1),
    usage_hours NUMERIC(3,1),
    work_record_id BIGINT,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code),
    CONSTRAINT fk_leave_grant FOREIGN KEY (leave_grant_id) 
        REFERENCES attendance.leave_grant(id),
    CONSTRAINT chk_usage_type CHECK (usage_type IN (1,2,3,4))
);

CREATE INDEX idx_leave_usage_personal_date 
    ON attendance.leave_usage(personal_code, usage_date);
```

### 3.5 支店控除系テーブル

#### 3.5.1 支店控除テーブル（attendance.branch_deduction）

```sql
CREATE TABLE attendance.branch_deduction (
    id SERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    year_month VARCHAR(6) NOT NULL,
    -- 固定項目
    zenrosai_fire INTEGER DEFAULT 0,        -- 全労済（火災共済）
    zenrosai_traffic INTEGER DEFAULT 0,     -- 全労済（交通災害）
    parking_fee INTEGER DEFAULT 0,          -- 駐車場代
    dormitory_fee INTEGER DEFAULT 0,        -- 社宅共益費
    water_fee INTEGER DEFAULT 0,            -- 水道代
    congratulation_repay INTEGER DEFAULT 0, -- 合格祝金返済
    union_fee INTEGER DEFAULT 0,            -- 支部費（組合）
    -- 可変項目（JSONで柔軟に対応）
    other_deductions JSONB DEFAULT '[]'::jsonb,
    -- メタデータ
    created_by VARCHAR(5),
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_by VARCHAR(5),
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT uk_branch_deduction UNIQUE (personal_code, year_month),
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) 
        REFERENCES master.staff(personal_code)
);

-- JSONBのGINインデックス
CREATE INDEX idx_branch_deduction_other ON attendance.branch_deduction 
    USING GIN (other_deductions);

-- other_deductionsの構造例
-- [
--   {"name": "その他控除1", "amount": 5000},
--   {"name": "その他控除2", "amount": 3000}
-- ]
```

### 3.6 システム管理系テーブル

#### 3.6.1 システム制御テーブル（master.system_control）

```sql
CREATE TABLE master.system_control (
    id INTEGER PRIMARY KEY DEFAULT 1,
    system_enable BOOLEAN DEFAULT true,
    maintenance_mode BOOLEAN DEFAULT false,
    maintenance_message VARCHAR(200),
    data_lock_year_month VARCHAR(6),  -- この年月以前のデータは編集不可
    config JSONB DEFAULT '{}'::jsonb,  -- その他の設定値
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_by VARCHAR(5),
    
    CONSTRAINT chk_single_row CHECK (id = 1)
);

-- 1行のみ許可
CREATE UNIQUE INDEX idx_system_control_single ON master.system_control((1));
```

#### 3.6.2 操作ログテーブル（master.audit_log）

```sql
CREATE TABLE master.audit_log (
    id BIGSERIAL PRIMARY KEY,
    log_timestamp TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    personal_code VARCHAR(5),
    action_type VARCHAR(20) NOT NULL,  -- LOGIN, LOGOUT, UPDATE, DELETE等
    target_table VARCHAR(50),
    target_id VARCHAR(50),
    old_values JSONB,
    new_values JSONB,
    ip_address INET,
    user_agent TEXT,
    session_id VARCHAR(100),
    result VARCHAR(10),  -- SUCCESS, FAILURE
    error_message TEXT
) PARTITION BY RANGE (log_timestamp);

-- 月次パーティション
CREATE TABLE master.audit_log_y2024m01 
    PARTITION OF master.audit_log
    FOR VALUES FROM ('2024-01-01') TO ('2024-02-01');

-- インデックス
CREATE INDEX idx_audit_log_personal ON master.audit_log(personal_code);
CREATE INDEX idx_audit_log_action ON master.audit_log(action_type);
CREATE INDEX idx_audit_log_timestamp ON master.audit_log(log_timestamp);
```

### 3.7 ビューとマテリアライズドビュー

#### 3.7.1 勤務状況ビュー

```sql
CREATE VIEW attendance.v_work_status AS
SELECT 
    w.personal_code,
    s.staff_name,
    o.org_name,
    w.working_date,
    CASE 
        WHEN w.morning_work = '1' AND w.afternoon_work = '1' THEN '出勤'
        WHEN w.morning_work = '1' THEN '午前出勤'
        WHEN w.afternoon_work = '1' THEN '午後出勤'
        WHEN w.morning_holiday IS NOT NULL OR w.afternoon_holiday IS NOT NULL THEN '休暇'
        ELSE '欠勤'
    END AS work_status,
    w.work_minutes,
    w.overtime_minutes,
    w.is_approval,
    w.is_error
FROM attendance.work_record w
JOIN master.staff s ON w.personal_code = s.personal_code
JOIN master.organization o ON s.org_code = o.org_code
WHERE w.working_date >= CURRENT_DATE - INTERVAL '1 month';
```

#### 3.7.2 月次集計マテリアライズドビュー

```sql
CREATE MATERIALIZED VIEW attendance.mv_monthly_statistics AS
SELECT 
    year_month,
    org_code,
    COUNT(DISTINCT personal_code) as employee_count,
    AVG(total_work_minutes) as avg_work_minutes,
    AVG(overtime_minutes) as avg_overtime_minutes,
    SUM(CASE WHEN overtime_minutes > 2700 THEN 1 ELSE 0 END) as over_45h_count,
    SUM(CASE WHEN is_error THEN 1 ELSE 0 END) as error_count
FROM attendance.monthly_summary ms
JOIN master.staff s ON ms.personal_code = s.personal_code
GROUP BY year_month, org_code;

CREATE INDEX idx_mv_monthly_stats ON attendance.mv_monthly_statistics(year_month, org_code);
```

## 4. インデックス設計

### 4.1 インデックス設計方針

1. **主キー・ユニークキー**: 自動的にインデックス作成
2. **外部キー**: 参照整合性とJOIN性能のため作成
3. **検索条件**: WHERE句で頻繁に使用される列
4. **ソート条件**: ORDER BY句で使用される列
5. **部分インデックス**: 特定条件のデータのみ対象

### 4.2 主要インデックス一覧

```sql
-- 複合インデックス（頻繁なクエリパターン用）
CREATE INDEX idx_work_record_monthly 
    ON attendance.work_record(personal_code, working_date DESC);

CREATE INDEX idx_work_record_approval_check 
    ON attendance.work_record(is_approval, working_date DESC) 
    WHERE is_approval = false;

CREATE INDEX idx_staff_active 
    ON master.staff(is_enable, org_code) 
    WHERE is_enable = true;

-- カバリングインデックス（インデックスのみで結果を返す）
CREATE INDEX idx_monthly_summary_report 
    ON attendance.monthly_summary(year_month, personal_code) 
    INCLUDE (overtime_minutes, total_work_minutes, is_closed);
```

## 5. 制約設計

### 5.1 制約一覧

```sql
-- ドメイン制約
CREATE DOMAIN year_month AS VARCHAR(6) 
    CHECK (VALUE ~ '^\d{6}$');

CREATE DOMAIN time_hhmm AS VARCHAR(4) 
    CHECK (VALUE ~ '^([0-4]\d|5[0-7])[0-5]\d$');

-- チェック制約の例
ALTER TABLE attendance.work_record 
    ADD CONSTRAINT chk_work_time_order 
    CHECK (work_begin < work_end);

ALTER TABLE attendance.work_record 
    ADD CONSTRAINT chk_break_time_order 
    CHECK (break_begin1 < break_end1);

-- トリガーによる複雑な制約
CREATE OR REPLACE FUNCTION check_overtime_limit()
RETURNS TRIGGER AS $$
DECLARE
    monthly_overtime INTEGER;
BEGIN
    -- 月間残業時間チェック
    SELECT SUM(overtime_minutes) INTO monthly_overtime
    FROM attendance.work_record
    WHERE personal_code = NEW.personal_code
    AND DATE_TRUNC('month', working_date) = DATE_TRUNC('month', NEW.working_date)
    AND id != COALESCE(NEW.id, -1);
    
    IF monthly_overtime + NEW.overtime_minutes > 2700 THEN -- 45時間
        NEW.is_error := true;
        NEW.error_code := 'OT001';
        NEW.error_message := '月間残業時間が45時間を超えています';
    END IF;
    
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_check_overtime_limit
BEFORE INSERT OR UPDATE ON attendance.work_record
FOR EACH ROW EXECUTE FUNCTION check_overtime_limit();
```

## 6. データ移行

### 6.1 移行手順

```sql
-- 1. 文字コード変換とデータクレンジング
-- SQL Serverからのエクスポート時にUTF-8変換

-- 2. 一時テーブルへのロード
CREATE SCHEMA migration;

CREATE TABLE migration.worktbl_temp AS 
SELECT * FROM (VALUES (NULL)) AS dummy WHERE false;
-- COPY コマンドでCSVインポート

-- 3. データ変換とロード
INSERT INTO attendance.work_record (
    personal_code,
    working_date,
    morning_work,
    afternoon_work,
    -- ... その他のカラム
    overtime_minutes,  -- 時間から分に変換
    created_at,
    updated_at
)
SELECT 
    TRIM(personalcode),
    TO_DATE(workingdate, 'YYYYMMDD'),
    morningwork,
    afternoonwork,
    -- ... その他のカラム
    CASE 
        WHEN overtime ~ '^\d{4}$' THEN 
            (SUBSTRING(overtime, 1, 2)::INT * 60 + SUBSTRING(overtime, 3, 2)::INT)
        ELSE 0
    END,
    CURRENT_TIMESTAMP,
    CURRENT_TIMESTAMP
FROM migration.worktbl_temp;

-- 4. データ検証
-- 件数チェック
SELECT COUNT(*) FROM migration.worktbl_temp;
SELECT COUNT(*) FROM attendance.work_record;

-- サンプルデータ比較
SELECT * FROM attendance.work_record 
ORDER BY RANDOM() LIMIT 100;
```

### 6.2 パスワード移行戦略

```python
# 初回ログイン時の移行処理
def migrate_password_on_login(username, password):
    user = get_user(username)
    
    if user.password_hash.startswith('sha1:'):
        # 旧SHA1ハッシュ
        old_hash = user.password_hash[5:]
        if verify_sha1(password, old_hash):
            # 新しいbcryptハッシュに更新
            new_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
            update_user_password(username, new_hash)
            return True
    else:
        # 既にbcrypt
        return bcrypt.checkpw(password.encode('utf-8'), user.password_hash)
    
    return False
```

## 7. パフォーマンス最適化

### 7.1 パーティショニング戦略

```sql
-- 自動パーティション作成関数
CREATE OR REPLACE FUNCTION create_monthly_partitions()
RETURNS void AS $$
DECLARE
    start_date DATE;
    end_date DATE;
    partition_name TEXT;
BEGIN
    -- 今月と来月のパーティションを作成
    FOR i IN 0..1 LOOP
        start_date := DATE_TRUNC('month', CURRENT_DATE) + (i || ' month')::INTERVAL;
        end_date := start_date + INTERVAL '1 month';
        
        -- work_record
        partition_name := 'work_record_y' || TO_CHAR(start_date, 'YYYY') || 
                         'm' || TO_CHAR(start_date, 'MM');
        
        EXECUTE format('
            CREATE TABLE IF NOT EXISTS attendance.%I 
            PARTITION OF attendance.work_record
            FOR VALUES FROM (%L) TO (%L)',
            partition_name, start_date, end_date);
            
        -- timecard
        partition_name := 'timecard_y' || TO_CHAR(start_date, 'YYYY') || 
                         'm' || TO_CHAR(start_date, 'MM');
        
        EXECUTE format('
            CREATE TABLE IF NOT EXISTS attendance.%I 
            PARTITION OF attendance.timecard
            FOR VALUES FROM (%L) TO (%L)',
            partition_name, start_date, end_date);
    END LOOP;
END;
$$ LANGUAGE plpgsql;

-- 月次実行（cronで自動化）
SELECT create_monthly_partitions();
```

### 7.2 統計情報とVACUUM戦略

```sql
-- 自動VACUUM設定
ALTER TABLE attendance.work_record SET (
    autovacuum_vacuum_scale_factor = 0.1,
    autovacuum_analyze_scale_factor = 0.05
);

-- 統計情報の精度向上
ALTER TABLE attendance.work_record 
    ALTER COLUMN personal_code SET STATISTICS 1000;
```

## 8. バックアップとリストア

### 8.1 バックアップ戦略

```bash
#!/bin/bash
# 日次バックアップスクリプト

DATE=$(date +%Y%m%d)
BACKUP_DIR="/backup/postgresql"

# フルバックアップ（日曜日）
if [ $(date +%w) -eq 0 ]; then
    pg_dump -h localhost -U postgres -d attendance_db \
        -F custom -b -v -f "${BACKUP_DIR}/full_${DATE}.backup"
fi

# 増分バックアップ（WALアーカイブ）
pg_basebackup -h localhost -U replication -D "${BACKUP_DIR}/base_${DATE}" \
    -F tar -z -P -v

# 古いバックアップの削除（30日以前）
find ${BACKUP_DIR} -name "*.backup" -mtime +30 -delete
```

### 8.2 リストア手順

```bash
# フルリストア
pg_restore -h localhost -U postgres -d attendance_db_restore \
    -v /backup/postgresql/full_20240131.backup

# ポイントインタイムリカバリ
# recovery.conf設定
restore_command = 'cp /backup/postgresql/wal/%f %p'
recovery_target_time = '2024-01-31 15:30:00'
```

## 9. セキュリティ設計

### 9.1 行レベルセキュリティ（RLS）

```sql
-- 行レベルセキュリティの有効化
ALTER TABLE attendance.work_record ENABLE ROW LEVEL SECURITY;

-- ポリシーの作成
CREATE POLICY work_record_personal ON attendance.work_record
    FOR ALL
    TO app_user
    USING (personal_code = current_setting('app.current_user')::VARCHAR);

CREATE POLICY work_record_superior ON attendance.work_record
    FOR SELECT
    TO app_user
    USING (
        personal_code IN (
            SELECT s.personal_code 
            FROM master.staff s
            WHERE s.org_code IN (
                SELECT org_code 
                FROM master.org_permission
                WHERE personal_code = current_setting('app.current_user')::VARCHAR
                AND manage_class = 2  -- 上長
            )
        )
    );
```

### 9.2 暗号化

```sql
-- 個人情報の暗号化
CREATE EXTENSION IF NOT EXISTS pgcrypto;

-- 暗号化関数
CREATE OR REPLACE FUNCTION encrypt_pii(text_value TEXT)
RETURNS TEXT AS $$
BEGIN
    RETURN encode(
        encrypt(
            text_value::BYTEA, 
            current_setting('app.encryption_key')::BYTEA, 
            'aes'
        ), 
        'base64'
    );
END;
$$ LANGUAGE plpgsql;

-- 復号化関数
CREATE OR REPLACE FUNCTION decrypt_pii(encrypted_value TEXT)
RETURNS TEXT AS $$
BEGIN
    RETURN convert_from(
        decrypt(
            decode(encrypted_value, 'base64'), 
            current_setting('app.encryption_key')::BYTEA, 
            'aes'
        ), 
        'UTF8'
    );
END;
$$ LANGUAGE plpgsql;
```