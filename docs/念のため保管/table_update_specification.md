# テーブル更新仕様書 - worktblとdutyrostertbl

## 1. worktbl（勤務記録テーブル）更新仕様

### 1.1 更新タイミング
- **個人勤務入力時**: 日々の勤務入力保存時
- **一括入力時**: 人事担当者による一括入力時
- **タイムカード反映時**: タイムカードデータ取込時

### 1.2 入力フィールド（ユーザー入力値をそのまま格納）

| フィールド名 | 型 | 説明 | 入力値 |
|------------|---|------|--------|
| personalcode | char(5) | 個人コード | ログインユーザーのID |
| workingdate | char(8) | 勤務日 | YYYYMMDD形式 |
| morningwork | char(1) | 午前勤務区分 | 0:なし, 1:振替出勤, 2:休日出勤, 4:通常勤務, 9:勤務 |
| afternoonwork | char(1) | 午後勤務区分 | 同上 |
| morningholiday | char(1) | 午前休暇区分 | 1:公休日, 2:振替休日, 3:有給, 4:代休, 5:特休, 6:保存休, 7:欠勤, 9:コアタイム有休, A:法定休日 |
| afternoonholiday | char(1) | 午後休暇区分 | 同上 |
| summons | char(1) | 呼出区分 | 0:なし, 1:通常, 2:深夜 |
| overtime_begin | char(4) | 時間外開始 | HHMM形式 |
| overtime_end | char(4) | 時間外終了 | HHMM形式 |
| rest_begin | char(4) | 時間外休憩開始 | HHMM形式 |
| rest_end | char(4) | 時間外休憩終了 | HHMM形式 |
| requesttime_begin | char(4) | 時間代休開始 | HHMM形式 |
| requesttime_end | char(4) | 時間代休終了 | HHMM形式 |
| latetime_begin | char(4) | 深夜割増開始 | HHMM形式 |
| latetime_end | char(4) | 深夜割増終了 | HHMM形式 |
| vacationtime_begin | char(4) | 時間有給開始 | HHMM形式 |
| vacationtime_end | char(4) | 時間有給終了 | HHMM形式 |
| nightduty | char(1) | 宿直 | 0:なし, 1:責任者, 2:処理者 |
| dayduty | char(1) | 日直 | 0:なし, 1:責任者, 2:処理者, 3:土曜出番 |
| operator | char(1) | シフト区分 | 0:なし, 1:甲番, 2:乙番, 3-8:各種シフト |
| memo | char(100) | 備考 | 自由入力 |
| memo2 | char(2) | メモ2 | typetblの値 |
| weekovertime | char(4) | 週超過時間 | HHMM形式（非フレックスのみ） |

### 1.3 フレックスタイム専用フィールド

| フィールド名 | 型 | 説明 | 入力値 |
|------------|---|------|--------|
| work_begin | char(4) | 勤務開始 | HHMM形式 |
| work_end | char(4) | 勤務終了 | HHMM形式 |
| break_begin1 | char(4) | 休憩1開始 | HHMM形式 |
| break_end1 | char(4) | 休憩1終了 | HHMM形式 |
| break_begin2 | char(4) | 中抜け開始 | HHMM形式 |
| break_end2 | char(4) | 中抜け終了 | HHMM形式 |

### 1.4 計算フィールド（システム自動計算）

#### 1.4.1 基本時間計算

```python
# requesttime - 時間代休時間数
def calculate_requesttime(requesttime_begin, requesttime_end):
    if not requesttime_begin or not requesttime_end:
        return "0000"
    
    # フレックス勤務者は常に"0000"
    if is_flex_worker:
        return "0000"
    
    # 時間差分計算
    minutes = time_diff_minutes(requesttime_begin, requesttime_end)
    
    # 昼休み控除（12:00-13:00を跨ぐ場合60分減算）
    if overlaps_lunch(requesttime_begin, requesttime_end):
        minutes -= 60
    
    return minutes_to_hhmm(minutes)

# vacationtime - 時間有給時間数
def calculate_vacationtime(vacationtime_begin, vacationtime_end):
    if not vacationtime_begin or not vacationtime_end:
        return "0000"
    
    minutes = time_diff_minutes(vacationtime_begin, vacationtime_end)
    
    # 昼休み控除
    if overlaps_lunch(vacationtime_begin, vacationtime_end):
        minutes -= 60
    
    return minutes_to_hhmm(minutes)

# latetime - 深夜割増時間数
def calculate_latetime(latetime_begin, latetime_end, operator_type):
    # オペレータ乙番は自動設定
    if operator_type == "2" and not latetime_begin:
        return "0700"  # 22:00-05:00
    
    if not latetime_begin or not latetime_end:
        return "0000"
    
    return minutes_to_hhmm(time_diff_minutes(latetime_begin, latetime_end))
```

#### 1.4.2 残業時間計算（compOverTime関数）

```python
def compOverTime(work_type, holiday_type, overtime_begin, overtime_end, 
                 rest_begin, rest_end, standard_hours=460):
    """
    残業時間を4種類に分類して計算
    
    戻り値: {
        'overtime': 普通残業,
        'overtimelate': 普通残業深夜,
        'holidayshift': 休日出勤,
        'holidayshiftovertime': 休日時間外,
        'holidayshiftlate': 休日深夜,
        'holidayshiftovertimelate': 休日時間外深夜
    }
    """
    
    # 総労働時間計算
    total_minutes = time_diff_minutes(overtime_begin, overtime_end)
    rest_minutes = time_diff_minutes(rest_begin, rest_end) if rest_begin else 0
    work_minutes = total_minutes - rest_minutes
    
    # 昼休み控除
    if overlaps_lunch(overtime_begin, overtime_end):
        work_minutes -= 60
    
    # 深夜時間（22:00-05:00）の計算
    late_night_minutes = calculate_late_night_overlap(overtime_begin, overtime_end)
    normal_minutes = work_minutes - late_night_minutes
    
    result = {
        'overtime': "0000",
        'overtimelate': "0000",
        'holidayshift': "0000",
        'holidayshiftovertime': "0000",
        'holidayshiftlate': "0000",
        'holidayshiftovertimelate': "0000"
    }
    
    # 休日出勤の場合
    if work_type in ["2", "6"]:  # 休日出勤
        if work_minutes <= standard_hours:
            # 7時間40分以内
            result['holidayshift'] = minutes_to_hhmm(normal_minutes)
            result['holidayshiftlate'] = minutes_to_hhmm(late_night_minutes)
        else:
            # 7時間40分超過
            overtime = work_minutes - standard_hours
            
            # 深夜時間の配分
            if late_night_minutes <= overtime:
                result['holidayshift'] = minutes_to_hhmm(standard_hours)
                result['holidayshiftovertime'] = minutes_to_hhmm(normal_minutes - standard_hours)
                result['holidayshiftovertimelate'] = minutes_to_hhmm(late_night_minutes)
            else:
                base_late = late_night_minutes - overtime
                result['holidayshift'] = minutes_to_hhmm(normal_minutes)
                result['holidayshiftlate'] = minutes_to_hhmm(base_late)
                result['holidayshiftovertimelate'] = minutes_to_hhmm(overtime)
    
    # 通常勤務・振替出勤の場合
    else:
        result['overtime'] = minutes_to_hhmm(normal_minutes)
        result['overtimelate'] = minutes_to_hhmm(late_night_minutes)
    
    return result
```

#### 1.4.3 フレックスタイム勤務時間計算

```python
def calculate_flex_workmin(work_begin, work_end, break_begin1, break_end1,
                          break_begin2, break_end2, morning_holiday, afternoon_holiday):
    """
    フレックスタイム勤務者の実労働時間（分）計算
    """
    if not work_begin or not work_end:
        return 0
    
    # 基本勤務時間
    total_minutes = time_diff_minutes(work_begin, work_end)
    
    # 休憩時間控除
    if break_begin1 and break_end1:
        total_minutes -= time_diff_minutes(break_begin1, break_end1)
    
    if break_begin2 and break_end2:
        total_minutes -= time_diff_minutes(break_begin2, break_end2)
    
    # 半日休暇の場合の調整
    if morning_holiday in ["3", "5", "6", "7"]:  # 午前休暇
        total_minutes = min(total_minutes, 210)  # 午後のみ3.5時間
    elif afternoon_holiday in ["3", "5", "6", "7"]:  # 午後休暇
        total_minutes = min(total_minutes, 250)  # 午前のみ4時間10分
    
    return total_minutes
```

#### 1.4.4 エラーフラグ設定

```python
def set_error_flag(total_overtime_minutes, grade_code):
    """
    労働時間エラーフラグ設定
    - 管理職（等級033以上）は除外
    - それ以外は違反チェック
    """
    if int(grade_code) >= 33:
        return "0"
    
    # エラー条件（実装により異なる）
    if total_overtime_minutes > 840:  # 14時間超
        return "1"
    
    return "0"
```

### 1.5 更新処理フロー

```sql
-- 1. 既存レコードチェック
SELECT id FROM worktbl 
WHERE personalcode = @personalcode AND workingdate = @workingdate

-- 2A. 存在する場合：UPDATE
UPDATE worktbl SET
    morningwork = @morningwork,
    afternoonwork = @afternoonwork,
    -- ... 全フィールド
    overtime = @calculated_overtime,
    overtimelate = @calculated_overtimelate,
    -- ... 計算フィールド
    updatetime = CURRENT_TIMESTAMP
WHERE id = @id

-- 2B. 存在しない場合：INSERT
INSERT INTO worktbl (
    personalcode, workingdate, morningwork, afternoonwork,
    -- ... 全フィールド
    overtime, overtimelate,
    -- ... 計算フィールド
    is_approval
) VALUES (
    @personalcode, @workingdate, @morningwork, @afternoonwork,
    -- ... 値
    @calculated_overtime, @calculated_overtimelate,
    -- ... 計算値
    '0'  -- 承認フラグは初期値0
)
```

## 2. dutyrostertbl（月次集計テーブル）更新仕様

### 2.1 更新タイミング
- worktbl更新の都度、該当月の全データを再集計
- 月次締め処理時の最終集計

### 2.2 集計処理フロー

```python
def update_monthly_summary(personal_code, year_month):
    """
    月次集計更新メイン処理
    """
    # 1. 該当月の全勤務データ取得
    work_records = get_monthly_work_records(personal_code, year_month)
    
    # 2. 初期化
    summary = initialize_summary_record(personal_code, year_month)
    
    # 3. 日次データ集計
    for record in work_records:
        aggregate_daily_record(summary, record)
    
    # 4. 月間計算
    calculate_monthly_totals(summary)
    
    # 5. 残日数計算
    calculate_remaining_balances(summary)
    
    # 6. DB更新
    upsert_monthly_summary(summary)
```

### 2.3 フィールド別集計ロジック

#### 2.3.1 勤務日数関連

```python
# workdays - 可出勤日数
def calculate_workdays(year_month, holidays):
    days_in_month = get_days_in_month(year_month)
    holiday_count = count_holidays_in_month(holidays, year_month)
    return days_in_month - holiday_count

# realworkdays - 実出勤日数
def calculate_realworkdays(work_records):
    days = 0.0
    for record in work_records:
        if record.morningwork in ["1", "4", "5", "9"]:
            days += 0.5
        if record.afternoonwork in ["1", "4", "5", "9"]:
            days += 0.5
        
        # 休日出勤は時間から日数換算
        if record.holidayshift_minutes > 0:
            days += record.holidayshift_minutes / 460.0  # 7時間40分 = 1日
    
    return round(days, 1)

# 各種休暇日数
def calculate_leave_days(work_records):
    leave_counts = {
        'workholidays': 0.0,      # 代休
        'paidvacations': 0.0,     # 有給
        'preservevacations': 0.0,  # 保存休
        'specialvacations': 0.0,   # 特休
        'absencedays': 0.0        # 欠勤
    }
    
    for record in work_records:
        # 午前分
        if record.morningholiday == "4":
            leave_counts['workholidays'] += 0.5
        elif record.morningholiday == "3":
            leave_counts['paidvacations'] += 0.5
        elif record.morningholiday == "6":
            leave_counts['preservevacations'] += 0.5
        elif record.morningholiday == "5":
            leave_counts['specialvacations'] += 0.5
        elif record.morningholiday == "7":
            leave_counts['absencedays'] += 0.5
        elif record.morningholiday == "9":  # コアタイム有休
            leave_counts['paidvacations'] += 0.25
        
        # 午後分（同様の処理）
        # ...
    
    return leave_counts
```

#### 2.3.2 時間集計（10進数時間で格納）

```python
# 時間文字列(HHMM)を10進数時間に変換
def hhmm_to_decimal(hhmm):
    """
    "0130" → 1.5 (1時間30分)
    "0050" → 0.9 (50分は0.9時間として計算)
    """
    if not hhmm or hhmm == "0000":
        return 0.0
    
    hours = int(hhmm[:2])
    minutes = int(hhmm[2:])
    
    # 特殊な10分単位変換
    decimal_minutes = {
        0: 0.0, 10: 0.2, 20: 0.4,
        30: 0.5, 40: 0.7, 50: 0.9
    }
    
    return hours + decimal_minutes.get(minutes, minutes/60.0)

# overtime - 時間外集計
def calculate_total_overtime(work_records, is_flex=False):
    if is_flex:
        # フレックスタイムは月間清算
        total_work_min = sum(r.workmin for r in work_records)
        required_min = get_monthly_required_minutes(year_month)
        overtime_min = max(0, total_work_min - required_min)
        return minutes_to_decimal_hours(overtime_min)
    else:
        # 通常勤務は日次累計
        total = 0.0
        for record in work_records:
            total += hhmm_to_decimal(record.overtime)
        return total

# その他の時間集計
def aggregate_time_fields(work_records):
    return {
        'overtime': sum(hhmm_to_decimal(r.overtime) for r in work_records),
        'overtimelate': sum(hhmm_to_decimal(r.overtimelate) for r in work_records),
        'holidayshifttime': sum(hhmm_to_decimal(r.holidayshift) for r in work_records),
        'holidayshiftovertime': sum(hhmm_to_decimal(r.holidayshiftovertime) for r in work_records),
        'holidayshiftlate': sum(hhmm_to_decimal(r.holidayshiftlate) for r in work_records),
        'holidayshiftovertimelate': sum(hhmm_to_decimal(r.holidayshiftovertimelate) for r in work_records),
        'workholidaytime': sum(hhmm_to_decimal(r.requesttime) for r in work_records),
        'latepremium': sum(hhmm_to_decimal(r.latetime) for r in work_records),
        'weekovertime': sum(hhmm_to_decimal(r.weekovertime) for r in work_records)
    }
```

#### 2.3.3 カウント集計

```python
# 宿直・日直・呼出等のカウント
def count_duties(work_records):
    counts = {
        'nightduty_a': 0, 'nightduty_b': 0,
        'nightduty_c': 0, 'nightduty_d': 0,
        'dayduty': 0,
        'summons': 0, 'summonslate': 0
    }
    
    for record in work_records:
        # 宿直（A-Dは別ロジックで判定）
        if record.nightduty == "1":
            counts['nightduty_a'] += 1
        elif record.nightduty == "2":
            counts['nightduty_b'] += 1
        
        # 日直
        if record.dayduty in ["1", "2", "3"]:
            counts['dayduty'] += 1
        
        # 呼出
        if record.summons == "1":
            counts['summons'] += 1
        elif record.summons == "2":
            counts['summonslate'] += 1
    
    return counts

# シフト勤務カウント
def count_shift_work(work_records):
    shifts = {
        'shiftwork_kou': 0.0,   # 甲番
        'shiftwork_otsu': 0.0,  # 乙番
        'shiftwork_hei': 0.0,   # 丙番
        'shiftwork_a': 0.0,
        'shiftwork_b': 0.0
    }
    
    for record in work_records:
        if record.operator == "1":  # 甲番
            if record.morningwork and record.afternoonwork:
                shifts['shiftwork_kou'] += 1.0
            else:
                shifts['shiftwork_kou'] += 0.5
        elif record.operator == "2":  # 乙番
            # 特殊計算ロジック
            pass
    
    return shifts
```

#### 2.3.4 残数計算

```python
# vacationnumber - 有給休暇残
def calculate_vacation_balance(personal_code, year_month, used_this_month):
    # 前月残
    prev_balance = get_previous_vacation_balance(personal_code, year_month)
    
    # 新規付与チェック
    new_grant = check_new_vacation_grant(personal_code, year_month)
    
    # 今月使用
    used = used_this_month  # paidvacations + vacationtime/460
    
    return prev_balance + new_grant - used

# holidaynumber - 代休残
def calculate_holiday_balance(personal_code, year_month, earned, used):
    prev_balance = get_previous_holiday_balance(personal_code, year_month)
    
    # 獲得：休日出勤8時間で1日
    earned_days = earned / 8.0
    
    return prev_balance + earned_days - used
```

#### 2.3.5 フレックスタイム専用

```python
# workingmins - 実労働時間（分）
def calculate_working_minutes(work_records):
    return sum(r.workmin for r in work_records)

# currentworkmin - 当月必要労働時間（分）
def calculate_required_minutes(year_month, personal_code):
    base_minutes = get_base_work_minutes(year_month)  # 基準時間
    
    # 休暇による減算
    for record in work_records:
        if record.morningholiday in ["3", "5", "6"]:  # 有給等
            base_minutes -= 250  # 午前4時間10分
        if record.afternoonholiday in ["3", "5", "6"]:
            base_minutes -= 210  # 午後3時間30分
    
    return base_minutes

# legalholiday_extra_min - 法定休日割増時間
def calculate_legal_holiday_minutes(work_records):
    total = 0
    for record in work_records:
        if record.morningholiday == "A" or record.afternoonholiday == "A":
            # 法定休日の勤務時間
            total += record.workmin
    return total
```

### 2.4 特殊処理

#### 2.4.1 オペレータ（シフト勤務者）

```python
def process_operator_records(record):
    # 乙番は自動的に深夜7時間
    if record.operator == "2":
        record.latetime = "0700"
    
    # 半日カウントの特殊処理
    if record.operator in ["3", "4", "5"]:
        # 午前/午後で0.5日カウント
        pass
```

#### 2.4.2 コミュニケータ

```python
def process_communicator_records(work_records, workshift):
    # 土曜/平日を分けて集計
    saturday_min = 0
    weekday_min = 0
    
    for record in work_records:
        if is_saturday(record.workingdate):
            saturday_min += record.workmin
        else:
            weekday_min += record.workmin
    
    # 残業計算は人事で別途実施
    return {
        'saturday_workmin': saturday_min,
        'weekdays_workmin': weekday_min,
        'overtime': 0.0  # 0固定
    }
```

### 2.5 データベース更新SQL

```sql
-- 既存レコード確認
SELECT id FROM dutyrostertbl 
WHERE personalcode = @personalcode AND ymb = @year_month

-- 更新処理（UPSERT）
MERGE dutyrostertbl AS target
USING (
    SELECT 
        @personalcode as personalcode,
        @year_month as ymb,
        @workdays as workdays,
        -- ... 全集計値
) AS source
ON target.personalcode = source.personalcode AND target.ymb = source.ymb
WHEN MATCHED THEN
    UPDATE SET
        workdays = source.workdays,
        -- ... 全フィールド更新
        updatetime = CURRENT_TIMESTAMP
WHEN NOT MATCHED THEN
    INSERT (personalcode, ymb, workdays, ...)
    VALUES (source.personalcode, source.ymb, source.workdays, ...);
```

## 3. トランザクション管理

```python
def save_work_record(form_data):
    """
    勤務記録保存のトランザクション処理
    """
    try:
        # トランザクション開始
        begin_transaction()
        
        # 1. 入力値検証
        errors = validate_input(form_data)
        if errors:
            rollback_transaction()
            return {'success': False, 'errors': errors}
        
        # 2. 計算処理
        calculated_fields = calculate_all_fields(form_data)
        
        # 3. worktbl更新
        work_id = upsert_worktbl(form_data, calculated_fields)
        
        # 4. dutyrostertbl再集計
        update_monthly_summary(form_data['personalcode'], form_data['year_month'])
        
        # 5. 監査ログ
        log_audit_trail(form_data['personalcode'], 'UPDATE_WORK', work_id)
        
        # コミット
        commit_transaction()
        return {'success': True, 'work_id': work_id}
        
    except Exception as e:
        rollback_transaction()
        log_error(e)
        return {'success': False, 'error': str(e)}
```

## 4. パフォーマンス考慮事項

1. **インデックス利用**
   - personalcode + workingdate での検索最適化
   - personalcode + ymb での月次集計最適化

2. **バッチ更新**
   - 一括入力時は複数レコードをまとめて処理
   - 月次集計は1回のクエリで実行

3. **キャッシュ活用**
   - 休暇残数は月初に計算してキャッシュ
   - 基準労働時間は年度初めに計算

4. **非同期処理**
   - 36協定チェックは別スレッドで実行
   - 大量データ更新時は進捗表示