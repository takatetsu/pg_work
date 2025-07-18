# 組織アクセス制御仕様書 - orgtblによる権限管理

## 1. 概要

orgtbl（組織権限テーブル）は、ユーザーがどの組織のデータにアクセスできるかを制御する重要なテーブルです。管理区分（manageclass）により、3つの異なる権限レベルを提供します。

**重要**: 組織コード（orgcode）は並列の関係にあり、階層構造を持ちません。各組織への権限は個別に設定する必要があります。

## 2. orgtbl テーブル構造

| カラム名 | 型 | 説明 |
|---------|---|------|
| personalcode | char(5) | 管理者の個人コード |
| orgcode | char(6) | 管理対象の組織コード |
| manageclass | char(1) | 管理区分 |

### 2.1 管理区分（manageclass）の定義

| 値 | 権限名称 | 説明 | 利用画面 |
|----|---------|------|---------|
| 0 | 支店控除担当 | 支店控除データの入力権限 | inputdeduction.asp |
| 1 | 全体入力担当者 | 全社員の勤務データ一括入力権限 | inputall.asp |
| 2 | 上長 | 部下の勤務承認権限 | checklist.asp, workstatus.asp |

### 2.2 データ構造例

```sql
-- 例：山田太郎(10001)が複数の権限を持つ場合

|personalcode | orgcode | manageclass|
|-------------|---------|------------|
|10001        | 111000  | 2  -- 東京支店の上長|
|10001        | 112000  | 2  -- 仙台支店の上長|
|10001        | 110000  | 1  -- 東日本営業部の全体入力担当|
|10001        | 111000  | 0  -- 東京支店の支店控除担当|
```

## 3. 画面別アクセス制御仕様

### 3.1 workstatus.asp（勤務状況確認画面）

#### 機能
部門の勤務状況を一覧表示

#### アクセス制御ロジック
```sql
-- ユーザーが上長として登録されている組織の職員を表示
SELECT s.personalcode, s.staffname, s.orgcode, s.gradecode, s.workshift, s.is_operator 
FROM orgtbl o 
RIGHT OUTER JOIN stafftbl s ON o.orgcode = s.orgcode 
WHERE s.is_input = '1'        -- 勤務入力対象者
  AND s.is_enable = '1'        -- 有効な職員
  AND o.personalcode = :user_id
  AND o.manageclass = '2'      -- 上長権限

UNION 

-- ユーザーと同じ組織の職員も表示
SELECT s.personalcode, s.staffname, s.orgcode, s.gradecode, s.workshift, s.is_operator 
FROM stafftbl o 
RIGHT OUTER JOIN stafftbl s ON o.orgcode = s.orgcode 
WHERE s.is_input = '1' 
  AND s.is_enable = '1' 
  AND o.personalcode = :user_id 

ORDER BY s.orgcode, s.gradecode DESC, s.personalcode
```

#### 表示仕様
- 自組織の職員 + 上長権限を持つ組織の職員を表示
- 等級順（降順）→ 個人コード順でソート
- 勤務状況を色分け表示（出勤/休暇/未入力）

### 3.2 inputdeduction.asp（支店控除入力画面）

#### 機能
支店単位での控除金額（駐車場代、社宅費等）の入力

#### アクセス制御ロジック
```sql
SELECT s.personalcode, s.staffname 
FROM orgtbl o
RIGHT OUTER JOIN stafftbl s ON o.orgcode = s.orgcode 
WHERE s.is_enable = '1'         -- 有効な職員
  AND o.personalcode = :user_id
  AND o.manageclass = '0'       -- 支店控除担当権限
ORDER BY s.orgcode, s.gradecode DESC, s.personalcode
```

#### 入力可能項目
- 全労済（火災共済）控除額
- 全労済（交通災害）控除額
- 駐車場代控除額
- 社宅共益費控除額
- 水道代控除額
- 合格祝金控除額
- 支部費（組合）控除額
- その他控除（3項目まで、名称指定可）

### 3.3 inputall.asp（勤務表全体入力画面）

#### 機能
管理組織の全職員の勤務データを一括入力

#### アクセス制御ロジック
```sql
SELECT * FROM 
  -- ユーザーが全体入力権限を持つ組織を取得
  (SELECT orgcode FROM orgtbl 
   WHERE personalcode = :user_id AND manageclass = '1') ORG 
LEFT JOIN 
  -- その組織の職員情報を取得
  (SELECT personalcode AS pcode, staffname, orgcode AS org, gradecode AS grade 
   FROM stafftbl WHERE is_enable = '1') STAFF 
ON ORG.orgcode = STAFF.org 
LEFT JOIN 
  -- 月次集計データを結合
  (SELECT * FROM dutyrostertbl WHERE ymb = :target_month) DUTY 
ON STAFF.pcode = DUTY.personalcode 
LEFT JOIN 
  -- 承認済み件数を集計
  (SELECT personalcode AS countpcode, COUNT(*) AS count 
   FROM worktbl 
   WHERE workingdate LIKE :target_month + '%' AND is_approval = '1' 
   GROUP BY personalcode) APPROVAL 
ON STAFF.pcode = APPROVAL.countpcode 
WHERE pcode IS NOT NULL 
ORDER BY STAFF.org ASC, STAFF.grade DESC, STAFF.pcode ASC
```

#### 入力制御
- 未承認データが存在する場合は入力不可
- 過去月のデータは編集不可
- 一括で以下の項目を入力可能：
  - 出勤日数、欠勤日数
  - 有給休暇、特別休暇、保存休暇
  - 残業時間、休日出勤
  - 宿直、日直回数

### 3.4 checklist.asp（勤務表確認・承認画面）

#### 機能
部下の勤務表を確認し、承認処理を実行

#### アクセス制御ロジック
```sql
SELECT s.personalcode, s.staffname, s.orgcode, s.gradecode 
FROM orgtbl o
RIGHT OUTER JOIN stafftbl s ON o.orgcode = s.orgcode 
WHERE s.is_input = '1'          -- 勤務入力対象者
  AND s.is_enable = '1'         -- 有効な職員
  AND o.personalcode = :user_id
  AND o.manageclass = '2'       -- 上長権限
ORDER BY s.orgcode, s.gradecode DESC, s.personalcode
```

#### 承認仕様
- 承認状態を記号で表示：
  - ○：承認済み
  - ×：未承認
  - －：承認不要（管理職等級033以上）
- 個別承認と一括承認が可能
- 承認後は本人による編集不可

## 4. 特殊ケース処理

### 4.1 等級による承認除外
```python
# 承認不要の判定
def is_approval_exempt(grade_code):
    # 等級033以上または等級000は承認不要
    return int(grade_code) >= 33 or grade_code == "000"
```

### 4.2 複数組織管理
- 1人のユーザーが複数の組織に対して異なる権限を持つことが可能
- 例：東京支店の上長 かつ 東日本営業部の全体入力担当者

### 4.3 組織コードの扱い
- 各組織コード（orgcode）は独立した並列の組織を表す
- 上位組織の権限を持っていても、下位組織への権限は自動的に付与されない
- 例：東日本営業部（110000）の権限を持っていても、東京支店（111000）の権限は別途必要
- 複数組織の管理が必要な場合は、orgtblに組織ごとのレコードを明示的に登録する必要がある

## 5. セキュリティ実装

### 5.1 SQLインジェクション対策
```asp
' パラメータバインディングの使用
cmd.CommandText = "SELECT * FROM orgtbl WHERE personalcode = ? AND manageclass = ?"
cmd.Parameters.Append cmd.CreateParameter("personalcode", adChar, adParamInput, 5, Session("MM_Username"))
cmd.Parameters.Append cmd.CreateParameter("manageclass", adChar, adParamInput, 1, "2")
```

### 5.2 セッションチェック
```asp
' 各画面でのセッション確認
If Session("MM_Username") = "" Then
    Response.Redirect "index.asp"
End If

' 権限チェック
Set rs = getOrgPermissions(Session("MM_Username"), requiredManageClass)
If rs.EOF Then
    Response.Write "アクセス権限がありません"
    Response.End
End If
```

### 5.3 監査ログ
- 誰が、いつ、どの組織のデータにアクセスしたかを記録
- 承認操作の履歴を保持

## 6. 実装ガイドライン

### 6.1 新システムでの実装方針

#### データベース設計
```sql
-- PostgreSQLでの実装例
CREATE TABLE organization_permission (
    id SERIAL PRIMARY KEY,
    personal_code VARCHAR(5) NOT NULL,
    org_code VARCHAR(6) NOT NULL,
    manage_class SMALLINT NOT NULL CHECK (manage_class IN (0, 1, 2)),
    valid_from DATE DEFAULT CURRENT_DATE,
    valid_to DATE,
    created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    
    CONSTRAINT uk_org_permission UNIQUE (personal_code, org_code, manage_class),
    CONSTRAINT fk_staff FOREIGN KEY (personal_code) REFERENCES staff(personal_code),
    CONSTRAINT fk_org FOREIGN KEY (org_code) REFERENCES organization(org_code)
);

-- インデックス
CREATE INDEX idx_org_perm_personal ON organization_permission(personal_code);
CREATE INDEX idx_org_perm_org ON organization_permission(org_code);
CREATE INDEX idx_org_perm_valid ON organization_permission(valid_from, valid_to);
```

#### アプリケーション層での実装
```python
class OrganizationPermission:
    """組織権限管理クラス"""
    
    MANAGE_CLASS = {
        'BRANCH_DEDUCTION': 0,  # 支店控除担当
        'GENERAL_INPUT': 1,     # 全体入力担当者
        'SUPERIOR': 2           # 上長
    }
    
    @classmethod
    def get_accessible_employees(cls, user_id, manage_class, include_same_org=False):
        """
        アクセス可能な職員リストを取得
        
        Args:
            user_id: ユーザーID
            manage_class: 管理区分
            include_same_org: 同一組織の職員を含むか
            
        Returns:
            職員リスト
        """
        query = """
        SELECT DISTINCT s.personal_code, s.staff_name, s.org_code
        FROM organization_permission op
        JOIN staff s ON op.org_code = s.org_code
        WHERE op.personal_code = %s
          AND op.manage_class = %s
          AND s.is_enable = true
          AND CURRENT_DATE BETWEEN COALESCE(op.valid_from, '1900-01-01') 
                               AND COALESCE(op.valid_to, '9999-12-31')
        """
        
        if include_same_org:
            query += """
            UNION
            SELECT s2.personal_code, s2.staff_name, s2.org_code
            FROM staff s1
            JOIN staff s2 ON s1.org_code = s2.org_code
            WHERE s1.personal_code = %s
              AND s2.is_enable = true
            """
        
        query += " ORDER BY org_code, personal_code"
        
        # 実行処理...
```

### 6.2 権限チェックミドルウェア
```python
def require_org_permission(manage_class):
    """組織権限チェックデコレータ"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            user_id = session.get('user_id')
            if not user_id:
                return redirect('/login')
            
            # 権限チェック
            permissions = OrganizationPermission.get_user_permissions(user_id)
            if manage_class not in permissions:
                abort(403)  # Forbidden
            
            # 権限情報をコンテキストに追加
            g.org_permissions = permissions
            return f(*args, **kwargs)
        
        return decorated_function
    return decorator

# 使用例
@app.route('/workstatus')
@require_org_permission(OrganizationPermission.MANAGE_CLASS['SUPERIOR'])
def work_status():
    # 上長権限が必要な処理
    pass
```

## 7. 移行時の注意事項

1. **データ移行**
   - orgtblのデータは完全に移行
   - 有効期限（valid_from/valid_to）を追加して履歴管理

2. **権限の統合**
   - 複数の権限を持つユーザーの扱い
   - UIでの権限切り替え機能の実装

3. **パフォーマンス**
   - 大規模組織でのクエリ最適化
   - 権限情報のキャッシング

4. **監査対応**
   - アクセスログの記録
   - 権限変更履歴の保持