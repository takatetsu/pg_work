# 機能仕様書補遺 - 勤務入力チェック詳細実装仕様

## 1. チェック実行タイミングとフロー

### 1.1 リアルタイムチェック（onChange/onBlur）
```
フィールド値変更時
├─ 時刻フォーマット検証
├─ ペア入力検証（開始・終了）
├─ 時系列検証
└─ 即座にエラー表示/クリア
```

### 1.2 行単位チェック（行フォーカスアウト時）
```
行の編集完了時
├─ 勤務区分組み合わせチェック
├─ 時間計算（実働、残業、深夜）
├─ 休憩時間法定チェック
└─ 行単位エラーフラグ設定
```

### 1.3 月次一括チェック（保存/申請時）
```
保存ボタン押下時
├─ 全行の再検証
├─ 休暇残日数チェック
├─ 36協定累計チェック
├─ 法定休日チェック（フレックス）
└─ エラーサマリー表示
```

## 2. 具体的なチェックロジック実装詳細

### 2.1 時刻フォーマット検証関数
```python
def validate_time_format(time_str):
    """
    時刻入力の検証と正規化
    
    入力例:
    - "830" → "08:30" (変換成功)
    - "8:30" → "08:30" (変換成功)
    - "0830" → "08:30" (変換成功)
    - "2500" → エラー（範囲外）
    - "12:70" → エラー（分が無効）
    - "" → None（空白許可）
    """
    if not time_str or time_str.strip() == "":
        return None, True  # 空白は有効
    
    # 数字のみの場合
    if time_str.isdigit():
        if len(time_str) <= 2:
            # "8" → "08:00"
            hours = int(time_str)
            minutes = 0
        elif len(time_str) == 3:
            # "830" → "08:30"
            hours = int(time_str[0])
            minutes = int(time_str[1:3])
        elif len(time_str) == 4:
            # "0830" → "08:30"
            hours = int(time_str[0:2])
            minutes = int(time_str[2:4])
        else:
            return None, False
    
    # コロン区切りの場合
    elif ":" in time_str:
        parts = time_str.split(":")
        if len(parts) != 2:
            return None, False
        try:
            hours = int(parts[0])
            minutes = int(parts[1])
        except ValueError:
            return None, False
    else:
        return None, False
    
    # 範囲チェック（翌日対応: 00:00-47:59）
    if hours < 0 or hours > 47:
        return None, False
    if minutes < 0 or minutes > 59:
        return None, False
    
    # 正規化された時刻を返す
    return f"{hours:02d}:{minutes:02d}", True
```

### 2.2 勤務区分組み合わせチェックマトリックス
```python
# 有効な組み合わせを定義
VALID_COMBINATIONS = {
    # (午前出勤区分, 午後出勤区分, 午前休暇区分, 午後休暇区分): 有効性
    ("1", "1", None, None): True,  # 通常勤務
    ("1", "1", "1", "1"): False,   # エラー01: 公休日に勤務
    ("2", "2", "1", "1"): True,    # 公休日に振替出勤
    ("3", "3", "1", "1"): True,    # 公休日に休日出勤（非フレックスのみ）
    ("2", "3", None, None): False, # エラー04: 振替と休出の混在
    # ... 他の組み合わせ
}

def check_work_type_combination(morning_work, afternoon_work, 
                               morning_holiday, afternoon_holiday,
                               is_flex=False):
    """
    勤務区分の組み合わせチェック
    
    戻り値: (is_valid, error_code, error_message)
    """
    # エラー01: 休暇区分の不一致チェック
    if morning_holiday != afternoon_holiday:
        if morning_holiday not in ["1", "A"] and afternoon_holiday not in ["1", "A"]:
            return False, "01", "午前と午後の休暇等区分の値が違います"
    
    # エラー54: フレックス勤務者の公休日休日出勤
    if is_flex and morning_holiday == "1" and morning_work == "3":
        return False, "54", "フレックス勤務者は公休日に休日出勤を入力できません"
    
    # 他のチェックロジック...
    return True, None, None
```

### 2.3 時間計算の詳細実装
```python
def calculate_work_hours_detailed(work_start, work_end, break_start, break_end,
                                overtime_start=None, overtime_end=None,
                                rest_start=None, rest_end=None):
    """
    勤務時間の詳細計算
    
    戻り値: {
        'actual_minutes': 480,      # 実働時間（分）
        'overtime_minutes': 60,     # 残業時間（分）
        'late_night_minutes': 30,   # 深夜時間（分）
        'holiday_minutes': 0,       # 休日勤務時間（分）
        'break_minutes': 60,        # 休憩時間（分）
        'lunch_deduction': 60       # 昼休み控除（分）
    }
    """
    result = {}
    
    # 1. 総勤務時間計算
    total_minutes = time_diff_minutes(work_start, work_end)
    
    # 2. 休憩時間計算
    break_minutes = 0
    if break_start and break_end:
        break_minutes = time_diff_minutes(break_start, break_end)
    
    # 3. 昼休み自動控除チェック
    lunch_deduction = 0
    if overlaps_lunch_time(work_start, work_end):
        lunch_deduction = 60
    
    # 4. 実働時間
    actual_minutes = total_minutes - break_minutes - lunch_deduction
    
    # 5. 残業時間（8時間超過分、10分単位切り捨て）
    overtime_minutes = max(0, actual_minutes - 480)
    overtime_minutes = (overtime_minutes // 10) * 10
    
    # 6. 深夜時間計算
    late_night_minutes = calculate_late_night_minutes(work_start, work_end)
    
    return {
        'actual_minutes': actual_minutes,
        'overtime_minutes': overtime_minutes,
        'late_night_minutes': late_night_minutes,
        'break_minutes': break_minutes,
        'lunch_deduction': lunch_deduction
    }

def calculate_late_night_minutes(start_time, end_time):
    """
    深夜時間（22:00-05:00）の計算
    
    例:
    - 21:00-23:00 → 60分
    - 22:00-02:00 → 240分
    - 04:00-06:00 → 60分
    - 09:00-18:00 → 0分
    """
    # 深夜時間帯の定義
    NIGHT_START = time(22, 0)  # 22:00
    NIGHT_END = time(5, 0)     # 05:00
    
    # 実装ロジック...
```

### 2.4 36協定チェックの累計計算
```python
class Agreement36Checker:
    """36協定準拠チェッククラス"""
    
    # 限度時間定義
    LIMITS = {
        'daily_overtime_warning': 14 * 60,      # 14時間（分）
        'monthly_overtime': 29 * 60,            # 29時間
        'monthly_holiday': 15 * 60 + 20,        # 15時間20分
        'monthly_total_max': 60 * 60,           # 60時間（改革法）
        'yearly_overtime': 176 * 60,            # 176時間
        'yearly_holiday': 184 * 60,             # 184時間
        'yearly_total_max': 398 * 60,           # 398時間（改革法）
        'yearly_holiday_count': 42              # 42回
    }
    
    def check_monthly_overtime(self, personal_code, target_month):
        """
        月間残業時間チェック
        
        戻り値: {
            'overtime_minutes': 1740,           # 29時間
            'holiday_minutes': 920,             # 15時間20分
            'total_minutes': 2660,              # 44時間20分
            'violations': [
                {'type': 'monthly_overtime', 'limit': 1740, 'actual': 1800}
            ]
        }
        """
        # データベースから当月の勤務データを取得
        records = get_monthly_records(personal_code, target_month)
        
        overtime_minutes = 0
        holiday_minutes = 0
        
        for record in records:
            if record.is_holiday:
                holiday_minutes += record.overtime_minutes
            else:
                overtime_minutes += record.overtime_minutes
        
        # 違反チェック
        violations = []
        if overtime_minutes > self.LIMITS['monthly_overtime']:
            violations.append({
                'type': 'monthly_overtime',
                'limit': self.LIMITS['monthly_overtime'],
                'actual': overtime_minutes,
                'message': f"月間時間外労働時間（休日除く）が36協定限度({self.LIMITS['monthly_overtime']//60}時間)を超過"
            })
        
        return {
            'overtime_minutes': overtime_minutes,
            'holiday_minutes': holiday_minutes,
            'total_minutes': overtime_minutes + holiday_minutes,
            'violations': violations
        }
```

### 2.5 フレックスタイム特有のチェック
```python
def check_flex_time_rules(work_record, is_flex):
    """
    フレックスタイム勤務者向けチェック
    """
    errors = []
    
    if not is_flex:
        return errors
    
    # エラー39: 勤務時間の片方のみ入力
    if bool(work_record.work_begin) != bool(work_record.work_end):
        errors.append({
            'code': '39',
            'field': ['work_begin', 'work_end'],
            'message': '勤務開始時刻と終了時刻は両方入力してください'
        })
    
    # エラー47: 時刻の論理順序
    if all([work_record.work_begin, work_record.work_end,
            work_record.break_begin1, work_record.break_end1]):
        
        # 勤務開始 < 休憩開始 < 休憩終了 < 勤務終了
        times = [
            ('勤務開始', work_record.work_begin),
            ('休憩開始', work_record.break_begin1),
            ('休憩終了', work_record.break_end1),
            ('勤務終了', work_record.work_end)
        ]
        
        for i in range(len(times) - 1):
            if time_to_minutes(times[i][1]) >= time_to_minutes(times[i+1][1]):
                errors.append({
                    'code': '47',
                    'field': ['work_begin', 'work_end', 'break_begin1', 'break_end1'],
                    'message': '勤務時間、休憩、中抜けの開始終了順序が正しくありません'
                })
                break
    
    # 法定休日チェック（週1日）
    week_records = get_week_records(work_record.personal_code, work_record.date)
    legal_holidays = [r for r in week_records if r.morning_holiday == 'A']
    
    if len(legal_holidays) == 0:
        errors.append({
            'code': 'FLEX01',
            'message': '1週間に1日以上の法定休日が必要です'
        })
    elif len(legal_holidays) > 1:
        errors.append({
            'code': 'FLEX02',
            'message': '1週間に2日以上の法定休日は設定できません'
        })
    
    return errors
```

## 3. エラー表示とユーザーフィードバック

### 3.1 エラー表示パターン
```javascript
// フィールドレベルエラー
function showFieldError(fieldElement, errorCode, errorMessage) {
    // 1. フィールドを赤枠表示
    fieldElement.classList.add('is-invalid');
    fieldElement.setAttribute('data-error-code', errorCode);
    
    // 2. エラーメッセージ表示
    const errorDiv = fieldElement.nextElementSibling;
    if (errorDiv && errorDiv.classList.contains('invalid-feedback')) {
        errorDiv.textContent = errorMessage;
        errorDiv.style.display = 'block';
    }
    
    // 3. 行にエラーフラグ設定
    const row = fieldElement.closest('tr');
    row.classList.add('has-error');
    
    // 4. エラーアイコン表示
    const errorIcon = row.querySelector('.error-indicator');
    if (errorIcon) {
        errorIcon.innerHTML = '<i class="bi bi-exclamation-triangle text-danger"></i>';
        errorIcon.title = errorMessage;
    }
}

// 月次チェックエラー
function showMonthlyErrors(errors) {
    const errorSummary = document.getElementById('errorSummary');
    errorSummary.innerHTML = '';
    
    errors.forEach(error => {
        const alertDiv = document.createElement('div');
        alertDiv.className = 'alert alert-danger alert-dismissible';
        alertDiv.innerHTML = `
            <strong>エラー ${error.code}:</strong> ${error.message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        errorSummary.appendChild(alertDiv);
    });
    
    // エラーがある日付をハイライト
    errors.forEach(error => {
        if (error.date) {
            const row = document.querySelector(`tr[data-date="${error.date}"]`);
            if (row) {
                row.classList.add('table-danger');
            }
        }
    });
}
```

### 3.2 チェック実行順序と依存関係
```
1. 基本フォーマットチェック
   └─ 時刻形式、必須入力
2. ペアチェック
   └─ 開始・終了時刻の組
3. 時系列チェック
   └─ 時刻の前後関係
4. 業務ルールチェック
   ├─ 勤務区分組み合わせ
   ├─ 休憩時間規定
   └─ 勤務体系別ルール
5. 集計チェック
   ├─ 休暇残日数
   └─ 36協定準拠
```

## 4. テストケース例

### 4.1 正常系
```yaml
test_case_1:
  description: "通常勤務（残業なし）"
  input:
    work_begin: "08:30"
    work_end: "17:30"
    break_begin1: "12:00"
    break_end1: "13:00"
  expected:
    actual_minutes: 480  # 8時間
    overtime_minutes: 0
    errors: []

test_case_2:
  description: "残業あり（深夜なし）"
  input:
    work_begin: "08:30"
    work_end: "20:00"
    break_begin1: "12:00"
    break_end1: "13:00"
  expected:
    actual_minutes: 630  # 10.5時間
    overtime_minutes: 150  # 2.5時間
    errors: []
```

### 4.2 異常系
```yaml
test_case_error_1:
  description: "時刻フォーマットエラー"
  input:
    work_begin: "8時30分"
    work_end: "17:30"
  expected:
    errors:
      - code: "FORMAT"
        field: "work_begin"
        message: "時刻を正しく入力してください"

test_case_error_2:
  description: "36協定違反"
  input:
    work_begin: "08:00"
    work_end: "23:00"  # 15時間勤務
  expected:
    errors:
      - code: "36AGR"
        message: "当日時間外労働が36協定の限度時間(14時間)を超えています"
```

## 5. 実装上の注意点

1. **パフォーマンス考慮**
   - リアルタイムチェックは軽量な処理のみ
   - 重い集計処理は保存時のみ実行
   - 月次データはキャッシュ活用

2. **エラー処理の優先度**
   - 致命的エラー（36協定違反）は最優先表示
   - フォーマットエラーは入力時即座に表示
   - 警告レベルは黄色、エラーは赤色で区別

3. **ユーザビリティ**
   - エラー修正後は即座にエラー表示をクリア
   - Tab移動でも適切にチェック実行
   - エラーがある行は視覚的に強調

4. **データ整合性**
   - 一時保存時もチェック実行（警告のみ）
   - 承認申請時は全チェック必須
   - エラーがある場合は承認申請不可