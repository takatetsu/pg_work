# 休出日数集計に関する調査報告書

## 調査日: 2025-06-19

## 1. 調査対象
- 休出回数の累積集計
- 当月の休出日数の集計
- dutyrostertbl.holidayshiftsへの値設定
- フレックス勤務者のholidayshiftsが0になる問題

## 2. 休出回数の累積集計

### 2.1 集計条件
出勤区分が以下のいずれかの場合を休出としてカウント：
- `2`：休日出勤
- `3`：祝日出勤  
- `6`：休出遅刻

### 2.2 集計処理

#### A. 当月の休出回数（view_proc.asp：382-391行目）
```asp
If v_morningwork = "2" Or _
   v_morningwork = "3" Or _
   v_morningwork = "6" Or _
   v_afternoonwork = "2" Or _
   v_afternoonwork = "3" Or _
   v_afternoonwork = "6" Then
    sumHolidayWork = sumHolidayWork + 1
End If
```
- 日単位でカウント（午前・午後どちらかが該当すれば1回）

#### B. 年度累計の休出回数（view_init.asp：242-263行目）
```sql
SELECT COUNT(*) AS holidaywork FROM dbo.worktbl 
WHERE personalcode = ? AND workingdate >= ? AND workingdate < ? AND 
(morningwork IN ('2', '3', '6') OR afternoonwork IN ('2', '3', '6'))
```
- 対象期間：当年度開始（businessYear）から当月の前月末まで

#### C. 当月の休出回数（view_init.asp：264-284行目）
- 対象期間：当月のみ
- データベースから直接集計

### 2.3 画面表示（inputwork.asp：1158-1167行目）
- **累積休出回数** = `yearlyHolidaywork`（年度累計） + `monthlyHolidaywork`（当月分）
- **警告表示**：
  - 42回以上：赤色（abnormality）
  - 35回以上：黄色（warning）

## 3. 当月の休出日数の集計

### 3.1 集計方法の違い

#### A. 通常勤務者（workshift ≠ 9）
1. **時間の集計**（view_proc.asp）：
   - `sumHolidayshifttime`：休日出勤時間（分）
   - `sumHolidayshiftlate`：休出深夜時間（分）

2. **日数への変換**（inputwork.asp：895行目）：
   ```asp
   temp_holidayshifts = mm2FloatDay(sumHolidayshifttime + sumHolidayshiftlate)
   ```

#### B. フレックス勤務者（workshift = 9）
1. **時間の集計**（view_proc.asp：82-91行目）：
   - 出勤区分が2（休日出勤）または6（休出遅刻）の場合
   - `sumFlex_holidayshift`に実際の勤務時間（workmin）を加算

2. **日数への変換**（inputwork.asp：893行目）：
   ```asp
   temp_holidayshifts = mm2FloatDay(sumFlex_holidayshift)
   ```

### 3.2 日数変換の仕組み（mm2FloatDay関数）
- 460分（7時間40分）を1日として計算
- 小数点第1位まで表示（第2位以下は切り上げ）

### 3.3 警告表示
- 2日以上：赤色（abnormality）
- 1.5日以上：黄色（warning）

## 4. dutyrostertbl.holidayshiftsへの値設定

### 4.1 時間の集計（upsert_dutyrostertbl.asp：281-322行目）

**通常勤務者**：
- `total_holidayshifttime`：worktblの`holidayshift`フィールドから集計
- `total_holidayshiftlate`：worktblの`holidayshiftlate`フィールドから集計

**フレックス勤務者**：
- `total_holidayshifttime`：worktblの`workmin`から集計（法定休日を除く）
- `total_holidayshiftlate`：0

### 4.2 日数への変換と格納
```asp
temp_holidayshifts = mm2FloatDay(total_holidayshifttime + total_holidayshiftlate)
```

## 5. フレックス勤務者のholidayshiftsが0になる問題

### 5.1 原因
フレックス勤務者が**法定休日（日曜日）**に休出した場合：

1. 集計条件（282-290行目）により、法定休日は`total_holidayshifttime`への加算から除外
   ```asp
   If ((Rs_worktbl_sum.Fields.Item("morningholiday").value <> "A" And _
        Rs_worktbl_sum.Fields.Item("afternoonholiday").value <> "A") And ...
   ```

2. 法定休日の勤務時間は`legalholiday_extra_min`に別途集計（421-435行目）

3. `temp_holidayshifts`の計算時、`total_holidayshifttime`は0のまま

4. 結果として`holidayshifts`は0になる

### 5.2 設計意図
- 法定休日の休出は給与計算システムの都合により別管理
- **時間代休フィールド**（workholidaytime）に法定休日割増時間を格納
- **休日出勤時間フィールド**（holidayshifttime）に合計時間を格納

### 5.3 データの格納先
フレックス勤務者の場合：
- `holidayshifts`：法定休日以外の休出日数
- `workholidaytime`：法定休日割増時間（時間単位）
- `holidayshifttime`：法定休日割増時間 + 通常休出時間

## 6. まとめ

1. **休出回数**：出勤区分2,3,6を日単位でカウント（全勤務形態共通）

2. **休出日数**：
   - 通常勤務者：休出時間と休出深夜時間から計算
   - フレックス勤務者：実勤務時間から計算（法定休日除く）

3. **フレックス勤務者の法定休日休出**：
   - holidayshiftsには反映されない（仕様）
   - 別フィールドで管理される

## 7. 関連ファイル
- `/inc/view_proc.asp`：画面表示用の集計処理
- `/inc/view_init.asp`：データベースからの集計読込
- `/inc/upsert_dutyrostertbl.asp`：dutyrostertblへの更新処理
- `/inputwork.asp`：画面表示
- `/inc/util.asp`：時間・日数変換関数