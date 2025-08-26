# データベーススキーマ情報

## 主要テーブル一覧

### セッション情報（ウェブシステム）
- **MM_Username**: 個人コード (char(5))
- **MM_staffname**: 氏名 (char(30))
- **MM_orgcode**: 組織コード (char(6))
- **MM_is_input**: 勤怠入力フラグ (char(1))
- **MM_is_deduction**: 支店控除入力担当者フラグ (char(1))
- **MM_is_charge**: 全体入力担当者フラグ (char(1))
- **MM_is_superior**: 上長フラグ (char(1))
- **MM_orgname**: 組織名称 (char(100))
- **MM_workshift**: 勤務体系 (char(1))

### stafftbl（社員テーブル）
- **personalcode**: 個人コード (char(5))
- **staffname**: 氏名
- **orgcode**: 組織コード (char(6))
- **password**: パスワード（SHA1ハッシュ）
- **is_enable**: 有効フラグ ('1': 有効, '0': 無効)
- **is_operator, is_input, is_deduction, is_charge, is_superior**: 各種権限フラグ
- **workshift**: 勤務体系

### worktbl（勤務テーブル）
- 日次勤務データを格納
- 出勤・退勤時刻、休憩時間、勤務区分など

### dutyrostertbl（勤務表テーブル）
- 月次集計データを格納
- 労働時間、残業時間、休暇日数など42項目

### その他主要テーブル
- **baseworktimetbl**: 基準労働時間テーブル
- **deductiontbl**: 支店控除テーブル
- **controltbl**: 処理制御テーブル
- **orgnametbl**: 組織名テーブル

## 勤務体系区分
- 0: 一般勤務
- 1: コミュ全日
- 2: コミュ午前
- 3: コミュ午後
- 9: フレックス勤務