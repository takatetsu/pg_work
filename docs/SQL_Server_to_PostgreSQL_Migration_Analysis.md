# SQL Server から PostgreSQL への移行分析（更新版）

## 概要
このドキュメントは、現在 SQL Server を使用している勤怠管理システムを PostgreSQL に移行するため、すべてのファイルで SQL を使用している箇所を洗い出し、修正が必要な箇所と変更内容をまとめたものです。

**【重要】** 新しいテーブル定義ファイル（CreateTable_SQLServer.sql と CreateTable_PostgreSQL.sql）を確認した結果、追加のテーブルと改良された構造が判明しました。

## 対象ファイル一覧
以下のファイルでSQL文が使用されています：

### メインページ
- `sample.asp`
- `checklist.asp`
- `inputall.asp`
- `check_holidaywork.asp`
- `check_holiday.asp`
- `index.asp`
- `inputdeduction.asp`
- `changeworktime.asp`
- `xls_inputall.asp`
- `check_holiday_charge.asp`
- `inputwork.asp`
- `check_overtime.asp`
- `changepassword.asp`
- `check_holidaywork_charge.asp`
- `workstatus.asp`
- `timecard.asp`
- `check_overtime_charge.asp`

### インクルードファイル
- `inc/view_init.asp`
- `inc/upsert_worktbl.asp`
- `inc/RestrictAccess.asp`
- `inc/select_stafftbl.asp`
- `inc/insert_timetbl.asp`
- `inc/properTimeCheck.asp`
- `inc/upsert_dutyrostertbl.asp`
- `inc/inputCommon1.asp`
- `inc/update_worktbl_is_approval.asp`
- `inc/view_proc.asp`

### 設定ファイル
- `Connections/workdbms.asp`
- `Connections/workdbms_テスト環境.asp`
- `Connections/workdbms_本番DB.asp`

### データベース定義
- `docs/CreateTable_SQLServer.sql` (新版)
- `docs/CreateTable_PostgreSQL.sql` (新版)

## テーブル構成の変更点

### 追加テーブル（新しく発見）
現在のASPコードには存在しないが、新テーブル定義に含まれるテーブル：

1. **`orgnametbl`** - 組織名マスタテーブル
   - 今後のコード修正時に使用される可能性あり

2. **`pctimemergetbl`** - PC電源時刻統合用テーブル 
   - バッチ処理等で使用される可能性あり

3. **`timecardtbl`** - タイムカードデータテーブル
   - `timecard.asp`で参照されている可能性あり

## SQL Server特有の機能と PostgreSQL への対応

### 1. データ型の変更（更新版）

#### 基本データ型変換
| SQL Server | PostgreSQL | 説明 | 影響ファイル |
|---|---|---|---|
| `int identity` | `SERIAL PRIMARY KEY` | 自動増分主キー | 全テーブル定義 |
| `timestamp` | `TIMESTAMP WITH TIME ZONE` | タイムスタンプ（タイムゾーン付き） | 全テーブル |
| `float(53)` | `DOUBLE PRECISION` | 倍精度浮動小数点数 | dutyrostertbl, remainvacationtbl |
| `char(n)` | `CHAR(n)` | 固定長文字列（維持） | 全テーブル |
| `[dbo].[テーブル名]` | `テーブル名` | スキーマプレフィックス削除 | 全ASPファイル |

### 2. 関数の変更

#### NULL処理関数
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `check_holiday.asp:344`, `inc/view_init.asp:356` | `IsNULL(r.remainvacation, 0)` | `COALESCE(r.remainvacation, 0)` |

#### 暗号化関数
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `index.asp:47`, `changepassword.asp:49` | `hashbytes('sha1', ?)` | `digest(?, 'sha1')` |
**注意**: PostgreSQLでは`CREATE EXTENSION pgcrypto;`の実行が必要

#### データ型変換関数
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `inc/view_init.asp:296` | `CONVERT(int,w1.updatetime)` | `CAST(w1.updatetime AS integer)` |
| `inc/view_init.asp:306` | `CONVERT(NVARCHAR, DATEADD(...), 112)` | `to_char(... + INTERVAL '...', 'YYYYMMDD')` |

#### ウィンドウ関数（対応不要）
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `check_holiday.asp:347`, `inc/view_init.asp:359` | `ROW_NUMBER() OVER (ORDER BY ymb DESC)` | 変更不要（PostgreSQLでも同じ構文） |

#### 日付フォーマット関数
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `inc/view_init.asp:310`, `inputwork.asp:122` | `FORMAT(date, 'yyyyMMdd')` | `to_char(date, 'YYYYMMDD')` |

#### 日付部分抽出関数
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `inc/view_init.asp:310`, `inputwork.asp:122` | `DATEPART(WEEKDAY, date)` | `extract(DOW FROM date)` |
**注意**: PostgreSQLでは日曜=0、SQL Serverでは日曜=1

#### 日付演算関数
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `inc/view_init.asp:306` | `DATEADD(day, -1, date)` | `date - INTERVAL '1 day'` |
| `inputwork.asp:122` | `DATEADD(day, -1*DATEPART(WEEKDAY, date), date)` | `date - (extract(DOW FROM date) || ' days')::interval` |

### 3. INSERT文の修正

#### DEFAULT値の使用
| 箇所 | SQL Server | PostgreSQL |
|---|---|---|
| `inputall.asp:241` | `INSERT INTO dbo.dutyrostertbl VALUES(DEFAULT, ...)` | `INSERT INTO dutyrostertbl VALUES(DEFAULT, ...)` |
| `inc/upsert_worktbl.asp:307` | `INSERT INTO dbo.worktbl VALUES(DEFAULT, ...)` | `INSERT INTO worktbl VALUES(DEFAULT, ...)` |
| `inc/insert_timetbl.asp:23` | `INSERT INTO dbo.timetbl VALUES(DEFAULT, ...)` | `INSERT INTO timetbl VALUES(DEFAULT, ...)` |
| `inc/upsert_dutyrostertbl.asp:461` | `INSERT INTO dbo.dutyrostertbl VALUES(DEFAULT, ...)` | `INSERT INTO dutyrostertbl VALUES(DEFAULT, ...)` |

### 4. データベース接続の変更

#### 接続文字列の修正（3ファイル）
| ファイル | 現在（SQL Server） | 変更後（PostgreSQL） |
|---|---|---|
| `Connections/workdbms.asp` | `dsn=test_work;uid=sg;pwd=kouka08;` | PostgreSQL ODBC接続文字列に変更 |
| `Connections/workdbms_テスト環境.asp` | `dsn=test_work;uid=sg;pwd=kouka08;` | PostgreSQL ODBC接続文字列に変更 |
| `Connections/workdbms_本番DB.asp` | (同様のDSN接続) | PostgreSQL ODBC接続文字列に変更 |

**推奨接続文字列例:**
```
Driver={PostgreSQL ODBC Driver};Server=localhost;Port=5432;Database=workdb;Uid=sg;Pwd=kouka08;
```

### 5. VBScript日付関数（ASPサーバーサイド処理）

以下のVBScript関数は変更不要（サーバーサイド処理のため）：
- `DateAdd()` - VBScriptの日付関数
- `Year()`, `Month()`, `Day()` - VBScriptの日付関数

## 修正優先度（更新版）

### 【最高】新テーブル定義の適用
1. **PostgreSQL用テーブル定義の適用** - `docs/CreateTable_PostgreSQL.sql`を使用
2. **pgcrypto拡張の有効化** - `CREATE EXTENSION IF NOT EXISTS pgcrypto;`
3. **インデックスとコメントの適用** - パフォーマンス向上のため

### 【高】必須修正項目
1. **データベース接続文字列の修正**（Connections/内3ファイル）
2. **IsNULL関数の修正**（4箇所）
3. **hashbytes関数の修正**（2箇所）
4. **DEFAULT値を使用するINSERT文の修正**（4箇所）

### 【中】機能に影響する修正項目
1. **CONVERT関数の修正**（3箇所）
2. **FORMAT関数の修正**（2箇所）
3. **DATEPART関数の修正**（5箇所）
4. **timestamp型関連の修正** - タイムゾーン対応

### 【低】スキーマ名の修正
1. **dbo.プレフィックスの削除**（多数箇所）
2. **未使用テーブルへの対応** - orgnametbl, pctimemergetbl, timecardtbl

## 注意事項

### 1. pgcrypto拡張の有効化
暗号化関数を使用するため、PostgreSQLで以下のコマンドを実行する必要があります：
```sql
CREATE EXTENSION pgcrypto;
```

### 2. 日付の週番号について
- SQL Server: 日曜日 = 1
- PostgreSQL: 日曜日 = 0

週の計算ロジックを含む箇所は注意深く検証が必要です。

### 3. データ型の精度
`float(53)` → `double precision` への変更により、数値の精度が変わる可能性があります。

### 4. 文字エンコーディング
SQL Server から PostgreSQL への移行時は、文字エンコーディング（UTF-8）の設定を確認してください。

## 推奨作業手順（更新版）

### フェーズ1: 環境準備
1. **PostgreSQL環境の準備**
2. **pgcrypto拡張の有効化** - `CREATE EXTENSION IF NOT EXISTS pgcrypto;`
3. **新テーブル定義の適用** - `docs/CreateTable_PostgreSQL.sql`を実行

### フェーズ2: 接続設定とコア機能修正
1. **データベース接続文字列の修正**（Connections/内3ファイル）
2. **IsNULL関数の修正**（4箇所）
3. **hashbytes関数の修正**（2箇所）
4. **基本動作確認テスト**

### フェーズ3: 高度機能の修正
1. **CONVERT関数の修正**（3箇所）
2. **FORMAT関数の修正**（2箇所）
3. **DATEPART関数の修正**（5箇所）
4. **日付・時刻関連機能のテスト**

### フェーズ4: 最適化と仕上げ
1. **dbo.プレフィックスの削除**（多数箇所）
2. **パフォーマンステスト**
3. **追加テーブルの活用検討** (orgnametbl, pctimemergetbl, timecardtbl)
4. **総合テスト**

## 総修正箇所数（更新版）

- **テーブル定義**: 16テーブル（3テーブル追加）
- **接続文字列**: 3ファイル  
- **関数修正**: 約50箇所
- **スキーマプレフィックス**: 多数箇所
- **新機能**: インデックス追加、コメント追加、タイムゾーン対応

## 追加考慮事項

### 新テーブルの活用
1. **`timecardtbl`**: `timecard.asp`での活用を検討
2. **`orgnametbl`**: 組織名表示の改善に利用可能
3. **`pctimemergetbl`**: バッチ処理でのパフォーマンス向上

### パフォーマンス向上
- PostgreSQL版では適切なインデックスが追加済み
- `TIMESTAMP WITH TIME ZONE`によるタイムゾーン対応
- テーブルコメントによる保守性向上

修正作業は段階的に行い、各フェーズで十分なテストを実施することを推奨します。