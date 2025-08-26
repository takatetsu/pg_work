# 勤務表管理システム - プロジェクト概要

## プロジェクトの目的
従業員の勤務時間管理、残業時間管理、休暇管理を一元的に行うWebベースの勤務表管理システム。日本の労働基準法および36協定に準拠した労働時間管理を実現。

## 技術スタック
- **プログラミング言語**: ASP Classic (VBScript)
- **Webサーバー**: IIS (Internet Information Services) 
- **データベース**: Microsoft SQL Server (ODBC接続)
- **フロントエンド**: HTML, CSS, JavaScript (jQuery使用)
- **文字エンコーディング**: UTF-8 (CODEPAGE 65001)
- **認証方式**: SHA1ハッシュ + セッション管理

## プロジェクト構造（2025年8月現在）
```
/
├── .claude/               # Claude Code設定
├── .serena/               # Serenaメモリ・キャッシュ
├── Connections/           # データベース接続設定
│   ├── workdbms.asp
│   ├── workdbms_テスト環境.asp
│   └── workdbms_本番DB.asp
├── css/                   # CSSファイル
│   ├── default.css
│   ├── style.css
│   └── superTables_compressed.css
├── docs/                  # 仕様書・ドキュメント
│   ├── テーブル情報.MD
│   └── CreateTable.SQL
├── inc/                   # インクルードファイル（共通処理）
│   ├── RestrictAccess.asp
│   ├── footer.source
│   ├── header.source
│   ├── inputCommon1.asp
│   ├── inputworkCheck.asp
│   ├── insert_timetbl.asp
│   ├── properTimeCheck.asp
│   ├── select_stafftbl.asp
│   ├── update_worktbl_is_approval.asp
│   ├── upsert_dutyrostertbl.asp
│   ├── upsert_worktbl.asp
│   ├── util.asp
│   ├── view_init.asp
│   └── view_proc.asp
├── js/                    # JavaScriptファイル
│   ├── redips-drag-min.js
│   ├── script.js
│   └── superTables_compressed.js
├── *.asp                  # ASPページファイル（メイン機能）
├── web.config            # IIS設定ファイル
├── .mcp.json             # MCP設定
└── 36協定書.pdf, フレックスタイム制に関する協定.pdf
```

## 主要機能
1. **認証・権限管理** - ユーザー認証、役割ベースアクセス制御
2. **勤務時間管理** - 日次勤務時間入力、自動計算
3. **承認ワークフロー** - 上長による勤務実績承認
4. **休暇管理** - 有給休暇、代休、特別休暇管理
5. **コンプライアンス機能** - 36協定違反チェック
6. **データ入出力** - CSV/Excelインポート・エクスポート

## 最近の変更点（2025年8月）
- 不要なバックアップファイルとアーカイブされた仕様書を削除
- プロジェクト構成の整理とクリーンアップを実施
- 多数のファイルに対してコードの改善・修正を実施