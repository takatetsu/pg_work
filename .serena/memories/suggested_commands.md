# 推奨コマンド

## 開発環境
このプロジェクトはASP Classic (VBScript) + IIS + SQL Server環境で動作します。

## よく使用するコマンド

### ファイル操作
```bash
# ファイル一覧表示
ls -la

# ディレクトリ構造表示
find . -type d -name ".*" -prune -o -type d -print | sort

# ASPファイル検索
find . -name "*.asp" -type f

# インクルードファイル検索
find inc/ -name "*.asp" -type f
```

### テキスト検索
```bash
# 特定の文字列を含むファイル検索
grep -r "検索文字列" . --include="*.asp"

# 関数定義検索
grep -r "Function\|Sub" . --include="*.asp"

# データベーステーブル参照検索
grep -r "stafftbl\|worktbl\|dutyrostertbl" . --include="*.asp"
```

### Git操作
```bash
# 変更状況確認
git status

# 変更差分確認
git diff

# コミット履歴確認
git log --oneline -10

# 特定ファイルの変更履歴
git log --oneline -- "ファイル名"
```

### Claude Code & Serena操作
```bash
# Serenaメモリ一覧確認
ls .serena/memories/

# プロジェクト設定確認
cat .serena/project.yml

# Claude設定確認
cat .claude/settings.local.json
```

## 現在のプロジェクト状況（2025年8月）
- 多数のファイルが修正中（Git status参照）
- バックアップファイルとアーカイブ仕様書を削除済み
- プロジェクト構成をクリーンアップ完了

## 注意事項
- このプロジェクトは従来のASP Classicで開発されており、Node.js、npm、モダンなビルドツールは使用していません
- テストフレームワークやリンター、フォーマッターなどのモダンな開発ツールは設定されていません
- 開発は主にテキストエディタとIISでの動作確認で行います
- データベースはSQL Server Management Studioなどで管理します