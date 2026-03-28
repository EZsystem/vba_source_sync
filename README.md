# VBA Management Tools

VBA開発における「資産のバックアップ・同期」と「クラスモジュールのインポート修復」を効率化するためのツール群です。

## 1. VBA Source Sync System (Exporter & Sync)
OfficeファイルからVBAコードやデータベース構造を抽出し、GitHubで管理します。

### 🛠 主な機能
- **VBAエクスポート**: Excel/Accessからモジュール、クラス、フォームを自動抽出。
- **Access SQL抽出**: クエリのSQL文および主キーを含むテーブル構造（DDL）の書き出し。
- **Excel構成情報の可視化**: シート一覧、テーブル名、範囲、列定義をテキスト化。
- **GitHub自動同期**: 変更分を自動コミットしてPush。

### 💻 実行環境
- **Python**: 32-bit版（Office COM操作のため）
- **主なライブラリ**: `pywin32`, `loguru`, `GitPython`
- **場所**: `D:\My_code\11_workspaces\VBA_manager\vba_source_sync`

---

## 2. VBA Class Header Fixer (Importer Prep)
外部からインポートする際にクラスモジュール（.cls）が「標準モジュール」として誤認識される問題を解消します。

### 🛠 主な機能
- **一括属性付与**: 選択したファイルの先頭に、VBA環境に必要なクラス属性定義を挿入。
- **Shift-JIS維持**: VBAとの互換性を保つため `CP932` エンコーディングを固定。
- **自動クラス名設定**: ファイル名から `Attribute VB_Name` を自動生成。

### 💻 開発環境
- **Python**: 3.12 (64-bit)
- **GUI**: `customtkinter` (Modern UI)
- **場所**: `D:\My_code\03_Tools\vba_class_fixer`

---

## 📁 全体フォルダ構成

```text
VBA_manager/
├── vba_source_sync/ (32-bit Env)
│   ├── bin/           # main.py (実行スクリプト)
│   ├── src/           # 抽出・Git同期ロジック
│   └── workspace/     # 抽出済み資産 (vba/, sql/, excel_info/)
└── vba_class_fixer/ (64-bit Env)
    ├── bin/           # main.py, config.py
    └── .env           # DEFAULT_VBA_DIR 等の設定


🚀 共通の事前準備
Excelセキュリティ設定 (Source Sync用)
ExcelのVBAを抽出するために、以下の設定が必要です：
[ファイル] > [オプション] > [トラスト センター] > [マクロの設定] を開く。
「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」 にチェック。
仮想環境の構築
各プロジェクトのフォルダにて仮想環境を作成し、必要なライブラリをインストールしてください。
Bash
# Source Sync (32-bit)
pip install pywin32 loguru GitPython

# Class Fixer (64-bit)
pip install customtkinter python-dotenv

