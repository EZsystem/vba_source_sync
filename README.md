# VBAクラス変換 (Class Header Fixer)

## 概要
VBAのクラスモジュール（.cls）を外部からインポートする際、属性ヘッダーがないために「標準モジュール」として誤認識される問題を解消するためのツールです。
エクスポートされたテキストファイルの先頭に、VBA実行環境がクラスとして認識するために必要な属性定義を一括で付与します。

## 主な機能
- **一括属性付与**: 複数の .cls / .txt ファイルを選択し、一括でヘッダーを挿入します。
- **二重処理防止**: 既に `VERSION 1.0 CLASS` 等のヘッダーが存在するファイルは自動的にスキップします。
- **Shift-JIS維持**: VBAとの互換性を保つため、エンコーディングを `Shift-JIS (CP932)` で固定して処理します。
- **自動クラス名設定**: ファイル名を `Attribute VB_Name` に自動反映します。

## 開発環境 / 使用技術
- [cite_start]**Python**: 3.12 (64-bit) [cite: 1106]
- [cite_start]**GUI**: customtkinter (Modern UI) [cite: 1103]
- [cite_start]**Environment**: python-dotenv [cite: 1103]

## 構築手順
1. `D:\My_code\03_Tools\vba_class_fixer` にて仮想環境を作成。
   ```bash
   C:\python64\python.exe -m venv .venv

必要ライブラリのインストール。
Bash
pip install customtkinter python-dotenv




フォルダ構成
bin/: 実行スクリプト（main.py, config.py）
.venv/: プロジェクト専用仮想環境
.env: 環境設定（DEFAULT_VBA_DIR 等）
開発メモ
エンコーディング注意: VBAは内部的にShift-JISを期待するため、UTF-8への変換は行わないこと。
属性定義: Attribute VB_Exposed = False 等の標準的なクラス設定をデフォルト値として採用。
