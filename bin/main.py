import os
import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from loguru import logger

# 自作モジュールのインポート用パス設定（インポート前に実行）
sys.path.append(str(Path(__file__).parent.parent))
from src.exporter import VBAExporter
from src.git_handler import GitManager

def select_file():
    """ファイル選択ダイアログを表示する"""
    root = tk.Tk()
    root.withdraw()  # メインウィンドウは表示しない
    
    # 最初はAccessファイルを優先して表示
    file_types = [
        ("Access Database", "*.accdb;*.mdb"),
        ("Excel Macro-Enabled Workbook", "*.xlsm"),
        ("All Files", "*.*")
    ]
    
    selected_path = filedialog.askopenfilename(
        title="バックアップ対象のファイルを選択してください",
        filetypes=file_types
    )
    return selected_path

def main():
    # 1. ログの設定
    project_root = Path(__file__).parent.parent
    log_dir = project_root / "logs"
    log_dir.mkdir(exist_ok=True)
    logger.add(log_dir / "exec_{time:YYYYMMDD}.log", rotation="1 day", level="INFO")
    
    logger.info("=== VBA Export & Sync Process Started ===")

    # 2. ファイル選択
    target_file = select_file()
    if not target_file:
        logger.warning("ファイルが選択されませんでした。処理を中断します。")
        return

    logger.info(f"対象ファイル: {target_file}")

    # 3. エクスポート準備
    workspace_dir = project_root / "workspace"
    exporter = VBAExporter(workspace_dir)
    
    ext = Path(target_file).suffix.lower()
    
    try:
        # 4. エクスポート実行
        if ext in ['.accdb', '.mdb']:
            logger.info("Accessとして処理を開始します...")
            output_path = exporter.export_access(target_file)
        elif ext == '.xlsm':
            logger.info("Excelとして処理を開始します...")
            output_path = exporter.export_excel(target_file)
        else:
            logger.error(f"未対応の拡張子です: {ext}")
            return

        logger.success(f"エクスポート完了: {output_path}")

        # 5. GitHub同期の実行
        logger.info("GitHub同期を開始します...")
        git_mgr = GitManager(project_root)
        
        # コミットメッセージに対象ファイル名を含める
        file_name = Path(target_file).name
        git_mgr.sync_to_github(f"Auto-sync VBA: {file_name}")

        print(f"\n【完了】エクスポートとGitHubへのPushが成功しました。")
        print(f"場所: {output_path}")

    except Exception as e:
        logger.exception(f"実行中にエラーが発生しました: {e}")
        print(f"\n【エラー】処理に失敗しました。詳細はログを確認してください。")

if __name__ == "__main__":
    main()
