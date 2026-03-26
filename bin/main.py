import os
import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from loguru import logger

# パス設定
sys.path.append(str(Path(__file__).parent.parent))
from src.exporter import VBAExporter, AccessSQLExtractor
from src.git_handler import GitManager

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_types = [
        ("Access Project Summary", "RN_ProjectSummary.accdb"),
        ("Access Database", "*.accdb;*.mdb"),
        ("Excel Macro-Enabled Workbook", "*.xlsm"),
        ("All Files", "*.*")
    ]
    return filedialog.askopenfilename(title="対象ファイルを選択", filetypes=file_types)

def main():
    project_root = Path(__file__).parent.parent
    workspace_dir = project_root / "workspace"
    
    target_file = select_file()
    if not target_file: return

    ext = Path(target_file).suffix.lower()
    is_access = ext in ['.accdb', '.mdb']

    try:
        # 1. VBAエクスポート
        exporter = VBAExporter(workspace_dir)
        if is_access:
            vba_path = exporter.export_access(target_file)
        elif ext == '.xlsm':
            vba_path = exporter.export_excel(target_file)
        
        logger.success(f"VBAエクスポート完了: {vba_path}")

        # 2. Accessの場合のみSQL抽出を実行
        if is_access:
            logger.info("SQL資産の抽出を開始します...")
            sql_extractor = AccessSQLExtractor(workspace_dir)
            sql_path = sql_extractor.extract(target_file)
            logger.success(f"SQL抽出完了: {sql_path}")

        # 3. GitHub同期
        git_mgr = GitManager(project_root)
        file_name = Path(target_file).name
        git_mgr.sync_to_github(f"Auto-sync VBA & SQL: {file_name}")

        print(f"\n【完了】全資産のバックアップとPushに成功しました。")

    except Exception as e:
        logger.exception(f"エラー: {e}")

if __name__ == "__main__":
    main()