import os
import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from loguru import logger
import gc  # ガベージコレクション（メモリ解放）用に追加

# 自作モジュールのインポート用パス設定
sys.path.append(str(Path(__file__).parent.parent))
from src.exporter import VBAExporter, AccessSQLExtractor
from src.git_handler import GitManager

def select_file():
    """ファイル選択ダイアログを表示する（Excel/Access両対応版）"""
    root = tk.Tk()
    root.withdraw()
    
    file_types = [
        ("Office Macro Files", "*.accdb;*.mdb;*.xlsm"),
        ("Access Database", "*.accdb;*.mdb"),
        ("Excel Macro-Enabled Workbook", "*.xlsm"),
        ("All Files", "*.*")
    ]
    
    return filedialog.askopenfilename(
        title="バックアップ対象のファイルを選択してください",
        filetypes=file_types
    )

def main():
    project_root = Path(__file__).parent.parent
    workspace_dir = project_root / "workspace"
    
    # 2. ファイル選択
    target_file = select_file()
    if not target_file:
        logger.warning("ファイルが選択されませんでした。")
        return

    ext = Path(target_file).suffix.lower()
    is_access = ext in ['.accdb', '.mdb']

    try:
        # 4. エクスポート・抽出実行
        exporter = VBAExporter(workspace_dir)
        if is_access:
            logger.info(f"Accessとして処理を開始します: {target_file}")
            output_path = exporter.export_access(target_file)
            
            logger.info("SQL資産の抽出を開始します...")
            sql_extractor = AccessSQLExtractor(workspace_dir)
            sql_extractor.extract(target_file)
        else:
            logger.info(f"Excelとして処理を開始します: {target_file}")
            output_path = exporter.export_excel(target_file)

        logger.success(f"エクスポート完了: {target_file}")

        # 5. GitHub同期
        logger.info("GitHub同期を開始します...")
        git_mgr = GitManager(project_root)
        file_name = Path(target_file).name
        git_mgr.sync_to_github(f"Auto-sync: {file_name}")

        print(f"\n【完了】バックアップとGitHubへのPushが成功しました。")

    except Exception as e:
        logger.exception(f"実行中にエラーが発生しました: {e}")
        print(f"\n【エラー】処理に失敗しました。")

    finally:
        # エラーメッセージ（WinError 6）対策：
        # 明示的にオブジェクトを解放し、ガベージコレクションを実行することで
        # 終了時のハンドル無効エラーを抑制しやすくします。
        if 'git_mgr' in locals(): del git_mgr
        gc.collect()

if __name__ == "__main__":
    main()