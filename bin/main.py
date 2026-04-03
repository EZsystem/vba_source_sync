import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
from loguru import logger
import gc  # ガベージコレクション（メモリ解放）用に追加

# 自作モジュールのインポート用パス設定
sys.path.append(str(Path(__file__).parent.parent))
from src.exporter import VBAExporter, AccessSQLExtractor, ExcelInfoExtractor, AccessDataSampler
from src.git_handler import GitManager
from src.history_mgr import HistoryManager

def select_file_and_encoding():
    """ファイル選択ダイアログと文字コード選択画面を表示する（Excel/Access両対応版）"""
    root = tk.Tk()
    root.title("バックアップ対象と文字コードの選択")
    root.geometry("450x150")
    
    history_mgr = HistoryManager()
    selected_file = tk.StringVar()
    selected_encoding = tk.StringVar(value="デフォルト")
    
    # 履歴を読み込む
    history_list = history_mgr.get_history()
    if history_list:
        selected_file.set(history_list[0])
    
    result = {"file": "", "encoding": ""}
    
    def browse_file():
        file_types = [
            ("Office Macro Files", "*.accdb;*.mdb;*.xlsm"),
            ("Access Database", "*.accdb;*.mdb"),
            ("Excel Macro-Enabled Workbook", "*.xlsm"),
            ("All Files", "*.*")
        ]
        filename = filedialog.askopenfilename(
            title="バックアップ対象のファイルを選択してください",
            filetypes=file_types
        )
        if filename:
            selected_file.set(filename)
            # 参照で選んだ場合も、コンボボックスの値を更新し、履歴追加の準備をする
            file_cb['values'] = history_mgr.get_history()
            
    def execute():
        if not selected_file.get():
            messagebox.showwarning("ファイル未選択", "バックアップ対象のファイルを選択してください。")
            return
        result["file"] = selected_file.get()
        result["encoding"] = selected_encoding.get()
        # 履歴を保存
        history_mgr.save_history(result["file"])
        root.destroy()
        
    def cancel():
        root.destroy()

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)
    
    ttk.Label(frame, text="対象ファイル:").grid(row=0, column=0, sticky=tk.W, pady=5)
    file_cb = ttk.Combobox(frame, textvariable=selected_file, width=40)
    file_cb['values'] = history_list
    file_cb.grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(frame, text="参照...", command=browse_file).grid(row=0, column=2, pady=5)
    
    ttk.Label(frame, text="文字コード:").grid(row=1, column=0, sticky=tk.W, pady=5)
    encoding_cb = ttk.Combobox(frame, textvariable=selected_encoding, state="readonly", width=15)
    encoding_cb['values'] = ("デフォルト(utf-8)", "utf-8-sig", "shift-jis", "cp932")
    encoding_cb.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
    
    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=2, column=0, columnspan=3, pady=15)
    ttk.Button(btn_frame, text="実行", command=execute).pack(side=tk.LEFT, padx=10)
    ttk.Button(btn_frame, text="キャンセル", command=cancel).pack(side=tk.LEFT, padx=10)
    
    root.eval('tk::PlaceWindow . center')
    root.mainloop()
    
    return result["file"], result["encoding"]

def main():
    project_root = Path(__file__).parent.parent
    workspace_dir = project_root / "workspace"
    
    # 2. ファイルと文字コード選択
    target_file, encoding_choice = select_file_and_encoding()
    if not target_file:
        logger.warning("ファイルが選択されませんでした（またはキャンセルされました）。")
        return

    # 文字コードの決定ロジック
    vba_enc = "utf-8"       # デフォルト (VBA同期マクロ側で変換するため、保存はUTF-8でOK)
    other_enc = "utf-8"     # デフォルト
    
    if encoding_choice == "utf-8-sig":
        vba_enc = "utf-8-sig"
        other_enc = "utf-8-sig"
    elif encoding_choice == "shift-jis":
        vba_enc = "shift_jis"
        other_enc = "shift_jis"
    elif encoding_choice == "cp932":
        vba_enc = "cp932"
        other_enc = "cp932"
    elif encoding_choice == "デフォルト(utf-8)":
        vba_enc = "utf-8"
        other_enc = "utf-8"
    
    logger.info(f"選択文字コード: {encoding_choice} (VBA={vba_enc}, 他={other_enc})")

    ext = Path(target_file).suffix.lower()
    is_access = ext in ['.accdb', '.mdb']

    try:
        # 4. エクスポート・抽出実行
        exporter = VBAExporter(workspace_dir, encoding=vba_enc)
        if is_access:
            logger.info(f"Accessとして処理を開始します: {target_file}")
            output_path = exporter.export_access(target_file)
            
            logger.info("SQL資産の抽出を開始します...")
            sql_extractor = AccessSQLExtractor(workspace_dir, encoding=other_enc)
            sql_extractor.extract(target_file)
            
            logger.info("レコード数とサンプルデータの抽出を開始します...")
            sampler = AccessDataSampler(workspace_dir, encoding=other_enc)
            sampler.extract_samples(target_file)
        else:
            logger.info(f"Excelとして処理を開始します: {target_file}")
            output_path = exporter.export_excel(target_file)
        
            # --- ここを追加 ---
            logger.info("Excel構成情報の抽出を開始します...")
            excel_extractor = ExcelInfoExtractor(workspace_dir, encoding=other_enc)
            excel_extractor.extract(target_file)
            # -----------------

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