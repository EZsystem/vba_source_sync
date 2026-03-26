import os
import win32com.client
import shutil
from pathlib import Path
from loguru import logger

class VBAExporter:
    def __init__(self, workspace_dir):
        self.workspace_dir = Path(workspace_dir)

    def _get_prefix(self, app_type, comp_type):
        """ルールブックに基づいたプレフィックスを決定する"""
        app_prefix = "xl_" if app_type == "excel" else "acc_"
        # 1:標準(mod), 2:クラス(cls), 3:フォーム(frm), 100:Document(cls)
        type_map = {1: "mod", 2: "cls", 3: "frm", 100: "cls"}
        kind_prefix = type_map.get(comp_type, "mod")
        return f"{app_prefix}{kind_prefix}_"

    def _prepare_directory(self, file_path):
        """出力先フォルダをクリーンアップして準備する"""
        target_dir = self.workspace_dir / Path(file_path).stem
        if target_dir.exists():
            logger.info(f"既存のフォルダを削除してクリーンアップします: {target_dir}")
            shutil.rmtree(target_dir)  # フォルダごと削除（Gitには影響なし）
        target_dir.mkdir(parents=True, exist_ok=True)
        return target_dir

    def export_excel(self, file_path):
        file_path = os.path.abspath(file_path)
        target_dir = self._prepare_directory(file_path)
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(file_path)
            for comp in wb.VBProject.VBComponents:
                if comp.Type in [1, 2, 3, 100]:
                    self._export_component(comp, "excel", target_dir)
            wb.Close(False)
            return target_dir
        finally:
            excel.Quit()

    def export_access(self, file_path):
        file_path = os.path.abspath(file_path)
        target_dir = self._prepare_directory(file_path)
        
        access = win32com.client.Dispatch("Access.Application")
        try:
            access.OpenCurrentDatabase(file_path)
            # AccessもVBProject経由で走査することで型を正確に取得
            for comp in access.VBE.ActiveVBProject.VBComponents:
                if comp.Type in [1, 2, 3]:
                    self._export_component(comp, "access", target_dir)
            access.CloseCurrentDatabase()
            return target_dir
        finally:
            access.Quit()

    def _export_component(self, comp, app_type, target_dir):
        """個別コンポーネントのエクスポート実行ロジック"""
        # 接頭辞の構成要素を分解して取得
        app_prefix = "xl_" if app_type == "excel" else "acc_"
        type_map = {1: "mod", 2: "cls", 3: "frm", 100: "cls"}
        kind_name = type_map.get(comp.Type, "mod")
        full_prefix = f"{app_prefix}{kind_name}_"
        
        name = comp.Name
        
        # 1. Accessクラスかつ "com_" を含む特殊ケース
        is_access_class = (app_type == "access" and comp.Type in [2, 100])
        if is_access_class and ("com_" in name.lower()):
            if name.lower().startswith("acc_cls_"):
                name = name[len("acc_cls_"):]
            if name.lower().startswith("com_"):
                pass 
            else:
                name = self._apply_prefix_if_needed(name, full_prefix, app_prefix, kind_name)
        # 2. 通常のケース
        else:
            name = self._apply_prefix_if_needed(name, full_prefix, app_prefix, kind_name)
            
        ext = ".bas" if comp.Type == 1 else ".cls"
        if comp.Type == 3: ext = ".frm"
        
        export_path = target_dir / (name + ext)
        comp.Export(str(export_path))

    def _apply_prefix_if_needed(self, name, full_prefix, app_prefix, kind_name):
        """接頭辞の重複を構造的にチェックして適用する"""
        name_l = name.lower()
        kind_prefix = kind_name + "_"  # 例: "mod_"
        
        # パターンA: すでにフル接頭辞がある (例: "acc_mod_..." または "acc_mod...")
        full_base = full_prefix.rstrip("_")
        if name_l.startswith(full_prefix.lower()) or name_l.startswith(full_base.lower()):
            return name
            
        # パターンB: 種類接頭辞だけがある (例: "mod_..." または "mod...")
        # この場合はアプリ名(acc_)だけを先頭に付与する
        if name_l.startswith(kind_prefix.lower()) or name_l.startswith(kind_name.lower()):
            return app_prefix + name
            
        # パターンC: 何も付いていない場合はフルで付与
        return full_prefix + name