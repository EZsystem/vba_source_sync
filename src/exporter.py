import os
import win32com.client
import shutil
from pathlib import Path
from loguru import logger

class VBAExporter:
    def __init__(self, workspace_dir, encoding="cp932"):
        self.workspace_dir = Path(workspace_dir)
        self.encoding = encoding

    def _get_prefix(self, app_type, comp_type):
        app_prefix = "xl_" if app_type == "excel" else "acc_"
        type_map = {1: "mod", 2: "cls", 3: "frm", 100: "cls"}
        kind_prefix = type_map.get(comp_type, "mod")
        return f"{app_prefix}{kind_prefix}_"

    def _prepare_directory(self, file_path):
        """出力先フォルダ（vba用）を準備する"""
        # workspace/ファイル名/vba という階層にする
        target_dir = self.workspace_dir / Path(file_path).stem / "vba"
        if target_dir.exists():
            shutil.rmtree(target_dir)
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
            for comp in access.VBE.ActiveVBProject.VBComponents:
                if comp.Type in [1, 2, 3]:
                    self._export_component(comp, "access", target_dir)
            access.CloseCurrentDatabase()
            return target_dir
        finally:
            access.Quit()

    def _export_component(self, comp, app_type, target_dir):
        app_prefix = "xl_" if app_type == "excel" else "acc_"
        type_map = {1: "mod", 2: "cls", 3: "frm", 100: "cls"}
        kind_name = type_map.get(comp.Type, "mod")
        full_prefix = f"{app_prefix}{kind_name}_"
        
        name = comp.Name
        
        # Accessクラスかつ "com_" を含む特殊ケースの判定
        is_access_class = (app_type == "access" and comp.Type in [2, 100])
        if is_access_class and ("com_" in name.lower()):
            if name.lower().startswith("acc_cls_"):
                name = name[len("acc_cls_"):]
            if not name.lower().startswith("com_"):
                name = self._apply_prefix_if_needed(name, full_prefix, app_prefix, kind_name)
        else:
            name = self._apply_prefix_if_needed(name, full_prefix, app_prefix, kind_name)
            
        ext = ".bas" if comp.Type == 1 else ".cls"
        if comp.Type == 3: ext = ".frm"
        
        export_path = target_dir / (name + ext)
        comp.Export(str(export_path))
        
        # 指定された文字コードへ再エンコード (デフォルトはShift-JISで出力されるため)
        try:
            with open(export_path, 'r', encoding='shift_jis', errors='replace') as f:
                content = f.read()
            with open(export_path, 'w', encoding=self.encoding) as f:
                f.write(content)
        except Exception as e:
            logger.error(f"ファイル {export_path.name} の文字コード変換に失敗しました: {e}")

    def _apply_prefix_if_needed(self, name, full_prefix, app_prefix, kind_name):
        name_l = name.lower()
        kind_prefix = kind_name + "_"
        full_base = full_prefix.rstrip("_")
        
        if name_l.startswith(full_prefix.lower()) or name_l.startswith(full_base.lower()):
            return name
        if name_l.startswith(kind_prefix.lower()) or name_l.startswith(kind_name.lower()):
            return app_prefix + name
        return full_prefix + name

class AccessSQLExtractor:
    """Accessのクエリとテーブル構造をSQLとして抽出するクラス"""
    def __init__(self, workspace_dir, encoding="utf-8"):
        self.workspace_dir = Path(workspace_dir)
        self.encoding = encoding

    def extract(self, file_path):
        file_path = os.path.abspath(file_path)
        root_dir = self.workspace_dir / Path(file_path).stem / "sql"
        
        # sqlフォルダ配下をクリーンアップ
        if root_dir.exists():
            shutil.rmtree(root_dir)
        root_dir.mkdir(parents=True, exist_ok=True)

        access = win32com.client.Dispatch("Access.Application")
        try:
            db = access.DBEngine.OpenDatabase(file_path)
            
            # 1. クエリ抽出
            query_dir = root_dir / "queries"
            query_dir.mkdir(exist_ok=True)
            for qdf in db.QueryDefs:
                if not qdf.Name.startswith("~"):
                    with open(query_dir / f"{qdf.Name}.sql", "w", encoding=self.encoding) as f:
                        f.write(qdf.SQL)

            # 2. テーブル構造抽出
            table_dir = root_dir / "tables"
            table_dir.mkdir(exist_ok=True)
            for tdf in db.TableDefs:
                if not tdf.Name.startswith("MSys") and not tdf.Name.startswith("~"):
                    fields = [f"    [{f.Name}] {self._get_type_name(f.Type)}" for f in tdf.Fields]
                    pk_fields = []
                    try:
                        for idx in tdf.Indexes:
                            if idx.Primary:
                                pk_fields = [f"[{f.Name}]" for f in idx.Fields]
                                break
                    except: pass
                    
                    sql_lines = fields.copy()
                    if pk_fields:
                        sql_lines.append(f"    CONSTRAINT PK_{tdf.Name} PRIMARY KEY ({', '.join(pk_fields)})")
                    
                    create_sql = f"CREATE TABLE [{tdf.Name}] (\n" + ",\n".join(sql_lines) + "\n);"
                    with open(table_dir / f"{tdf.Name}.sql", "w", encoding=self.encoding) as f:
                        f.write(create_sql)

            db.Close()
            return root_dir
        finally:
            access.Quit()

    def _get_type_name(self, type_id):
        type_map = {1: "YESNO", 3: "INTEGER", 4: "LONG", 5: "CURRENCY", 
                    6: "SINGLE", 7: "DOUBLE", 8: "DATETIME", 10: "TEXT", 12: "MEMO"}
        return type_map.get(type_id, "TEXT")
    
class ExcelInfoExtractor:
    """Excelのシート構成やテーブル（ListObject）情報を抽出するクラス"""
    def __init__(self, workspace_dir, encoding="utf-8"):
        self.workspace_dir = Path(workspace_dir)
        self.encoding = encoding

    def extract(self, file_path):
        file_path = os.path.abspath(file_path)
        root_dir = self.workspace_dir / Path(file_path).stem / "excel_info"
        
        if root_dir.exists():
            shutil.rmtree(root_dir)
        root_dir.mkdir(parents=True, exist_ok=True)

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(file_path, ReadOnly=True)
            info_lines = [f"Excel File: {Path(file_path).name}", "="*50, ""]

            for sheet in wb.Worksheets:
                info_lines.append(f"■ Sheet: {sheet.Name}")
                
                # テーブル（ListObject）の情報を取得
                if sheet.ListObjects.Count > 0:
                    for lo in sheet.ListObjects:
                        info_lines.append(f"  └─ Table: {lo.Name}")
                        info_lines.append(f"     Range: {lo.Range.Address}")
                        # ヘッダー項目を取得
                        headers = [str(cell.Value) for cell in lo.HeaderRowRange]
                        info_lines.append(f"     Columns: {', '.join(headers)}")
                else:
                    info_lines.append("  (No Tables)")
                info_lines.append("")

            # テキストファイルとして保存
            output_file = root_dir / "workbook_structure.txt"
            with open(output_file, "w", encoding=self.encoding) as f:
                f.write("\n".join(info_lines))

            wb.Close(False)
            return root_dir
        finally:
            excel.Quit()