import os
import win32com.client
from pathlib import Path

def extract_access_assets(db_path, output_base):
    db_path = os.path.abspath(db_path)
    if not os.path.exists(db_path):
        print(f"【エラー】Accessファイルが見つかりません: {db_path}")
        return

    db_name = Path(db_path).stem
    output_dir = Path(output_base) / f"{db_name}_assets"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"開始: {db_path}")
    
    access = win32com.client.Dispatch("Access.Application")
    try:
        db = access.DBEngine.OpenDatabase(db_path)
        
        # --- クエリの抽出 ---
        query_dir = output_dir / "queries"
        query_dir.mkdir(exist_ok=True)
        for qdf in db.QueryDefs:
            if not qdf.Name.startswith("~"):
                file_path = query_dir / f"{qdf.Name}.sql"
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(qdf.SQL)
        print(f" - クエリ完了: {len(db.QueryDefs)}件")
        
        # --- テーブル構造(DDL)の抽出 ---
        table_dir = output_dir / "tables"
        table_dir.mkdir(exist_ok=True)
        for tdf in db.TableDefs:
            if not tdf.Name.startswith("MSys") and not tdf.Name.startswith("~"):
                # 1. フィールド定義の取得
                fields = [f"    [{f.Name}] {get_type_name(f.Type)}" for f in tdf.Fields]
                
                # 2. 主キー(Primary Key)の取得
                pk_fields = []
                try:
                    for idx in tdf.Indexes:
                        if idx.Primary:
                            # 複合主キーの場合もあるためループで取得
                            pk_fields = [f"[{f.Name}]" for f in idx.Fields]
                            break
                except Exception:
                    pass # インデックスがないテーブル用
                
                # 3. SQLの組み立て
                sql_lines = fields.copy()
                if pk_fields:
                    sql_lines.append(f"    CONSTRAINT PK_{tdf.Name} PRIMARY KEY ({', '.join(pk_fields)})")
                
                create_sql = f"CREATE TABLE [{tdf.Name}] (\n" + ",\n".join(sql_lines) + "\n);"
                
                file_path = table_dir / f"{tdf.Name}.sql"
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(create_sql)
        print(f" - テーブル構造完了: {len(db.TableDefs)}件")

        db.Close()
        print(f"\n【成功】出力先: {output_dir}")

    except Exception as e:
        print(f"【実行エラー】: {e}")
    finally:
        access.Quit()

def get_type_name(type_id):
    """DAOのデータ型IDを文字列に変換"""
    type_map = {1: "YESNO", 3: "INTEGER", 4: "LONG", 5: "CURRENCY", 
                6: "SINGLE", 7: "DOUBLE", 8: "DATETIME", 10: "TEXT", 12: "MEMO"}
    return type_map.get(type_id, "TEXT")

if __name__ == "__main__":
    import sys
    # 第一引数にDBパス、第二引数に出力先を受け取る想定
    target_db = sys.argv[1] if len(sys.argv) > 1 else r"D:\My_DataBase\RN_ProjectSummary.accdb"
    output_path = os.path.dirname(os.path.abspath(__file__))
    extract_access_assets(target_db, output_path)