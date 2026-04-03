import sys
from pathlib import Path
import json

# srcディレクトリをパスに追加
sys.path.append(str(Path(__file__).parent.parent))
from src.history_mgr import HistoryManager

def test_history():
    history_file = "test_history.json"
    full_path = Path(__file__).parent.parent / history_file
    if full_path.exists():
        full_path.unlink()
        
    mgr = HistoryManager(history_file=history_file)
    
    # 1. 初期状態
    assert mgr.get_history() == []
    
    # 2. 追加
    mgr.save_history("path1")
    assert mgr.get_history() == ["path1"]
    
    # 3. 重複追加
    mgr.save_history("path2")
    mgr.save_history("path1")
    assert mgr.get_history() == ["path1", "path2"]  # path1がトップにくる
    
    # 4. 5件制限
    mgr.save_history("path3")
    mgr.save_history("path4")
    mgr.save_history("path5")
    mgr.save_history("path6")
    assert len(mgr.get_history()) == 5
    assert mgr.get_history()[0] == "path6"
    assert "path2" not in mgr.get_history()  # 最古のpath2が消える
    
    print("Test passed!")
    if full_path.exists():
        full_path.unlink()

if __name__ == "__main__":
    test_history()
