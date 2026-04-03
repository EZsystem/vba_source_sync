# =====================================
# Module: history_mgr.py
# 説明  : ファイル選択履歴（最新5件）を管理するクラス
# 作成日: 2026/04/03
# =====================================
import json
from pathlib import Path
from loguru import logger

class HistoryManager:
    def __init__(self, history_file="history.json"):
        # プロジェクトルート（srcの親）を基準に保存先を決定
        project_root = Path(__file__).parent.parent
        self.history_file = project_root / history_file
        self.max_history = 5
        self.history = self._load_history()

    def _load_history(self):
        """履歴ファイルを読み込む"""
        if not self.history_file.exists():
            return []
        try:
            with open(self.history_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data[:self.max_history]
        except Exception as e:
            logger.warning(f"履歴の読み込みに失敗しました: {e}")
        return []

    def save_history(self, new_path):
        """新しいパスを履歴に追加して保存する"""
        if not new_path:
            return
            
        # 重複を排除し、最新を先頭に持ってくる
        if new_path in self.history:
            self.history.remove(new_path)
        
        self.history.insert(0, new_path)
        
        # 最大件数に制限
        self.history = self.history[:self.max_history]
        
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logger.error(f"履歴の保存に失敗しました: {e}")

    def get_history(self):
        """現在の履歴リストを取得する"""
        return self.history
