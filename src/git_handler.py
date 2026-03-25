import os
from git import Repo
from dotenv import load_dotenv
from loguru import logger
from pathlib import Path

class GitManager:
    def __init__(self, repo_path):
        load_dotenv()
        self.repo_path = repo_path
        self.token = os.getenv("GITHUB_TOKEN")
        self.remote_url = os.getenv("GITHUB_REPO")
        
        # リモートURLにトークンを埋め込む（認証の自動化）
        if self.token and "https://" in self.remote_url:
            self.authenticated_url = self.remote_url.replace("https://", f"https://{self.token}@")
        else:
            self.authenticated_url = self.remote_url

    def sync_to_github(self, commit_message):
        try:
            # 1. リポジトリの取得または初期化
            if not (Path(self.repo_path) / ".git").exists():
                logger.info("リポジトリを初期化します...")
                repo = Repo.init(self.repo_path)
                repo.create_remote("origin", self.authenticated_url)
            else:
                repo = Repo(self.repo_path)

            # 2. 全ての変更をステージング (git add .)
            repo.git.add(A=True)

            # 3. 変更がある場合のみコミット
            if repo.is_dirty(untracked_files=True):
                repo.index.commit(commit_message)
                logger.info(f"コミット完了: {commit_message}")

                # 4. Push実行
                origin = repo.remote(name="origin")
                origin.push()
                logger.success("GitHubへのPushに成功しました。")
            else:
                logger.info("変更がないため、Pushをスキップしました。")

        except Exception as e:
            logger.error(f"Git操作中にエラーが発生しました: {e}")
            raise
