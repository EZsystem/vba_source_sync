import os
import sys
import re
import argparse
import httpx
from pathlib import Path
from loguru import logger

# ログ設定：標準出力には重要な成功/エラーメッセージのみを出す
logger.remove()
logger.add(sys.stderr, format="<green>{time:HH:mm:ss}</green> | <level>{level: <8}</level> | <level>{message}</level>", level="INFO")

def extract_headers(md_path):
    """Markdownファイルからテーブルのヘッダー行を抽出する（機密保持のため出力は最小限）"""
    try:
        with open(md_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("|") and line.endswith("|"):
                    if re.match(r"^\|[\s\-\|:]+\|$", line):
                        continue
                    headers = [h.strip() for h in line.split("|") if h.strip()]
                    if headers:
                        return headers
    except Exception:
        pass # インターフェースを汚さないよう詳細エラーは後続に任せる
    return None

def generate_ai_samples(headers, model_name="gemma-4-e4b-it"):
    """ローカルLLMに接続してテストデータを生成する（プロンプト等の中身はログ出力しない）"""
    url = "http://localhost:1234/v1/chat/completions"
    
    header_str = ", ".join(headers)
    system_prompt = (
        "あなたは岩手県の建設業界（岩手県庁、盛岡市役所などの公共工事発注案件）に非常に精通した専門家です。"
        "岩手県内の具体的な地名（盛岡市、北上市、釜石市、宮古市、一関市、二戸市など）や、"
        "現実的な工種（アスファルト舗装工、治山工事、河川改修、雪寒対策工事など）を理解しています。"
    )
    
    user_prompt = (
        f"以下のフィールド名を持つ、岩手県の建設工事らしい架空のテストデータを10件生成してください。\n"
        f"フィールド: [{header_str}]\n\n"
        "### 制約事項:\n"
        "1. 出力は Markdown テーブル形式のみとしてください。\n"
        "2. ヘッダー行を含め、10件のデータ行を出力してください。\n"
        "3. 余計な前置きや解説文は一切不要です。\n"
        "4. 岩手県内の実在する地名や、業界で使われるリアルな用語（砂利敷、重機損料、路盤材等）を盛り込んでください。"
    )

    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.7
    }

    # データの中身は表示せず、アクションのみを記録
    logger.info(f"ローカルAIにデータ生成を依頼中 (Model: {model_name})...")
    
    try:
        # タイムアウトを120秒に延長（大規模生成に備え）
        with httpx.Client(timeout=120.0) as client:
            response = client.post(url, json=payload)
            response.raise_for_status()
            result = response.json()
            return result["choices"][0]["message"]["content"]
    except httpx.ConnectError:
        logger.error("Error: LM Studio (localhost:1234) が起動していません。")
    except Exception as e:
        logger.error(f"Error: AI連携エラーが発生しました。")
    
    return None

def main():
    parser = argparse.ArgumentParser(description="MarkdownヘッダーからAIでサンプルデータを生成します（Silent Mode）")
    parser.add_argument("input", help="対象となるMarkdownファイルのパス")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        logger.error(f"Error: フォルダが見つかりません: {input_path.name}")
        return

    # 1. ヘッダー抽出（中身は表示しない）
    headers = extract_headers(input_path)
    if not headers:
        logger.error(f"Skip: {input_path.name} からヘッダーを読み取れませんでした。")
        return

    # 2. AIによるデータ生成
    ai_content = generate_ai_samples(headers)
    if not ai_content:
        return

    # 3. ファイル保存 (UTF-8 BOMなし)
    output_path = input_path.parent / f"{input_path.stem}_sample.md"
    try:
        with open(output_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(ai_content)
        # 成功メッセージのみを表示（データは non-traceable）
        logger.success(f"[完了] サンプルファイルを生成しました: {output_path.name}")
    except Exception:
        logger.error(f"Error: {output_path.name} の保存に失敗しました。")

if __name__ == "__main__":
    main()
