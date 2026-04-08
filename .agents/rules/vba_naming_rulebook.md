# 【Ez式】Excel テーブル全般 命名規則・ルールブック (v3.0)

本プロジェクトおよび関連するすべての Excel テーブル命名において、以下のルールを厳守するものとする。

## 1. 原則（Basic Principles）
*   **統一接頭辞 `xt_`**: Excel全般のシステム・作業用テーブルを示す。（従来の `xl_` は廃止し `xt_` へ統合・置換する）
*   **構成**: `[Prefix]_[Category]_[DetailConnection]`
    *   例: `xt_Ord_MonMin`
*   **文末連結ルール**: 最後の2つの単語（期間・属性等）はアンダーバー `_` を含まず、単語頭文字を大文字にして連結（結合）させる。
*   **キャメルケース**: 接頭辞（小文字）以外のすべての単語は、先頭を大文字とする。

## 2. カテゴリー（Common Categories）
*   `Ord` : 受注 (Order)
*   `Act` : 完工・実績 (Actuals)
*   `FCST`: 予測 (Forecast)
*   `exp` : 経費 (Expense)
*   `3Hist`: 3期分履歴 (3-term History)
*   `genkaS`: 原価管理システム (Cost Management System / genkaS)
*   `Mast`: 各種マスタ (Master)

## 3. 詳細および属性（Qualifiers & Logic）
*   `Mon` (月計), `Ql` (Q計), `Next` (次期), `Open` (期首)
*   `Min` (小口), `Pro` (大口プロジェクト), `Avg` (平均), `Trnd` (推移)
*   `Front` (支店/フロント), `Base` (基本/ベース), `Last` (最終), `No` (区分なし)

---
※このルールブックは .agents/rules/ に保存されており、本ディレクトリ下でのすべての Excel 操作において最優先で適用される。
