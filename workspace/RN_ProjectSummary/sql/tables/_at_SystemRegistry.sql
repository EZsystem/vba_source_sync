CREATE TABLE [_at_SystemRegistry] (
    [ID] LONG,
    [業務区分] TEXT,
    [処理名称] TEXT,
    [処理説明] TEXT,
    [ファイル種別] TEXT,
    [既定パス] TEXT,
    [実行マクロ名] TEXT,
    [実行前確認メッセージ] TEXT,
    [フォーム表示有効フラグ] YESNO,
    [最終実行日時] DATETIME,
    [最終実行結果] TEXT,
    CONSTRAINT PK__at_SystemRegistry PRIMARY KEY ([ID])
);