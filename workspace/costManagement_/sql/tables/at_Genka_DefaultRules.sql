CREATE TABLE [at_Genka_DefaultRules] (
    [配列列番号] LONG,
    [配列タイトル名] TEXT,
    [accテーブル名] TEXT,
    [取込フラグ] YESNO,
    [データ型] TEXT,
    [空欄対応モード] TEXT,
    CONSTRAINT PK_at_Genka_DefaultRules PRIMARY KEY ([配列列番号])
);