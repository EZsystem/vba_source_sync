CREATE TABLE [at_原価S_ColSetting] (
    [配列列番号] LONG,
    [配列タイトル名] TEXT,
    [accテーブル名] TEXT,
    [取込フラグ] YESNO,
    [データ型] TEXT,
    [空欄対応モード] TEXT,
    CONSTRAINT PK_at_原価S_ColSetting PRIMARY KEY ([配列列番号])
);