CREATE TABLE [✕at_原価S_ColSetting] (
    [配列列番号] LONG,
    [配列タイトル名] TEXT,
    [accテーブル名] TEXT,
    [取込フラグ] YESNO,
    [データ型] TEXT,
    [空欄対応モード] TEXT,
    [基本工事_仮テーブルタイトル名] TEXT,
    [基本工事_本テーブルタイトル名] TEXT,
    [枝番工事_仮テーブルタイトル名] TEXT,
    [枝番工事_本テーブルタイトル名] TEXT,
    CONSTRAINT PK_✕at_原価S_ColSetting PRIMARY KEY ([配列列番号])
);