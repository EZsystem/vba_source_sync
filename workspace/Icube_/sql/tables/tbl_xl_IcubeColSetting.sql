CREATE TABLE [tbl_xl_IcubeColSetting] (
    [配列列番号] LONG,
    [デフォルトタイトルフィールド名] TEXT,
    [タイトル名_デフォルト] TEXT,
    [タイトル名_置換え後] TEXT,
    [取込フラグ] YESNO,
    [データ型] TEXT,
    [空欄対応モード] TEXT,
    CONSTRAINT PK_tbl_xl_IcubeColSetting PRIMARY KEY ([デフォルトタイトルフィールド名])
);