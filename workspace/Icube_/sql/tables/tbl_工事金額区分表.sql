CREATE TABLE [tbl_工事金額区分表] (
    [ID] LONG,
    [工事金額区分コード] TEXT,
    [工事金額区分名] TEXT,
    [最小金額] CURRENCY,
    [最大金額] CURRENCY,
    [工事金額マイナス判定] YESNO,
    CONSTRAINT PK_tbl_工事金額区分表 PRIMARY KEY ([ID])
);