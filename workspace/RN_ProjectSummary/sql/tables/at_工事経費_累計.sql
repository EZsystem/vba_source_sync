CREATE TABLE [at_工事経費_累計] (
    [ID] LONG,
    [年月] DATETIME,
    [仮基本工事コード] TEXT,
    [工事名] TEXT,
    [経費名] TEXT,
    [経費額] CURRENCY,
    [作業所名] TEXT,
    [期] TEXT,
    [Q] TEXT,
    CONSTRAINT PK_at_工事経費_累計 PRIMARY KEY ([ID])
);