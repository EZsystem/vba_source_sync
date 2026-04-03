CREATE TABLE [at_受注額予測計数] (
    [ID] LONG,
    [作業所名] TEXT,
    [期_計算対象] TEXT,
    [期_荷重対象] TEXT,
    [加重率] DOUBLE,
    CONSTRAINT PK_at_受注額予測計数 PRIMARY KEY ([ID])
);