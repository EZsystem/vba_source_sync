CREATE TABLE [at_建物用途正誤表] (
    [ID] LONG,
    [誤_用途大区分] TEXT,
    [誤_用途大区分名] TEXT,
    [正_用途大区分] TEXT,
    [正_用途大区分名] TEXT,
    CONSTRAINT PK_at_建物用途正誤表 PRIMARY KEY ([ID])
);