CREATE TABLE [at_Work_小口完工推移] (
    [ID] LONG,
    [施工管轄組織名] TEXT,
    [受注期_] TEXT,
    [完工期_] TEXT,
    [受注月_] TEXT,
    [完工月_] TEXT,
    [工事価格の合計] CURRENCY,
    CONSTRAINT PK_at_Work_小口完工推移 PRIMARY KEY ([ID])
);