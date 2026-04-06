CREATE TABLE [at_Work_03_完工_推移割合] (
    [ID] LONG,
    [期_予測ターゲット] TEXT,
    [受注月] TEXT,
    [完工月] TEXT,
    [3期平均値] CURRENCY,
    [完工高割合] DOUBLE,
    CONSTRAINT PK_at_Work_03_完工_推移割合 PRIMARY KEY ([ID])
);