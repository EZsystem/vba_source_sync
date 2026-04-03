CREATE TABLE [at_Work_小口完工推移3期平均] (
    [ID] LONG,
    [期_予測ターゲット] TEXT,
    [受注月] TEXT,
    [完工月] TEXT,
    [3期平均値] CURRENCY,
    [完工高割合] DOUBLE,
    CONSTRAINT PK_at_Work_小口完工推移3期平均 PRIMARY KEY ([ID])
);