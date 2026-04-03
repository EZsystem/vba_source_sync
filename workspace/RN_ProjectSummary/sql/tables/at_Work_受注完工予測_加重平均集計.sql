CREATE TABLE [at_Work_受注完工予測_加重平均集計] (
    [ID] LONG,
    [予測ターゲット] TEXT,
    [施工管轄組織名] TEXT,
    [Q] TEXT,
    [受注月] TEXT,
    [加重集計値] CURRENCY,
    CONSTRAINT PK_at_Work_受注完工予測_加重平均集計 PRIMARY KEY ([ID])
);