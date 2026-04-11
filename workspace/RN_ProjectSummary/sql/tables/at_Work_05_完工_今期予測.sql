CREATE TABLE [at_Work_05_完工_今期予測] (
    [ID] LONG,
    [期_予測ターゲット] TEXT,
    [施工管轄組織名] TEXT,
    [受注月] TEXT,
    [完工月] TEXT,
    [完工Q] TEXT,
    [工事コード] TEXT,
    [元_受注予測額] CURRENCY,
    [適用比率] DOUBLE,
    [予測完工高] CURRENCY,
    [予測経費額] CURRENCY,
    CONSTRAINT PK_at_Work_05_完工_今期予測 PRIMARY KEY ([ID])
);