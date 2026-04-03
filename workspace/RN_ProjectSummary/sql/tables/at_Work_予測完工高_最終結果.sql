CREATE TABLE [at_Work_予測完工高_最終結果] (
    [ID] LONG,
    [期_予測ターゲット] TEXT,
    [施工管轄組織名] TEXT,
    [受注月] TEXT,
    [完工月] TEXT,
    [元_受注予測額] CURRENCY,
    [適用比率] DOUBLE,
    [予測完工高] CURRENCY,
    CONSTRAINT PK_at_Work_予測完工高_最終結果 PRIMARY KEY ([ID])
);