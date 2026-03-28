CREATE TABLE [at_基本工事_受注] (
    [No] LONG,
    [データ年月（受注計上年月）] LONG,
    [基本工事コード] TEXT,
    [基本工事名称] TEXT,
    [受注年度] INTEGER,
    [受注期] INTEGER,
    [受注Q] INTEGER,
    [受注月] INTEGER,
    [受注計上日_日付型] DATETIME,
    CONSTRAINT PK_at_基本工事_受注 PRIMARY KEY ([基本工事コード])
);