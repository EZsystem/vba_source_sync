CREATE TABLE [kt_基本工事_完工] (
    [データ年月（受注計上年月）] LONG,
    [基本工事コード] TEXT,
    [基本工事名称] TEXT,
    [完工年度] INTEGER,
    [完工期] INTEGER,
    [完工Q] INTEGER,
    [完工月] INTEGER,
    [完工日_日付型] DATETIME,
    [支払最終期] INTEGER,
    CONSTRAINT PK_kt_基本工事_完工 PRIMARY KEY ([基本工事コード])
);