CREATE TABLE [at_基本工事_作業所] (
    [No] LONG,
    [データ年月（受注計上年月）] LONG,
    [基本工事コード] TEXT,
    [基本工事名称] TEXT,
    [施工担当組織コード] TEXT,
    [施工担当組織名] TEXT,
    [施工管轄組織コード] TEXT,
    [施工管轄組織名] TEXT,
    [所属組織コード] TEXT,
    [所属組織名] TEXT,
    [一件工事判定] TEXT,
    CONSTRAINT PK_at_基本工事_作業所 PRIMARY KEY ([基本工事コード])
);