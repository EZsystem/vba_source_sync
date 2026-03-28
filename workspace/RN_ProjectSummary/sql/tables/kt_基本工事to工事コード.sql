CREATE TABLE [kt_基本工事to工事コード] (
    [基本工事コード] TEXT,
    [基本工事名称] TEXT,
    [仮基本工事コード] TEXT,
    [一件工事判定] TEXT,
    [基本工事名_作業所] TEXT,
    [基本工事名_期] TEXT,
    [基本工事名_Q] TEXT,
    [基本工事名_官民] TEXT,
    [基本工事名_繰越] TEXT,
    CONSTRAINT PK_kt_基本工事to工事コード PRIMARY KEY ([基本工事コード])
);