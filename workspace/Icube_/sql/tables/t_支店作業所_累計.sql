CREATE TABLE [t_支店作業所_累計] (
    [部門コード] TEXT,
    [部門] TEXT,
    [組織コード] TEXT,
    [組織名] TEXT,
    [組織略名] TEXT,
    [上位組織コード] TEXT,
    [上位組織] TEXT,
    [セグメント区分] TEXT,
    [組織種類] TEXT,
    [使用開始日] TEXT,
    [撤廃年月日] TEXT,
    [施工管轄組織コード] TEXT,
    CONSTRAINT PK_t_支店作業所_累計 PRIMARY KEY ([組織コード])
);