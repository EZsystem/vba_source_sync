CREATE TABLE [tbl_予定工事_随時更新] (
    [仮基本工事コード] TEXT,
    [基本工事コード_仮] TEXT,
    [基本工事名称_仮] TEXT,
    [施工管轄組織コード] TEXT,
    [施工管轄組織名] TEXT,
    [作業所_略称] TEXT,
    [作業所_県] TEXT,
    [官民] TEXT,
    [完工_期] INTEGER,
    [完工_Q] INTEGER,
    [一件工事判定] TEXT,
    CONSTRAINT PK_tbl_予定工事_随時更新 PRIMARY KEY ([基本工事コード_仮])
);