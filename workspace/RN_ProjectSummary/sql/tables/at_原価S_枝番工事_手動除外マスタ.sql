CREATE TABLE [at_原価S_枝番工事_手動除外マスタ] (
    [ID] LONG,
    [枝番工事コード] TEXT,
    [追加工事名称] TEXT,
    [除外_利益エラー] YESNO,
    [除外_支払エラー] YESNO,
    [備考] TEXT,
    CONSTRAINT PK_at_原価S_枝番工事_手動除外マスタ PRIMARY KEY ([ID])
);