CREATE TABLE [at_Genka_Edaban_Exclusion] (
    [ID] LONG,
    [枝番工事コード] TEXT,
    [追加工事名称] TEXT,
    [除外_利益エラー] YESNO,
    [除外_支払エラー] YESNO,
    [備考] TEXT,
    CONSTRAINT PK_at_Genka_Edaban_Exclusion PRIMARY KEY ([ID])
);