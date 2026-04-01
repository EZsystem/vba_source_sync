CREATE TABLE [_at_公的資格リスト] (
    [資格コード] LONG,
    [資格免許名称] TEXT,
    [表示対象] YESNO,
    [取得奨励資格免許] YESNO,
    [昇格要件資格_総合職_地域総合職] YESNO,
    [昇格要件資格_地域職] YESNO,
    [画像保管] YESNO,
    [備考] TEXT,
    CONSTRAINT PK__at_公的資格リスト PRIMARY KEY ([資格コード])
);