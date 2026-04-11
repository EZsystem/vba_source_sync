CREATE TABLE [_at_社員給与] (
    [ID] LONG,
    [資格名] TEXT,
    [本年度] DOUBLE,
    [2025年度] DOUBLE,
    [総合職該当] YESNO,
    [支払月オフセット] DOUBLE,
    CONSTRAINT PK__at_社員給与 PRIMARY KEY ([ID])
);