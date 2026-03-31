CREATE TABLE [_at_ExportConfig] (
    [ID] LONG,
    [ProcessName] TEXT,
    [QueryName] TEXT,
    [SQLTemplate] MEMO,
    [ExcelPath] MEMO,
    [ExcelSheet] TEXT,
    [ExcelTable] TEXT,
    [IsActive] YESNO,
    CONSTRAINT PK__at_ExportConfig PRIMARY KEY ([ID])
);