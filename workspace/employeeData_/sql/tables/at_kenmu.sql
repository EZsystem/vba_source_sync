CREATE TABLE [at_kenmu] (
    [ImportID] LONG,
    [No] TEXT,
    [年月] DATETIME,
    [工事コード] TEXT,
    [工事名] TEXT,
    [コメント] TEXT,
    [社員名] TEXT,
    [兼務率割合] DOUBLE,
    [作業所名] TEXT,
    [元ファイルパス] TEXT,
    CONSTRAINT PK_at_kenmu PRIMARY KEY ([ImportID])
);