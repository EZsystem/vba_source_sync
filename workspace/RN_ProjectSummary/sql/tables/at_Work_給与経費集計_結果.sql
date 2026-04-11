CREATE TABLE [at_Work_給与経費集計_結果] (
    [ID] AUTOINCREMENT PRIMARY KEY,
    [データ区分] TEXT(20),
    [対象年月] DATETIME,
    [作業所名] TEXT(100),
    [工事コード] TEXT(50),
    [工事名] TEXT(255),
    [社員名] TEXT(100),
    [金額] CURRENCY
);
