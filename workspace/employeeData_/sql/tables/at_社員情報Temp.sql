CREATE TABLE [at_社員情報Temp] (
    [明細情報] TEXT,
    [社員番号] TEXT,
    [氏名_戸籍上] TEXT,
    [氏名カナ] TEXT,
    [氏名_ﾒｰﾙ表示用] TEXT,
    [資格] TEXT,
    [所属] TEXT,
    [役職] TEXT,
    [対外呼称] TEXT,
    CONSTRAINT PK_at_社員情報Temp PRIMARY KEY ([社員番号])
);