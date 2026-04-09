CREATE TABLE [at_社員情報] (
    [社員番号] TEXT,
    [氏名_戸籍上] TEXT,
    [氏名カナ] TEXT,
    [氏名_ﾒｰﾙ表示用] TEXT,
    [資格] TEXT,
    [資格_想定給与額] TEXT,
    [所属] TEXT,
    [役職] TEXT,
    [対外呼称] TEXT,
    [事務員区分] YESNO,
    [在籍区分] YESNO,
    CONSTRAINT PK_at_社員情報 PRIMARY KEY ([社員番号])
);