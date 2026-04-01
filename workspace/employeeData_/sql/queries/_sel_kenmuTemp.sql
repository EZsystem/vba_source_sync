SELECT [at_社員情報main].社員番号, [at_社員情報main].氏名_戸籍上, [at_社員情報main].資格, 社員情報_事務職区分.事務職, 社員情報_在籍区分.在籍区分_年度管理
FROM 社員情報_在籍区分 RIGHT JOIN (社員情報_事務職区分 RIGHT JOIN (社員情報_rank RIGHT JOIN at_社員情報main ON 社員情報_rank.社員番号=[at_社員情報main].社員番号) ON 社員情報_事務職区分.社員番号=[at_社員情報main].社員番号) ON 社員情報_在籍区分.社員番号=[at_社員情報main].社員番号
WHERE (((社員情報_在籍区分.在籍区分_年度管理)="在籍"));
