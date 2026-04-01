SELECT K.年月, K.作業所名, K.工事コード, K.工事名, K.社員名, Sum(G.本年度 * K.兼務率割合) AS 想定給与額, Sum(IIf(G.総合職該当 = True, K.兼務率割合, 0)) AS 総合職数, Sum(IIf(S.事務員区分 = True, 0, K.兼務率割合)) AS 人員数
FROM (at_kenmu AS K INNER JOIN at_社員情報 AS S ON K.社員名 = S.氏名_ﾒｰﾙ表示用) INNER JOIN _at_社員給与 AS G ON S.資格_想定給与額 = G.資格名
WHERE (((K.工事コード) <> "-") AND ((S.在籍区分) = True))
GROUP BY K.年月, K.作業所名, K.工事コード, K.工事名, K.社員名;
