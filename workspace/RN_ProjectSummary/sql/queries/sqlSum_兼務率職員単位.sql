SELECT Combined.作業所名, Combined.工事コード, Combined.工事名, Combined.社員名, Sum(Combined.金額) AS 想定給与額, Sum(Combined.総合職数) AS 総合職数, Sum(Combined.人員数) AS 人員数, Sum(Combined.数値割合) AS 数値
FROM (SELECT K.作業所名, K.工事コード, K.工事名, K.社員名, (G.本年度 * K.兼務率割合) AS 金額, IIf(G.総合職該当 = True, K.兼務率割合, 0) AS 総合職数, IIf(S.事務員区分 = True, 0, K.兼務率割合) AS 人員数, K.兼務率割合 AS 数値割合
  FROM (at_kenmu AS K INNER JOIN at_社員情報 AS S ON K.社員名 = S.氏名_ﾒｰﾙ表示用) INNER JOIN _at_社員給与 AS G ON S.資格_想定給与額 = G.資格名
  WHERE (((K.工事コード) <> "-") AND ((S.在籍区分) = True))
  
  UNION ALL
  
  SELECT 作業所名, 仮基本工事コード AS 工事コード, 工事名, 経費名 AS 社員名, 経費額 AS 金額, 0 AS 総合職数, 0 AS 人員数, 1 AS 数値割合
  FROM at_工事経費_累計
)  AS Combined
WHERE (Combined.工事コード <> "EE0000100")
GROUP BY Combined.作業所名, Combined.工事コード, Combined.工事名, Combined.社員名;
