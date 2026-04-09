SELECT Combined.仮基本工事コード, Combined.工事名, Combined.作業所名, Combined.期, Combined.Q, Sum(Combined.金額) AS [金額(集計)]
FROM (SELECT
    k.仮基本工事コード,
    k.工事名,
    k.作業所名,
    k.期,
    k.Q,
    (k.兼務率割合 * g.本年度) AS 金額
  FROM (_at_社員給与 AS g
  INNER JOIN at_社員情報 AS s ON g.資格名 = s.資格_想定給与額)
  INNER JOIN at_kenmu_累計 AS k ON s.氏名_ﾒｰﾙ表示用 = k.社員名
  WHERE s.在籍区分 = True
  
  UNION ALL
  

  SELECT
    仮基本工事コード,
    工事名,
    作業所名,
    期,
    Q,
    経費額 AS 金額
  FROM at_工事経費_累計
)  AS Combined
WHERE (Combined.工事名 NOT LIKE "*ＲＮ経費算出外*")
  AND (Combined.仮基本工事コード <> "EE0000100")
GROUP BY Combined.仮基本工事コード, Combined.工事名, Combined.作業所名, Combined.期, Combined.Q;
