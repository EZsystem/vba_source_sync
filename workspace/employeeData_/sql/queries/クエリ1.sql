SELECT 仮基本工事コード, 工事名, 作業所名, 期, Q, Sum(金額) AS [金額(集計)]
FROM (SELECT
    k.仮基本工事コード,
    k.工事名,
    k.作業所名,
    k.期,
    k.Q,
    (k.兼務率割合 * s.給与相当額) AS 金額
  FROM at_kenmu_累計 AS k
  INNER JOIN at_社員情報 AS s ON k.社員名 = s.氏名_戸籍上
  
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
GROUP BY 仮基本工事コード, 工事名, 作業所名, 期, Q;
