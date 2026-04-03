SELECT B.受注期, B.受注月, B.完工期, B.完工月, Sum(A.加重受注高) AS 全店変数A, Sum(B.受注額集計) AS 全店変数B, Sum(Nz(B.[受注額集計], 0)) / IIf(Sum(Nz(A.[加重受注高], 0)) = 0, Null, Sum(A.[加重受注高])) AS 配分比率
FROM (SELECT [施工管轄組織名], [受注期], [受注月], [完工期], [完工月], Sum([工事価格]) AS 受注額集計
        FROM [at_Icube_累計]
        WHERE [一件工事判定] = '小口工事'
        GROUP BY [施工管轄組織名], [受注期], [受注月], [完工期], [完工月]
    )  AS B INNER JOIN q_受注完工予測_加重平均集計 AS A ON (B.[受注月] = A.[受注月数値]) AND (B.[受注期] = A.[期_計算対象数値]) AND (B.[施工管轄組織名] = A.[施工管轄組織名])
GROUP BY B.受注期, B.受注月, B.完工期, B.完工月;
