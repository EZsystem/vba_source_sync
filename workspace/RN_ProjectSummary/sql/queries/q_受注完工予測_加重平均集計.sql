SELECT K.[期_計算対象], Val(Nz(K.[期_計算対象],0)) AS 期_計算対象数値, T.[施工管轄組織名], Switch(T.[受注月] In (4,5,6),"1Q", T.[受注月] In (7,8,9),"2Q", T.[受注月] In (10,11,12),"3Q", T.[受注月] In (1,2,3),"4Q") AS Q, T.[受注月] AS 受注月数値, T.[受注月] & "月" AS 受注月表示, Int(Nz(Sum(CDbl(Nz(T.[合計価格],0)) * CDbl(Nz(K.[加重率数値],0))), 0) + 0.5) AS 加重受注高
FROM (SELECT [施工管轄組織名], [受注月], [受注期], Sum([工事価格]) AS [合計価格]
        FROM [at_Icube_累計]
        WHERE [一件工事判定] = '小口工事' 
          AND [受注月] Is Not Null 
          AND [受注期] Is Not Null
        GROUP BY [施工管轄組織名], [受注月], [受注期]
    )  AS T INNER JOIN (SELECT [作業所名], [期_計算対象], Val(Replace(Nz([期_荷重対象],""), "期", "")) AS [荷重期数値], Val(Nz([加重率],0)) AS [加重率数値]
        FROM [at_受注額予測計数]
        WHERE [作業所名] Is Not Null 
          AND [期_荷重対象] Is Not Null
    )  AS K ON (T.[施工管轄組織名] = K.[作業所名]) AND (T.[受注期] = K.[荷重期数値])
GROUP BY K.[期_計算対象], Val(Nz(K.[期_計算対象],0)), T.[施工管轄組織名], Switch(T.[受注月] In (4,5,6),"1Q", T.[受注月] In (7,8,9),"2Q", T.[受注月] In (10,11,12),"3Q", T.[受注月] In (1,2,3),"4Q"), T.[受注月]
ORDER BY Val(Nz(K.[期_計算対象],0)), T.[施工管轄組織名], T.[受注月];
