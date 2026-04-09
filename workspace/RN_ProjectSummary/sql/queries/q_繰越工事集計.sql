SELECT [at_Icube_累計].[施工管轄組織名], [at_Icube_累計].[完工期] & '期' AS 期, Sum([at_Icube_累計].[工事価格]) AS 合計工事価格, Sum([at_Icube_累計].[粗利益額]) AS 合計粗利益額
FROM at_Icube_累計
WHERE ([at_Icube_累計].[一件工事判定] = '小口工事') AND ([at_Icube_累計].[基本工事名_繰越] = '(繰越)') AND ([at_Icube_累計].[所属組織名] = 'ＬＣＳ事業部') AND ([at_Icube_累計].[完工期] = Val(Replace('14期','期','')))
GROUP BY [at_Icube_累計].[施工管轄組織名], [at_Icube_累計].[完工期];
