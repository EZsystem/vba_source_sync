SELECT at_Icube_累計.[施工管轄組織名], at_Icube_累計.[受注期] & '期' AS 受注期_, at_Icube_累計.[完工期] & '期' AS 完工期_, at_Icube_累計.[受注月] & '月' AS 受注月_, at_Icube_累計.[完工月] & '月' AS 完工月_, CCur(Nz(Sum(at_Icube_累計.[工事価格]), 0)) AS 工事価格の合計
FROM at_Icube_累計
WHERE at_Icube_累計.[施工管轄組織名] IN ('青森ＲＮ（作）', '秋田ＲＮ（作）', '盛岡ＲＮ（作）', '山形ＲＮ（作）', '福島ＲＮ（作）', '仙台ＲＮ（作）')
    AND at_Icube_累計.[一件工事判定] = '小口工事'
    AND at_Icube_累計.[受注期] >= (SELECT Max([受注期]) FROM [at_Icube_累計]) - 2
GROUP BY at_Icube_累計.[施工管轄組織名], at_Icube_累計.[受注期], at_Icube_累計.[完工期], at_Icube_累計.[受注月], at_Icube_累計.[完工月]
ORDER BY at_Icube_累計.[施工管轄組織名], at_Icube_累計.[受注期], at_Icube_累計.[完工期], at_Icube_累計.[受注月], at_Icube_累計.[完工月];
