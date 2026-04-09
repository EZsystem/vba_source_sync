SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS 受注期表示, ([受注Q] & 'Q') AS 受注Q表示, [受注月] & '月' AS 受注月表示, at_Icube_累計.受注形態名, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計
FROM at_Icube_累計
WHERE (((at_Icube_累計.施工管轄組織名)<>'ビルサービスグループ') AND ([受注期]=Val(Replace('13期','期',''))) AND ((at_Icube_累計.所属組織名)='ＬＣＳ事業部') AND ((at_Icube_累計.一件工事判定)<>'一件工事'))
GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), ([受注Q] & 'Q'), [受注月] & '月', at_Icube_累計.受注形態名;
