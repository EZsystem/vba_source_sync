SELECT at_Icube_累計.基本工事コード, First(at_Icube_累計.基本工事名称) AS 基本工事名称, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計, First(at_Icube_累計.施工管轄組織名) AS 施工管轄組織名, First([完工期] & "期") AS 完工期表示, First([完工Q] & "Q") AS 完工Q表示, [完工月] & "月" AS 完工月表示, at_Icube_累計.基本工事名_官民
FROM at_Icube_累計
WHERE (((at_Icube_累計.施工管轄組織名)<>"ビルサービスグループ") AND (([完工期] & "期")='14期') AND ((at_Icube_累計.所属組織名)="ＬＣＳ事業部") AND ((at_Icube_累計.一件工事判定)<>"一件工事"))
GROUP BY at_Icube_累計.基本工事コード, [完工月] & "月", at_Icube_累計.基本工事名_官民;
