SELECT Icube_累計.施工管轄組織名, Icube_累計.一件工事判定, Icube_累計.枝番工事コード, Icube_累計.追加工事名称, Icube_累計.工事価格, Icube_累計.粗利益額, Icube_累計.受注期, Icube_累計.受注Q, Icube_累計.受注月, Icube_累計.完工期, Icube_累計.完工Q, Icube_累計.完工月
FROM Icube_累計
WHERE (((Icube_累計.所属組織名)="ＬＣＳ事業部"))
ORDER BY Icube_累計.施工管轄組織コード;
