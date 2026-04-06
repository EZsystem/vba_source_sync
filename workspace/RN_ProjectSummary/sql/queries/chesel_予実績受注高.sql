SELECT at_Icube_累計.追加工事名称, Sum(at_Icube_累計.工事価格) AS 工事価格の合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額の合計, at_Icube_累計.受注期, at_Icube_累計.受注Q, at_Icube_累計.施工管轄組織名, at_Icube_累計.一件工事判定, at_Icube_累計.所属組織名, at_Icube_累計.[リニューアル環境区分名]
FROM at_Icube_累計
GROUP BY at_Icube_累計.追加工事名称, at_Icube_累計.受注期, at_Icube_累計.受注Q, at_Icube_累計.施工管轄組織名, at_Icube_累計.一件工事判定, at_Icube_累計.所属組織名, at_Icube_累計.[リニューアル環境区分名]
HAVING (((at_Icube_累計.受注期)=13) AND ((at_Icube_累計.施工管轄組織名)<>"ビルサービスグループ") AND ((at_Icube_累計.所属組織名)="ＬＣＳ事業部") AND ((at_Icube_累計.[リニューアル環境区分名])="リニューアル")) OR (((at_Icube_累計.所属組織名)="建築部"));
