SELECT Icube_累計.追加工事名称, Sum(Icube_累計.工事価格) AS 工事価格の合計, Sum(Icube_累計.粗利益額) AS 粗利益額の合計, Icube_累計.受注期, Icube_累計.受注Q, Icube_累計.施工管轄組織名, Icube_累計.一件工事判定, Icube_累計.所属組織名, Icube_累計.[リニューアル環境区分名]
FROM Icube_累計
GROUP BY Icube_累計.追加工事名称, Icube_累計.受注期, Icube_累計.受注Q, Icube_累計.施工管轄組織名, Icube_累計.一件工事判定, Icube_累計.所属組織名, Icube_累計.[リニューアル環境区分名]
HAVING (((Icube_累計.受注期)=13) AND (Not (Icube_累計.施工管轄組織名)="ビルサービスグループ") AND ((Icube_累計.所属組織名)="ＬＣＳ事業部") AND ((Icube_累計.[リニューアル環境区分名])="リニューアル")) OR (((Icube_累計.所属組織名)="建築部"));
