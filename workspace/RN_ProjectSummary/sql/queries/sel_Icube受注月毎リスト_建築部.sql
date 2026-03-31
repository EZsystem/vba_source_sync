SELECT First([at_Icube_累計].施工管轄組織名) AS 施工管轄組織名, First([受注期] & "期") AS 受注期表示, [受注Q] & "Q" AS 受注Q表示, Sum([at_Icube_累計].工事価格) AS 工事価格_合計, Sum([at_Icube_累計].粗利益額) AS 粗利益額_合計, First([at_Icube_累計].一件工事判定) AS 一件工事判定, First([at_Icube_累計].[リニューアル環境区分名]) AS リニューアル環境区分名
FROM at_Icube_累計
WHERE ((([受注期] & "期")="13期") And (([at_Icube_累計].[リニューアル環境区分名])="リニューアル") And (([at_Icube_累計].所属組織名)="建築部"))
GROUP BY [受注Q] & "Q";
