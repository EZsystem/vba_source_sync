SELECT Icube_累計.工事コード, First(Icube_累計.工事名称) AS 工事名称, Sum(Icube_累計.工事価格) AS 工事価格_合計, Sum(Icube_累計.粗利益額) AS 粗利益額_合計, First(Icube_累計.施工管轄組織名) AS 施工管轄組織名, First([受注期] & "期") AS 受注期表示, First([受注Q] & "Q") AS 受注Q表示, [受注月] & "月" AS 受注月表示
FROM Icube_累計
WHERE (((Icube_累計.施工管轄組織名)<>"ビルサービスグループ") 
    AND (([受注期] & "期")="13期") 
    AND ((Icube_累計.一件工事判定)="一件工事") 
    AND ((Icube_累計.所属組織名)="ＬＣＳ事業部"))
GROUP BY Icube_累計.工事コード, [受注月] & "月";
