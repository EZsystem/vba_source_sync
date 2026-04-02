SELECT at_Icube_累計.施工管轄組織名, ([受注期] & '期') AS 受注期表示, at_Icube_累計.受注形態区分名, Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計
FROM at_Icube_累計
WHERE at_Icube_累計.施工管轄組織名 NOT IN ('ビルサービスグループ', '東北支店ＲＮ部')
    AND [受注期] = Val(Replace('13期','期',''))
    AND at_Icube_累計.所属組織名 = 'ＬＣＳ事業部'
GROUP BY at_Icube_累計.施工管轄組織名, ([受注期] & '期'), at_Icube_累計.受注形態区分名;
