SELECT First(at_Icube_累計.基本工事コード) AS [基本工事コード ], First(at_Icube_累計.基本工事名称) AS [基本工事名称 ], Sum(at_Icube_累計.工事価格) AS 工事価格_合計, Sum(at_Icube_累計.粗利益額) AS 粗利益額_合計, at_Icube_累計.施工管轄組織名, (完工期 & '期') AS 完工期表示, (完工Q & 'Q') AS 完工Q表示, (完工月 & '月') AS 完工月表示, 基本工事名_官民
FROM at_Icube_累計
WHERE (施工管轄組織名 <> 'ビルサービスグループ') AND (完工期 = Val(Replace('14期','期',''))) AND (所属組織名 = 'ＬＣＳ事業部') AND (一件工事判定 = '小口工事')
GROUP BY 施工管轄組織名, 完工期, 完工Q, 完工月, 基本工事名_官民;
