SELECT kt.s基本工事コード, kt.s基本工事名称, kt.工事コード, kt.工事名称, kt.枝番工事コード, kt.追加工事名称, kt.工事価格, kt.施工管轄組織名
FROM q_Icube_WithMngNo AS kt LEFT JOIN at_Genka_Edaban AS g ON (kt.工事コード = g.工事コード) AND (kt.管理番号 = g.管理番号)
WHERE kt.所属組織名 = "ＬＣＳ事業部" 
  AND kt.一件工事判定 = "小口工事" 
  AND kt.完工期 = 13 
  AND g.工事コード IS NULL
ORDER BY kt.施工管轄組織コード;
