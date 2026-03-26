SELECT kt.s基本工事コード, kt.s基本工事名称, kt.施工管轄組織名
FROM at_Icube_Archive AS kt LEFT JOIN at_Genka_Kihon AS g ON kt.基本工事コード = g.基本工事コード
GROUP BY kt.s基本工事コード, kt.s基本工事名称, kt.施工管轄組織名, kt.施工管轄組織コード, kt.所属組織名, kt.一件工事判定, kt.完工期, g.基本工事コード
HAVING (((kt.所属組織名)="ＬＣＳ事業部") AND ((kt.一件工事判定)="小口工事") AND ((kt.完工期)=13) AND ((g.基本工事コード) Is Null))
ORDER BY kt.施工管轄組織コード;
