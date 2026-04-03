SELECT Icube_.s基本工事コード, Icube_.s基本工事名称, Icube_.完工期, Icube_.施工管轄組織名, Icube_.一件工事判定
FROM Icube_
GROUP BY Icube_.s基本工事コード, Icube_.s基本工事名称, Icube_.完工期, Icube_.施工管轄組織名, Icube_.一件工事判定, Icube_.施工管轄組織コード, Icube_.所属組織名
HAVING (((Icube_.完工期)>=13) AND ((Icube_.施工管轄組織名)<>"ビルサービスグループ") AND ((Icube_.所属組織名)="ＬＣＳ事業部"))
ORDER BY Icube_.施工管轄組織コード;
