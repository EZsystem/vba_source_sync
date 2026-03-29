SELECT Icube_.基本工事コード, Icube_.基本工事名称, Icube_.仮基本工事コード, Icube_.一件工事判定, Icube_.基本工事名_作業所, Icube_.基本工事名_年度, Icube_.基本工事名_期, Icube_.基本工事名_Q, Icube_.基本工事名_官民, Icube_.基本工事名_繰越
FROM Icube_
GROUP BY Icube_.基本工事コード, Icube_.基本工事名称, Icube_.仮基本工事コード, Icube_.一件工事判定, Icube_.基本工事名_作業所, Icube_.基本工事名_年度, Icube_.基本工事名_期, Icube_.基本工事名_Q, Icube_.基本工事名_官民, Icube_.基本工事名_繰越
HAVING (((Icube_.一件工事判定)="小口工事"));
