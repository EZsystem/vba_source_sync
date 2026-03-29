SELECT Icube_.基本工事コード, Icube_.基本工事名称, Icube_.工事コード, Icube_.工事名称, Icube_.枝番工事コード, Icube_.追加工事名称, Icube_.受注期, Icube_.受注Q, Icube_.受注月, Icube_.完工期, Icube_.完工Q, Icube_.完工月, Icube_.工事価格, Icube_.粗利益額, Icube_.工事金額区分コード, Icube_.工事金額区分名, Icube_.工事金額マイナス判定, Icube_.一件工事判定, Icube_.発注者コード, Icube_.発注者略称, Icube_.受注形態区分, Icube_.受注形態区分名, Icube_.施主種別, Icube_.施主種別名, Icube_.官民区分, Icube_.官民区分名, Icube_.[リニューアル環境区分], Icube_.[リニューアル環境区分名], Icube_.所属組織コード, Icube_.所属組織名, Icube_.施工管轄組織コード, Icube_.施工管轄組織名, Icube_.s用途大区分, Icube_.s用途大区分名
FROM Icube_
WHERE (((Icube_.[リニューアル環境区分])=1));
