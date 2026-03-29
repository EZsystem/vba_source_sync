SELECT DISTINCT s基本工事コード, s基本工事名称, 完工期, 完工Q, 施工管轄組織名, 一件工事判定
FROM (SELECT 
        Icube_.s基本工事コード AS s基本工事コード,
        Icube_.s基本工事名称 AS s基本工事名称,
        Icube_.完工期 AS 完工期,
        Icube_.完工Q AS 完工Q,
        Icube_.施工管轄組織名 AS 施工管轄組織名,
        Icube_.一件工事判定 AS 一件工事判定
    FROM Icube_

    UNION

    SELECT 
        tbl_基本工事_期首入力.基本工事コード_仮 AS s基本工事コード,
        tbl_基本工事_期首入力.基本工事名称_仮 AS s基本工事名称,
        tbl_基本工事_期首入力.完工_期 AS 完工期,
        tbl_基本工事_期首入力.完工_Q AS 完工Q,
        tbl_基本工事_期首入力.施工管轄組織名 AS 施工管轄組織名,
        tbl_基本工事_期首入力.一件工事判定 AS 一件工事判定
    FROM tbl_基本工事_期首入力
)  AS 統合元;
