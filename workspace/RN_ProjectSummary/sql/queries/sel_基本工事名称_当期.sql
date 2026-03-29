SELECT 
    Icube_累計.s基本工事コード AS s基本工事コード,
    Icube_累計.s基本工事名称 AS s基本工事名称,
    Icube_累計.仮基本工事コード AS 仮基本工事コード,
    Icube_累計.完工期 AS 完工期,
    Icube_累計.完工Q AS 完工Q,
    Icube_累計.施工管轄組織名 AS 施工管轄組織名,
    Icube_累計.一件工事判定 AS 一件工事判定
FROM Icube_累計
WHERE 
    Icube_累計.完工期 = 13
    AND Icube_累計.施工管轄組織名 IN (
        "青森ＲＮ（作）", "秋田ＲＮ（作）", "盛岡ＲＮ（作）", 
        "山形ＲＮ（作）", "福島ＲＮ（作）", "仙台ＲＮ（作）"
    )

UNION

SELECT 
    A.基本工事コード_仮 AS s基本工事コード,
    A.基本工事名称_仮 AS s基本工事名称,
    A.仮基本工事コード AS 仮基本工事コード,
    A.完工_期 AS 完工期,
    A.完工_Q AS 完工Q,
    A.施工管轄組織名 AS 施工管轄組織名,
    A.一件工事判定 AS 一件工事判定
FROM tbl_基本工事_期首入力 AS A
LEFT JOIN Icube_累計 AS B
    ON A.基本工事コード_仮 = B.s基本工事コード
WHERE 
    B.s基本工事コード IS NULL
    AND A.完工_期 = 13
    AND A.施工管轄組織名 IN (
        "青森ＲＮ（作）", "秋田ＲＮ（作）", "盛岡ＲＮ（作）", 
        "山形ＲＮ（作）", "福島ＲＮ（作）", "仙台ＲＮ（作）"
    )

UNION SELECT 
    X.基本工事コード_仮 AS s基本工事コード,
    X.基本工事名称_仮 AS s基本工事名称,
    X.仮基本工事コード AS 仮基本工事コード,
    X.完工_期 AS 完工期,
    X.完工_Q AS 完工Q,
    X.施工管轄組織名 AS 施工管轄組織名,
    X.一件工事判定 AS 一件工事判定
FROM tbl_予定工事_随時更新 AS X
LEFT JOIN Icube_累計 AS Y
    ON X.基本工事コード_仮 = Y.s基本工事コード
WHERE 
    Y.s基本工事コード IS NULL
    AND X.完工_期 = 13
    AND X.施工管轄組織名 IN (
        "青森ＲＮ（作）", "秋田ＲＮ（作）", "盛岡ＲＮ（作）", 
        "山形ＲＮ（作）", "福島ＲＮ（作）", "仙台ＲＮ（作）"
    );
