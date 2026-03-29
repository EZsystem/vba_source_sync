SELECT g.基本工事コード, g.基本工事名, g.工事コード, g.工事名, kt.枝番工事コード, g.枝番工事コード, kt.追加工事名称, g.追加工事名称, g.枝番工事コードエラーチェック
FROM at_原価S_枝番工事 AS g INNER JOIN at_Icube_累計 AS kt ON g.枝番工事コード = kt.枝番工事コード;
