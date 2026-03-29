SELECT Mid([枝番工事コード], InStrRev([枝番工事コード], "-") + 1) AS 管理番号, *
FROM at_Icube_累計;
