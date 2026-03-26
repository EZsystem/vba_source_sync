SELECT at_Genka_Edaban.基本工事コード, at_Genka_Edaban.基本工事名, at_Genka_Edaban.工事コード, at_Genka_Edaban.工事名, at_Icube_Archive.枝番工事コード, at_Genka_Edaban.枝番工事コード, at_Icube_Archive.追加工事名称, at_Genka_Edaban.追加工事名称, at_Genka_Edaban.枝番工事コードエラーチェック
FROM at_Genka_Edaban INNER JOIN at_Icube_Archive ON at_Genka_Edaban.枝番工事コード = at_Icube_Archive.枝番工事コード;
