SELECT Format([対象年月], 'yyyy/mm') AS 年月, 作業所名, 工事コード, 工事名, 
       Sum([金額]) AS 集計金額
FROM at_Work_給与経費集計_結果
WHERE 工事コード <> 'EE0000100'
GROUP BY 対象年月, 作業所名, 工事コード, 工事名
ORDER BY 対象年月, 工事コード;
