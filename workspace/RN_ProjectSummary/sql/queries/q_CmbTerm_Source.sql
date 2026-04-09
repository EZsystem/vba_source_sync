SELECT T.期 & '期' AS 期表示
FROM (SELECT DISTINCT 受注期 AS 期 FROM at_Icube_累計
    UNION
    SELECT MAX(受注期) + 1 AS 期 FROM at_Icube_累計
)  AS T
WHERE T.期 IS NOT NULL
ORDER BY T.期 DESC;
