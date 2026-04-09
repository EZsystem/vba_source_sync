SELECT at_kenmu.作業所名, at_kenmu.社員名, [_at_社員給与].資格名, [_at_社員給与].本年度, [_at_社員給与].総合職該当, at_社員情報.事務員区分
FROM (_at_社員給与 INNER JOIN at_社員情報 ON [_at_社員給与].[資格名] = at_社員情報.[資格_想定給与額]) INNER JOIN at_kenmu ON at_社員情報.[氏名_ﾒｰﾙ表示用] = at_kenmu.[社員名]
WHERE (((at_社員情報.在籍区分)=True))
GROUP BY at_kenmu.作業所名, at_kenmu.社員名, [_at_社員給与].資格名, [_at_社員給与].本年度, [_at_社員給与].総合職該当, at_社員情報.事務員区分
ORDER BY at_kenmu.作業所名;
