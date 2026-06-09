-- 审核人 × 场次 × 属性项质量分析
-- 日期范围：2026-05-24 至 2026-06-02
-- 保持现有质量看板的属性项申报口径

WITH cte AS (
    SELECT
        DATE(asa.online1_complete_dt) AS online1_complete_date,
        asa.online1_observer_name AS observer_name,
        dl.cvalue_name AS sale_type,
        CASE
            WHEN asae.online1_cheat_proof_flag_name LIKE '%预质检测试成熟%' THEN '第1批'
            WHEN asae.online1_cheat_proof_flag_name LIKE '%预质检测试新%'   THEN '第2批'
            WHEN asae.online1_cheat_proof_flag_name RLIKE '预质检测试\\d+'
                THEN CONCAT(
                    '第',
                    REGEXP_EXTRACT(
                        asae.online1_cheat_proof_flag_name,
                        '预质检测试(\\d+)',
                        1
                    ),
                    '批'
                )
            WHEN asae.online1_cheat_proof_flag_name LIKE '%预质检测试%' THEN '第1批'
        END AS flag,
        asae.oneline1_property_names,
        asae.online1_mix_pass_property_name_names,
        CASE
            WHEN asa.online1_remark_txt LIKE '%举证%'
             AND asa.online1_result_name LIKE '%未通过%'
            THEN asae.oneline1_property_names
        END AS proof_failed_names,
        CASE
            WHEN asa.online1_result_name LIKE '%未通过%'
            THEN asae.oneline1_property_names
        END AS failed_names
    FROM dm.dm_centre_return_after_sale_analysis asa
    JOIN dm.dm_centre_return_after_sale_analysis_ext asae
        ON asa.return_order_no = asae.return_order_no
       AND asa.product_no = asae.product_no
    JOIN dw.dw_platform_settle_document_product s
        ON asa.settle_document_no = s.settle_document_no
       AND asa.product_no = s.product_no
       AND s.partition_flag = '1'
    LEFT JOIN dim.dim_lookup dl
        ON s.sale_type = dl.cvalue
       AND dl.ctype = 'sale_type_id'
    WHERE asa.online1_complete_dt >= '2026-05-24'
      AND asa.online1_complete_dt <  '2026-06-03'
      AND asae.online1_cheat_proof_flag_name LIKE '%预质检%'
      AND asa.online1_observer_name <> '系统'
      AND asa.online1_observer_name IS NOT NULL
      AND asa.submit_product_category_name = '手机'
      AND asa.online1_remark_txt NOT LIKE '%无理由%'
),

tags_flattened AS (
    SELECT
        online1_complete_date,
        observer_name,
        flag,
        TRIM(tag) AS tag,
        sale_type,
        'prop' AS src
    FROM cte
    LATERAL VIEW EXPLODE(
        SPLIT(COALESCE(oneline1_property_names, ''), CHR(59))
    ) t AS tag

    UNION ALL

    SELECT
        online1_complete_date,
        observer_name,
        flag,
        TRIM(tag) AS tag,
        sale_type,
        'mix' AS src
    FROM cte
    LATERAL VIEW EXPLODE(
        SPLIT(COALESCE(online1_mix_pass_property_name_names, ''), CHR(59))
    ) t AS tag

    UNION ALL

    SELECT
        online1_complete_date,
        observer_name,
        flag,
        TRIM(tag) AS tag,
        sale_type,
        'fail' AS src
    FROM cte
    LATERAL VIEW EXPLODE(
        SPLIT(COALESCE(failed_names, ''), CHR(59))
    ) t AS tag

    UNION ALL

    SELECT
        online1_complete_date,
        observer_name,
        flag,
        TRIM(tag) AS tag,
        sale_type,
        'proof_fail' AS src
    FROM cte
    LATERAL VIEW EXPLODE(
        SPLIT(COALESCE(proof_failed_names, ''), CHR(59))
    ) t AS tag
)

SELECT
    online1_complete_date AS `第一次线审完成时间`,
    observer_name AS `审核人`,
    CASE
        WHEN observer_name IN (
            '李海鹏', '朱熹航', '杨鹏', '乐越辉', '赵敏',
            '徐鑫杰', '周超', '候伟强', '滕济灿'
        ) THEN '常州_老人'
        WHEN observer_name IN (
            '潘姝羽', '李蓓', '朱桂丽'
        ) THEN '常州_新人'
        WHEN observer_name IN (
            '沈轶', '刘付强', '李英蕾', '薛航天', '马梦依'
        ) THEN '上海_第一批'
        WHEN observer_name IN (
            '沈维维', '徐瑞倩', '顾月', '李梦玮', '金晨',
            '范钰硕', '沈丰', '侯帅威', '陶懿酩'
        ) THEN '上海_第二批'
        WHEN observer_name IN (
            '贺然', '黄美林', '张雪', '陈佳蕊', '任敏'
        ) THEN '上海_第三批'
        ELSE '其他'
    END AS `团队`,
    sale_type AS `场次`,
    flag AS `批次`,
    tag AS `属性标签`,
    SUM(CASE WHEN src = 'prop'       THEN 1 ELSE 0 END) AS `申报次数`,
    SUM(CASE WHEN src = 'mix'        THEN 1 ELSE 0 END) AS `模糊通过次数`,
    SUM(CASE WHEN src = 'fail'       THEN 1 ELSE 0 END) AS `未通过次数`,
    SUM(CASE WHEN src = 'proof_fail' THEN 1 ELSE 0 END) AS `举证未通过次数`
FROM tags_flattened
WHERE tag IS NOT NULL
  AND tag <> ''
GROUP BY
    online1_complete_date,
    observer_name,
    CASE
        WHEN observer_name IN (
            '李海鹏', '朱熹航', '杨鹏', '乐越辉', '赵敏',
            '徐鑫杰', '周超', '候伟强', '滕济灿'
        ) THEN '常州_老人'
        WHEN observer_name IN (
            '潘姝羽', '李蓓', '朱桂丽'
        ) THEN '常州_新人'
        WHEN observer_name IN (
            '沈轶', '刘付强', '李英蕾', '薛航天', '马梦依'
        ) THEN '上海_第一批'
        WHEN observer_name IN (
            '沈维维', '徐瑞倩', '顾月', '李梦玮', '金晨',
            '范钰硕', '沈丰', '侯帅威', '陶懿酩'
        ) THEN '上海_第二批'
        WHEN observer_name IN (
            '贺然', '黄美林', '张雪', '陈佳蕊', '任敏'
        ) THEN '上海_第三批'
        ELSE '其他'
    END,
    sale_type,
    flag,
    tag;
