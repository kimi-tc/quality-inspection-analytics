-- 审核人场次质量绝对值（多属性项不炸开）
-- 日期范围：2026-05-24 至 2026-06-02
-- 统计单位：原始审核记录，一条记录最多计一次申报
-- 输出粒度：日期 × 审核人 × 场次

WITH source_data AS (
    SELECT
        DATE(asa.online1_complete_dt) AS online1_complete_date,
        asa.online1_observer_name AS observer_name,
        dl.cvalue_name AS sale_type,
        asae.oneline1_property_names,
        asae.online1_mix_pass_property_name_names,
        asa.online1_result_name,
        asa.online1_remark_txt
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
)

SELECT
    online1_complete_date AS `第一次线审完成时间`,
    observer_name AS `审核人`,
    sale_type AS `场次`,
    SUM(
        CASE
            WHEN COALESCE(TRIM(oneline1_property_names), '') <> ''
            THEN 1 ELSE 0
        END
    ) AS `申报次数`,
    SUM(
        CASE
            WHEN COALESCE(TRIM(oneline1_property_names), '') <> ''
             AND COALESCE(TRIM(online1_mix_pass_property_name_names), '') <> ''
            THEN 1 ELSE 0
        END
    ) AS `模糊通过次数`,
    SUM(
        CASE
            WHEN COALESCE(TRIM(oneline1_property_names), '') <> ''
             AND online1_result_name LIKE '%未通过%'
            THEN 1 ELSE 0
        END
    ) AS `未通过次数`,
    SUM(
        CASE
            WHEN COALESCE(TRIM(oneline1_property_names), '') <> ''
             AND online1_remark_txt LIKE '%举证%'
             AND online1_result_name LIKE '%未通过%'
            THEN 1 ELSE 0
        END
    ) AS `举证未通过次数`
FROM source_data
GROUP BY
    online1_complete_date,
    observer_name,
    sale_type;
