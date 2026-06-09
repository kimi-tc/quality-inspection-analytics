-- 审核人 × 场次质量分析（Hive 精简版）
-- 日期范围：2026-05-24 至 2026-06-02
-- 粒度：日期 × 审核人 × 团队 × 场次

WITH audit_detail AS (
    SELECT
        DATE(asa.online1_complete_dt) AS report_date,
        asa.online1_observer_name AS observer_name,
        COALESCE(dl.cvalue_name, '未知场次') AS sale_type,
        asa.return_order_no,
        asa.online1_result_name,
        asa.online1_remark_txt,
        asa.online1_diff_name_names
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
      AND asa.online1_observer_id <> 0
      AND asa.online1_observer_id IS NOT NULL
      AND asa.online1_observer_name <> '系统'
      AND asa.online1_observer_name IS NOT NULL
      AND asae.online1_cheat_proof_flag_name LIKE '%预质检%'
      AND asa.submit_product_category_name = '手机'
      AND asa.online1_remark_txt NOT LIKE '%无理由%'
),

audit_summary AS (
    SELECT
        report_date,
        observer_name,
        sale_type,

        COUNT(DISTINCT IF(
            online1_result_name IS NOT NULL,
            return_order_no,
            NULL
        )) AS audit_cnt,

        COUNT(DISTINCT IF(
            online1_result_name = '审核通过',
            return_order_no,
            NULL
        )) AS audit_pass_cnt,

        COUNT(DISTINCT IF(
            online1_result_name = '审核未通过',
            return_order_no,
            NULL
        )) AS audit_not_pass_cnt,

        COUNT(DISTINCT IF(
            online1_result_name = '审核未通过'
            AND online1_remark_txt LIKE '%举证%',
            return_order_no,
            NULL
        )) AS proof_refusal_cnt,

        COUNT(DISTINCT IF(
            online1_result_name = '审核通过'
            AND COALESCE(online1_diff_name_names, '') = '',
            return_order_no,
            NULL
        )) AS ambiguous_cnt

    FROM audit_detail
    GROUP BY
        report_date,
        observer_name,
        sale_type
)

SELECT
    report_date AS `日期`,
    observer_name AS `员工姓名`,

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
    CAST(audit_cnt AS BIGINT) AS `一审审核量`,
    CAST(audit_pass_cnt AS BIGINT) AS `一审通过量`,
    CAST(audit_pass_cnt - ambiguous_cnt AS BIGINT) AS `精准通过量`,
    CAST(audit_not_pass_cnt AS BIGINT) AS `未通过量`,
    CAST(proof_refusal_cnt AS BIGINT) AS `举证拒绝量`,
    CAST(ambiguous_cnt AS BIGINT) AS `模糊通过量`,

    CAST(audit_pass_cnt AS DOUBLE)
        / NULLIF(CAST(audit_cnt AS DOUBLE), 0.0) AS `通过率`,

    CAST(audit_pass_cnt - ambiguous_cnt AS DOUBLE)
        / NULLIF(CAST(audit_cnt AS DOUBLE), 0.0) AS `精准通过率`,

    CAST(ambiguous_cnt AS DOUBLE)
        / NULLIF(CAST(audit_cnt AS DOUBLE), 0.0) AS `模棱两可率`,

    CAST(audit_not_pass_cnt AS DOUBLE)
        / NULLIF(CAST(audit_cnt AS DOUBLE), 0.0) AS `拒绝率`,

    CAST(audit_cnt - ambiguous_cnt - proof_refusal_cnt AS DOUBLE)
        / NULLIF(CAST(audit_cnt AS DOUBLE), 0.0) AS `举证准确率`

FROM audit_summary;
