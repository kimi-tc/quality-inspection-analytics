-- 审核人在场次维度的审核质量
-- 统计周期：2026-05-24 至 2026-06-02（包含首尾日期）
-- 粒度：日期 × 审核人 × 团队 × 场次
-- 引擎：Hive
--
-- 口径说明：
-- 1. 仅统计第一次在线人工审核，质量指标与当前人效 SQL 保持订单级口径。
-- 2. 场次沿用质量 SQL：结算单商品销售类型 -> dim_lookup.sale_type_id。
-- 3. 举证准确率 =（一审审核量 - 模糊通过量 - 举证拒绝量）/ 一审审核量。
-- 4. 精准通过率 =（一审通过量 - 模糊通过量）/ 一审审核量。
-- 5. 使用半开日期区间，确保完整包含 2026-06-02 当天。

WITH sale_type_map AS (
    SELECT DISTINCT
        s.settle_document_no,
        s.product_no,
        COALESCE(dl.cvalue_name, '未知场次') AS sale_type
    FROM dw.dw_platform_settle_document_product s
    LEFT JOIN dim.dim_lookup dl
        ON s.sale_type = dl.cvalue
       AND dl.ctype = 'sale_type_id'
    WHERE s.partition_flag = '1'
),

audit_base AS (
    SELECT
        DATE(asa.online1_complete_dt) AS report_date,
        asa.online1_observer_name AS observer_name,
        COALESCE(stm.sale_type, '未知场次') AS sale_type,
        asa.return_order_no,
        asa.online1_result_name,
        asa.online1_remark_txt,
        asa.online1_diff_name_names,
        asae.oneline1_property_names,
        asae.online1_mix_pass_property_name_names
    FROM dm.dm_centre_return_after_sale_analysis asa
    JOIN dm.dm_centre_return_after_sale_analysis_ext asae
        ON asa.return_order_no = asae.return_order_no
       AND asa.product_no = asae.product_no
    LEFT JOIN sale_type_map stm
        ON asa.settle_document_no = stm.settle_document_no
       AND asa.product_no = stm.product_no
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

auditor_session_daily AS (
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

        -- 与当前人效 SQL 保持一致：审核通过且 diff 为空，记为模糊通过。
        COUNT(DISTINCT IF(
            online1_result_name = '审核通过'
            AND COALESCE(online1_diff_name_names, '') = '',
            return_order_no,
            NULL
        )) AS ambiguous_cnt,

        SUM(
            CASE
                WHEN online1_result_name IS NOT NULL
                 AND online1_diff_name_names = oneline1_property_names
                 AND INSTR(oneline1_property_names, CHR(59)) = 0
                 AND oneline1_property_names IN (
                    '售后案例情况',
                    '主板维修情况',
                    '屏幕维修情况',
                    '电池维修情况',
                    '后壳维修情况',
                    '零件维修情况',
                    '摄像头维修',
                    '受潮状况',
                    '整机风险项'
                 )
                THEN 1.1

                WHEN online1_result_name IS NOT NULL
                 AND online1_diff_name_names = oneline1_property_names
                 AND INSTR(oneline1_property_names, CHR(59)) = 0
                 AND oneline1_property_names IN (
                    '屏幕外观',
                    '边框背板',
                    '机身弯曲'
                 )
                THEN 0.8

                WHEN online1_result_name IS NOT NULL
                 AND (
                    INSTR(oneline1_property_names, CHR(59)) > 0
                    OR INSTR(online1_diff_name_names, CHR(59)) > 0
                    OR online1_diff_name_names <> oneline1_property_names
                 )
                THEN 1.2

                WHEN online1_result_name IS NOT NULL
                THEN 1

                ELSE 0
            END
        ) AS weighted_audit_cnt
    FROM audit_base
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
    CAST(weighted_audit_cnt AS DOUBLE) AS `加权审核量`,
    CAST(audit_pass_cnt AS BIGINT) AS `一审通过量`,
    CAST(audit_pass_cnt - ambiguous_cnt AS BIGINT) AS `精准通过量`,
    CAST(audit_not_pass_cnt AS BIGINT) AS `未通过量`,
    CAST(proof_refusal_cnt AS BIGINT) AS `举证拒绝量`,
    CAST(ambiguous_cnt AS BIGINT) AS `模糊通过量`,

    CAST(
        audit_pass_cnt / NULLIF(audit_cnt, 0)
        AS DOUBLE
    ) AS `通过率`,

    CAST(
        (audit_pass_cnt - ambiguous_cnt) / NULLIF(audit_cnt, 0)
        AS DOUBLE
    ) AS `精准通过率`,

    CAST(
        ambiguous_cnt / NULLIF(audit_cnt, 0)
        AS DOUBLE
    ) AS `模棱两可率`,

    CAST(
        audit_not_pass_cnt / NULLIF(audit_cnt, 0)
        AS DOUBLE
    ) AS `拒绝率`,

    CAST(
        1 - ((ambiguous_cnt + proof_refusal_cnt) / NULLIF(audit_cnt, 0))
        AS DOUBLE
    ) AS `举证准确率`

FROM auditor_session_daily
ORDER BY
    `日期`,
    `团队`,
    `员工姓名`,
    `场次`;
