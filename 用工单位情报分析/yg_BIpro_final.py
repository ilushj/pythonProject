import pandas as pd
import mysql.connector  # 假设使用MySQL数据库
from datetime import datetime

# 数据库连接配置（请根据实际情况修改）
db_config = {
    "host": "139.224.44.206",
    "user": "kaimai",
    "password": "kaimai1234",
    "db": "kaimai",
    "port": 3306,
    'charset': 'utf8mb4',

}

# 第一个SQL查询
query1 = """
SELECT
    ep.remark AS 用工单位,
    c.`name` AS 客户名称,
    s.`name` AS 业务员,
    MIN(ep.insustartdate) AS 最早投保日期,
    MAX(ep.insuenddate) AS 最晚投保日期,
    COUNT(DISTINCT claim.id) AS 出险次数,
    COUNT(DISTINCT CASE WHEN claim.CASETYPE = 2 THEN claim.id END) AS 骨折次数,
    SUM(claim.total_paid) / 100 AS 全量实付金额,
    SUM(claim.total_unpaid) / 100 AS 全量预估金额,
    SUM(claim.fracture_paid) / 100 AS 骨折实付金额,
    SUM(claim.fracture_unpaid) / 100 AS 骨折预估金额,
    -- 关键修复：撤案数按案件去重统计
    COUNT(DISTINCT CASE WHEN claim.`status` = 9 THEN claim.id END) AS 撤案数,
    COUNT(DISTINCT CASE WHEN claim.`status` = 9 AND claim.CASETYPE = 2 THEN claim.id END) AS 骨折撤案
FROM (
    SELECT 
        id,
        employee_id,
        paid AS total_paid,
        unpaid AS total_unpaid,
        CASETYPE,
				casedate,
        createdate,
        `status`,  -- 确保外层可访问 status 字段
        CASE WHEN CASETYPE = 8 THEN paid ELSE 0 END AS fracture_paid,
        CASE WHEN CASETYPE = 8 THEN unpaid ELSE 0 END AS fracture_unpaid
    FROM claimcase
) AS claim
LEFT JOIN kaimai.employeeperiod AS ep 
    ON claim.employee_id = ep.employee_id
    AND ep.insuenddate != ep.insustartdate
    AND ep.`status` = 2
    AND ep.remark IS NOT NULL
    AND ep.remark != ''
LEFT JOIN customer AS c 
    ON ep.customer_id = c.id
LEFT JOIN kaimai.customersalesman AS cs 
    ON cs.customer_id = c.id
LEFT JOIN kaimai.salesman AS s 
    ON s.ID = cs.salesman_id
WHERE 
    (
        (ep.insustartdate >= '2022-01-01' AND ep.insustartdate < '2025-03-01')
        AND 
        (claim.casedate >= '2022-01-01')
    )
		AND claim.casedate BETWEEN ep.insustartdate AND ep.insuenddate
    AND NOT EXISTS (
        SELECT 1
        FROM kaimai.customer c2
        WHERE c2.id = ep.customer_id
        AND c2.company_id IN (1, 3, 4, 5, 10, 901)
    )
GROUP BY
    ep.remark
"""

# 第二个SQL查询模板
query2_template = """
SELECT
    COUNT(DISTINCT ep.id) AS 在保人次,
    SUM(CASE WHEN ep.insustartdate <= CURDATE() AND ep.insuenddate >= CURDATE() THEN 1 ELSE 0 END) AS 当前在保人数,
    ep.remark
FROM
    kaimai.employeeperiod AS ep
    LEFT JOIN claimcase AS claim ON claim.employee_id = ep.employee_id
WHERE
    ep.insuenddate != ep.insustartdate
    AND ep.`status` = 2
    AND NOT EXISTS (
        SELECT 1
        FROM kaimai.customer c2
        WHERE c2.id = ep.customer_id
        AND c2.company_id IN (1, 3, 4, 5, 10, 901)
    )
    AND (
        (ep.insustartdate >= '2022-01-01' AND ep.insustartdate < '2025-04-01')
    )
    AND ep.remark IN ({})
GROUP BY
    ep.remark
"""

try:
    # 建立数据库连接
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor()

    # 执行第一个查询
    df1 = pd.read_sql(query1, conn)

    # 获取用工单位列表
    remark_list = df1['用工单位'].unique()
    remark_placeholder = ','.join(['%s'] * len(remark_list))

    # 修改第二个查询，添加用工单位过滤
    query2 = query2_template.format(remark_placeholder)

    # 执行第二个查询
    cursor.execute(query2, tuple(remark_list))
    results2 = cursor.fetchall()

    # 将第二个查询结果转为DataFrame
    df2 = pd.DataFrame(results2, columns=['在保人次', '当前在保人数', '用工单位'])

    # 合并两个DataFrame，以用工单位为键
    merged_df = pd.merge(
        df1,
        df2[['用工单位', '在保人次', '当前在保人数']],
        on='用工单位',
        how='left',
        suffixes=('', '_from_ep')
    )

    # 处理可能出现的重复列（当前在保人数）
    if '当前在保人数_from_ep' in merged_df.columns:
        merged_df['当前在保人数'] = merged_df['当前在保人数'].fillna(merged_df['当前在保人数_from_ep'])
        merged_df = merged_df.drop(columns=['当前在保人数_from_ep'])

    # 调整列顺序（可选）
    column_order = [
        '用工单位', '客户名称', '业务员', '最早投保日期', '最晚投保日期',
        '出险次数', '骨折次数', '全量实付金额', '全量预估金额', '骨折实付金额',
        '骨折预估金额', '撤案数', '骨折撤案', '在保人次', '当前在保人数'
    ]
    merged_df = merged_df[column_order]

    # 导出为CSV
    output_file = f'result_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
    merged_df.to_csv(output_file, index=False, encoding='utf-8-sig')  # utf-8-sig支持中文
    print(f"结果已保存为 {output_file}")

except mysql.connector.Error as e:
    print(f"数据库错误: {e}")
except Exception as e:
    print(f"发生错误: {e}")
finally:
    if 'cursor' in locals():
        cursor.close()
    if 'conn' in locals():
        conn.close()
