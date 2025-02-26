import pymysql
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

matplotlib.use('Agg')  # 设置非交互式后端
# 数据库连接配置
db_config = {
    "host": "139.224.44.206",
    "user": "kaimai",
    "password": "kaimai1234",
    "db": "kaimai",
    "port": 3306,
    'charset': 'utf8mb4',
    'cursorclass': pymysql.cursors.DictCursor
}

# 建立数据库连接
connection = pymysql.connect(**db_config)

# 定义日期范围
start_date = '2022-01-01'
end_date = '2025-03-01'

# SQL 查询
sql = """
SELECT
    ep.remark AS 用工单位,
	c.`name` AS 客户名称,
	s.`name` AS 业务员,
    MIN(ep.insustartdate) AS 最早投保日期,
    MAX(ep.insuenddate) AS 最晚投保日期,
    COUNT(claim.id) AS 出险次数,
	SUM(CASE WHEN claim.CASETYPE = 8 then 1 ELSE 0 END) AS 骨折次数,
    SUM(claim.paid) / 100 AS 全量实付金额,
    SUM(claim.unpaid) / 100 AS 全量预估金额,
	SUM(CASE WHEN claim.CASETYPE = 8 THEN claim.paid ELSE 0 END) / 100 AS 骨折实付金额,
    SUM(CASE WHEN claim.CASETYPE = 8 THEN claim.unpaid ELSE 0 END) / 100 AS 骨折预估金额,
	SUM(CASE WHEN claim.`status` = 9 THEN 1 ELSE 0 END) AS 撤案数,
	SUM(CASE WHEN claim.`status` = 9 and claim.CASETYPE = 8 THEN 1 ELSE 0 END) AS 骨折撤案,
 FROM
    claimcase AS claim
    LEFT JOIN kaimai.employeeperiod AS ep ON claim.employee_id = ep.employee_id
    LEFT JOIN customer AS c ON ep.customer_id = c.id
	LEFT JOIN kaimai.customersalesman AS cs ON cs.customer_id = c.id
    LEFT JOIN kaimai.salesman AS s ON s.ID = cs.salesman_id
WHERE
    (ep.insustartdate >= %s AND ep.insustartdate < %s)
    AND ep.insuenddate != ep.insustartdate
    AND ep.`status` = 2
    AND NOT EXISTS (
        SELECT 1
        FROM kaimai.customer c2
        WHERE c2.id = ep.customer_id
        AND c2.company_id IN (1, 3, 4, 5, 10, 901)
    )
    AND (claim.createdate >= %s AND claim.createdate < %s)
GROUP BY
    ep.remark,
    s.`name`,
    c.`name`;
"""

# 执行查询并调试
try:
    with connection.cursor() as cursor:
        cursor.execute(sql, (start_date, end_date, start_date, end_date))
        results = cursor.fetchall()
        print("查询结果行数:", len(results))
        if not results:
            raise ValueError("SQL 查询未返回任何数据")
finally:
    connection.close()

# 转换为 DataFrame
df = pd.DataFrame(results)
print("DataFrame 列名:", df.columns.tolist())
print("全量实付金额 示例数据:", df['全量实付金额'].head())

# 数据清洗
try:
    df['全量实付金额'] = pd.to_numeric(df['全量实付金额'], errors='coerce').fillna(0)
    df['出险次数'] = pd.to_numeric(df['出险次数'], errors='coerce').fillna(0)
    df['骨折实付金额'] = pd.to_numeric(df['骨折实付金额'], errors='coerce').fillna(0)
    df['骨折次数'] = pd.to_numeric(df['骨折次数'], errors='coerce').fillna(0)
    df['在保人次'] = pd.to_numeric(df['在保人次'], errors='coerce').fillna(0)
    df['撤案数'] = pd.to_numeric(df['撤案数'], errors='coerce').fillna(0)
    df['骨折撤案'] = pd.to_numeric(df['骨折撤案'], errors='coerce').fillna(0)
except KeyError as e:
    print(f"KeyError: 列 {e} 不存在，请检查 SQL 查询")
except Exception as e:
    print(f"数据清洗出错: {e}")

# 计算风险指标
df['出险率'] = df['出险次数'] / df['在保人次']
df['骨折出险率'] = df['骨折次数'] / df['在保人次']
df['平均实付金额'] = df['全量实付金额'] / df['出险次数']
df['平均实付金额'] = df['平均实付金额'].where(df['出险次数'] > 0, 0)
df['平均骨折实付金额'] = df['骨折实付金额'] / df['骨折次数']
df['平均骨折实付金额'] = df['平均骨折实付金额'].where(df['骨折次数'] > 0, 0)
df['撤案率'] = df['撤案数'] / df['出险次数']
df['撤案率'] = df['撤案率'].where(df['出险次数'] > 0, 0)
df['骨折撤案率'] = df['骨折撤案'] / df['骨折次数']
df['骨折撤案率'] = df['骨折撤案率'].where(df['骨折次数'] > 0, 0)

# 替换 NaN 为 0
df.fillna(0, inplace=True)

# 清理用工单位中的制表符
df['用工单位'] = df['用工单位'].str.replace('\t', '', regex=False)

# 打印前5个高风险用工单位
high_risk_units = df.sort_values(by='出险率', ascending=False)
print(high_risk_units[['用工单位', '出险率', '骨折出险率', '平均实付金额', '撤案率']].head())

# 保存结果到 CSV
df.to_csv('risk_analysis.csv', index=False, encoding='utf-8-sig')

# 可视化：显示前10高风险用工单位
plt.rcParams['font.sans-serif'] = ['SimHei']  # 支持中文显示
plt.rcParams['axes.unicode_minus'] = False    # 解决负号显示问题
top_n = 10
df_top = df.sort_values(by='出险率', ascending=False).head(top_n)
plt.figure(figsize=(10, 8))
sns.barplot(x='出险率', y='用工单位', hue='用工单位', data=df_top, palette='coolwarm_r', legend=False)
plt.title(f'前 {top_n} 高风险用工单位出险率')
plt.xlabel('出险率')
plt.ylabel('用工单位')
plt.tight_layout()
plt.savefig('risk_analysis.png', dpi=300, bbox_inches='tight')  # 保存图像

# 示例：获取 RGB 图像数据
fig, ax = plt.subplots()
ax.plot([1, 2, 3], [4, 5, 6])
canvas = fig.canvas
canvas.draw()  # 显式渲染画布
argb_data = canvas.tostring_argb()  # 使用 tostring_argb
width, height = canvas.get_width_height()
img = np.frombuffer(argb_data, dtype=np.uint8).reshape(height, width, 4)
rgb_img = img[:, :, 1:4]  # 从 ARGB 提取 RGB
print("RGB 图像数据形状:", rgb_img.shape)
plt.savefig('example_plot.png', dpi=300)  # 保存示例图像