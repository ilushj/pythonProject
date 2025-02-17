import mysql.connector
import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

# 从用户输入获取目录路径
from_date = input("请输入开始时间例如2024-01-01: ").strip()
to_date = from_date

# 数据库配置（需要修改为实际配置）
db_config = {
    "host": "139.224.44.206",
    "user": "kaimai",
    "password": "kaimai1234",
    "database": "kaimai",
    "port": 3306  # 数据库端口号，默认是 3306
}

# 时间参数设置
# from_date = '2024-01-01'
# to_date = '2024-12-31'

try:
    # 连接数据库
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor()

    # 设置用户变量
    cursor.execute("SET @from = %s;", (from_date,))
    cursor.execute("SET @to = %s;", (to_date,))

    # 执行查询
    query = """
    SELECT
        CAST(SUM(importCount) AS UNSIGNED) AS 人数,
        sum(premium / (DATEDIFF(t.enddate, t.STARTDATE) + 1) * (datediff(greatest(t.enddate, @to), least(t.STARTDATE, @from)) + 1)) as premium,
        sm.`name` AS 业务员姓名
    FROM
        kaimai.insupolicysalesman AS t
        INNER JOIN kaimai.customer AS c ON t.customer_id = c.id AND c.company_id != 1 
        INNER JOIN kaimai.salesman AS sm ON t.salesman_id = sm.id AND sm.`name` NOT LIKE '%/%'
    WHERE
        t.startdate <= @from OR t.enddate >= @to 
    GROUP BY
        sm.`name`
    ORDER BY
        人数 DESC;
    """

    cursor.execute(query)
    result = cursor.fetchall()
    columns = [i[0] for i in cursor.description]

    # 转换为DataFrame
    df = pd.DataFrame(result, columns=columns)

    # 生成Excel文件
    excel_file = f"{from_date}业务人数竞赛.xlsx"
    df.to_excel(excel_file, index=False, engine='openpyxl')

    # 添加图表
    wb = load_workbook(excel_file)
    ws = wb.active

    # **创建用于图表的 DataFrame (只包含前 10 行)**
    df_top10 = df.head(10)

    # 创建柱状图
    chart = BarChart()
    chart.title = "业务员排名"
    chart.y_axis.title = "人数"
    chart.x_axis.title = "业务员"
    chart.style = 13  # 使用预定义样式

    # 数据范围（假设数据从第2行开始）
    data = Reference(ws, min_col=1, min_row=1, max_row=len(df_top10) + 1, max_col=1)
    # data = Reference(ws, min_col=2, min_row=1, max_row=len(df) + 1, max_col=2)
    categories = Reference(ws, min_col=3, min_row=2, max_row=len(df_top10) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # 调整图表大小和位置
    chart.width = 20
    chart.height = 12
    ws.add_chart(chart, "E2")

    # 保存Excel文件
    wb.save(excel_file)
    print(f"文件已生成：{excel_file}")

except mysql.connector.Error as err:
    print(f"数据库错误：{err}")
finally:
    if 'conn' in locals() and conn.is_connected():
        cursor.close()
        conn.close()


