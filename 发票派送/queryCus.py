import mysql.connector
import pandas as pd

# 输入文件路径
file_path = input("请输入文件保存路径：")
# 文件名
file_name = "公司所属.xlsx"

# 完整文件路径
full_file_path = file_path + "/" + file_name

# 连接数据库
cnx = mysql.connector.connect(
    user='kaimai',
    password='kaimai1234',
    host='139.224.44.206',  # 例如 'localhost' 或 '192.168.1.100'
    port=3306,  # MySQL 默认端口号
    database='kaimai'
)

# 创建游标
cursor = cnx.cursor()

# SQL 查询
query = """  
    SELECT DISTINCT  
        s.`name` AS 业务员,  
        COALESCE(cmi.invoicename, c.invoicename) AS 客户发票抬头,  
        c.NAME AS 客户名称  
    FROM  
        `customer` AS c  
        LEFT JOIN kaimai.customermultiinsured AS cmi ON c.multiInsuredCustomers LIKE CONCAT( '%', cmi.NAME, '%' )  
        INNER JOIN kaimai.customersalesman AS cs ON cs.customer_id = c.id  
        INNER JOIN kaimai.salesman AS s ON s.ID = cs.salesman_id  
        INNER JOIN kaimai.company AS comp ON s.company_id = comp.id  
    WHERE  
        c.insurerSettlement = 1  
        AND comp.id NOT IN ( 1, 3, 4, 5, 10, 901 )  
        AND s.`status` = 0  
        AND c.NAME <> ''  
        AND NOT ( s.NAME = '方志伟' AND c.NAME = '上海博语稻企业管理集团有限公司' )  
        AND NOT ( c.invoicename = '' AND cmi.invoicename = '' )  
    ORDER BY  
        COALESCE(cmi.invoicename, c.invoicename)  
"""

# 执行查询
cursor.execute(query)

# 获取结果
result = cursor.fetchall()

# 将结果转换为 DataFrame
df = pd.DataFrame(result, columns=['业务员', '客户发票抬头', '客户名称'])

# 保存结果到 Excel 文件
df.to_excel(full_file_path, index=False)


# 关闭游标和连接
cursor.close()
cnx.close()
