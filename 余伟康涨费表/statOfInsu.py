import pymysql
import pandas as pd
from datetime import datetime
import os


# 数据库连接函数
def get_db_connection():
    return pymysql.connect(
        host='139.224.44.206',  # 替换为你的数据库主机
        user='kaimai',  # 替换为你的数据库用户名
        password='kaimai1234',  # 替换为你的数据库密码
        database='kaimai',  # 替换为你的数据库名称
        charset='utf8mb4'
    )


# 第一步：查询价格变化的客户
def query_price_changes():
    start_date = input("请输入开始时间：")  # 输入日期的当月第一天
    end_date = input("请输入结束时间：")  # 次年同月第一天

    query = """
    SELECT
        c.`name` AS 客户名称,
        p.productname AS 方案名称,
        GROUP_CONCAT(DISTINCT pp.PRICE ORDER BY pp.PRICE SEPARATOR ', ') AS 不同价格,
        pp.`name` AS 价格名称,
        '有变化' AS 价格变化 
    FROM
        policy AS p
        INNER JOIN policyprice AS pp ON pp.policy_id = p.ID
        INNER JOIN customer AS c ON p.Customer_id = c.ID
        INNER JOIN insupolicy AS insup ON insup.policy_id = p.ID 
    WHERE
        (insup.startdate >= %s and insup.startdate < %s)
        AND c.company_id NOT IN (1, 3, 4, 5, 10, 901)
    GROUP BY
        c.`name`,
        p.productname,
        pp.`name`
    HAVING
        COUNT(DISTINCT pp.PRICE) > 1 
        AND (NOT (不同价格 LIKE '0,%%' AND CHAR_LENGTH(不同价格) - CHAR_LENGTH(REPLACE(不同价格, ',', '')) = 1));
    """

    conn = get_db_connection()
    try:
        df = pd.read_sql(query, conn, params=(start_date, end_date))
        return df
    finally:
        conn.close()


# 第二步：根据第一步的客户名称查询在保情况
def query_insurance_status(customer_names):
    current_month_start = datetime.now().strftime('%Y-%m-01')  # 当前月第一天

    query = """
    SELECT
        insup.customer_id,
        c.`name` AS 客户名称,
        p.productname,
        '当前在保' AS 在保情况 
    FROM
        insupolicy AS insup
        INNER JOIN customer AS c ON insup.Customer_id = c.ID
        INNER JOIN policy AS p ON p.customer_id = insup.customer_id
    WHERE
        insup.startdate >= %s
        AND c.company_id NOT IN (1, 3, 4, 5, 10, 901)    
    """

    # 将 customer_names 转为逗号分隔的字符串，用于 IN 子句

    conn = get_db_connection()
    try:
        # 参数列表包含current_month_start和所有customer_names
        df = pd.read_sql(query, conn, params=(current_month_start,))
        return df
    finally:
        conn.close()


# 第三步：合并结果并去重
def merge_results(df_price_changes, df_insurance_status):
    if df_price_changes.empty:
        return pd.DataFrame()

    # 以客户名称为键合并两个 DataFrame
    merged_df = pd.merge(
        df_price_changes,
        df_insurance_status[['客户名称', 'productname', '在保情况']],
        on='客户名称',
        how='outer'
    )

    print("合并后的列名:", merged_df.columns.tolist())  # 调试用

    # 处理 productname，使用在保记录的 productname，否则用原有的方案名称
    merged_df['方案名称'] = merged_df.apply(
        lambda row: row['productname'] if pd.notna(row['productname']) else row['方案名称'],
        axis=1
    )

    # 删除多余的 productname 列
    merged_df = merged_df.drop(columns=['productname'])
    merged_df['在保情况'] = merged_df['在保情况'].fillna('不在保')
    merged_df['价格变化'] = merged_df['价格变化'].fillna('无变化')
    merged_df[['不同价格', '价格名称']] = merged_df[['不同价格', '价格名称']].fillna('')

    # 按“价格变化”字段从高到低排序（有变化在前，无变化在后）
    merged_df = merged_df.sort_values(by='价格变化', ascending=False)
    # 去重
    # merged_df = merged_df.drop_duplicates(subset=['客户名称', '方案名称'])

    return merged_df

# 主函数
def main():
    # 第一步：查询价格变化
    df_price_changes = query_price_changes()
    if df_price_changes.empty:
        print("没有找到价格变化的记录")
        return

    # 获取客户名称列表
    customer_names = df_price_changes['客户名称'].tolist()

    # 第二步：查询在保情况
    df_insurance_status = query_insurance_status(customer_names)

    # 第三步：合并结果
    result_df = merge_results(df_price_changes, df_insurance_status)

    # 输出结果
    print(result_df)
    # 可选：保存到文件


    result_df.to_csv('result.csv', index=False, encoding='utf-8-sig')

    # 将结果写入现有 Excel 文件，新增“价格变化” Sheet
    file_path = r'D:\数据总表\merged_1.xlsx'
    sheet_name = '价格变化'

    try:
        # 如果文件已存在，追加新 Sheet
        if os.path.exists(file_path):
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                result_df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # 如果文件不存在，创建新文件并写入
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"结果已成功写入 {file_path} 的 '{sheet_name}' Sheet")
    except Exception as e:
        print(f"写入 Excel 文件失败: {e}")


# 示例运行
if __name__ == "__main__":
    main()
