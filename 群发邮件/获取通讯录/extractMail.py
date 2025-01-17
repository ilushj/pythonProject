import pandas as pd

# 读取Excel文件
df = pd.read_excel('奇点保险公司通讯录.xlsx', sheet_name=None)

# 创建一个空的DataFrame来存储提取的数据
result_df = pd.DataFrame()

# 遍历每个sheet
for sheet_name, sheet_df in df.items():
    # 提取姓名和邮箱地址都有值的行
    sheet_df = sheet_df[(sheet_df['姓名'].notna()) & (sheet_df['邮箱'].notna())]

    # 提取姓名和邮箱列
    sheet_df = sheet_df[['姓名', '邮箱']]

    # 添加到result_df
    result_df = pd.concat([result_df, sheet_df], ignore_index=True)

# 去除重复行
result_df = result_df.drop_duplicates()

# 保存到新Excel文件
result_df.to_excel('公司通讯录.xlsx', index=False)