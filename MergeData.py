import os
import pandas as pd
from datetime import datetime

# 获取当前日期
today = datetime.today().strftime('%Y-%m-%d')

# 定义目录
directory = r'D:\千服日报'
today_directory = os.path.join(directory, today)

# 如果目录不存在，则创建目录
if not os.path.exists(today_directory):
    os.makedirs(today_directory)

# 初始化数据帧列表
batch_add = []
batch_subtract = []

# 遍历目录中的所有xlsx文件
for filename in os.listdir(directory):
    if filename.endswith('.xlsx') and '最新雇员清单' not in filename:
        file_path = os.path.join(directory, filename)
        # 读取文件，跳过第一行
        df = pd.read_excel(file_path, dtype={'新雇员证件号码': str})
        if '批增' in filename:
            batch_add.append(df)
        elif '批减' in filename:
            batch_subtract.append(df)

# 合并数据帧
if batch_add:
    df_batch_add = pd.concat(batch_add, ignore_index=True)
else:
    df_batch_add = pd.DataFrame()

if batch_subtract:
    df_batch_subtract = pd.concat(batch_subtract, ignore_index=True)
else:
    df_batch_subtract = pd.DataFrame()

# 找到重复数据
if not df_batch_add.empty and not df_batch_subtract.empty:
    # 使用“新雇员证件号码”和“投保方案”作为键进行合并
    merged_df = pd.merge(df_batch_add, df_batch_subtract, on=['新雇员证件号码', '投保方案'], how='inner', suffixes=('_add', '_subtract'))
    # 从df_batch_subtract中删除重复数据
    df_batch_subtract = df_batch_subtract[~df_batch_subtract[['新雇员证件号码', '投保方案']].isin(merged_df[['新雇员证件号码', '投保方案']]).all(axis=1)]
else:
    merged_df = pd.DataFrame()

# 创建一个新的Excel writer
output_file = os.path.join(today_directory, f'{today}.xlsx')
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_batch_add.to_excel(writer, sheet_name=f'{today}批增', index=False)
    df_batch_subtract.to_excel(writer, sheet_name=f'{today}批减', index=False)
    merged_df.to_excel(writer, sheet_name='重复数据', index=False)

print(f'数据已成功合并并保存为 {output_file}')
