import os
import pandas as pd
from datetime import datetime
import warnings

# 忽略UserWarning警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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
latest_employee_list = pd.DataFrame()

# 遍历目录中的所有xlsx文件，排除文件名包含“最新雇员清单”的文件
for filename in os.listdir(directory):
    file_path = os.path.join(directory, filename)
    if filename.endswith('.xlsx'):
        if '最新雇员清单' in filename:
            # 读取最新雇员清单
            latest_employee_list = pd.read_excel(file_path, dtype={'证件号码': str})
        elif '最新雇员清单' not in filename:
            # 读取文件，跳过第一行，并将“新雇员证件号码”列设为文本
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

# 修改批增数据中的特定列
if not df_batch_add.empty:
    # 条件一：投保方案等于“雇主责任险（1-4类）30万-17”
    mask1 = df_batch_add['投保方案'] == '雇主责任险（1-4类）50万-17'
    df_batch_add.loc[mask1, '新职业类别'] = '四类职业'
    df_batch_add.loc[mask1, '新岗位名称'] = '普工2'

    # 条件二：投保方案等于“雇主责任险（1-4类）80万-30”
    mask2 = df_batch_add['投保方案'] == '雇主责任险（1-4类）80万-30'
    df_batch_add.loc[mask2, '新职业类别'] = '四类职业'
    df_batch_add.loc[mask2, '新岗位名称'] = '普工2'

    # 条件三：投保方案等于“G_雇主责任险（1-3类）80+20万-36”
    mask3 = df_batch_add['投保方案'] == 'G_雇主责任险（1-3类）80+20万-36'
    df_batch_add.loc[mask3, '新职业类别'] = '三类职业'
    df_batch_add.loc[mask3, '新岗位名称'] = '邮件分拣员-邮件分拣员-2'

    # 条件四：投保方案等于“H_雇主责任险（1-3类）100万-52”
    mask4 = df_batch_add['投保方案'] == 'H_雇主责任险（1-3类）100万-52'
    df_batch_add.loc[mask4, '新职业类别'] = '三类职业'
    df_batch_add.loc[mask4, '新岗位名称'] = '市场管理员'

    # 条件五：投保方案等于“雇主责任险（1-4类）30万-12”
    mask5 = df_batch_add['投保方案'] == '雇主责任险（1-4类）30万-12'
    df_batch_add.loc[mask5, '新职业类别'] = '四类职业'
    df_batch_add.loc[mask5, '新岗位名称'] = '普工1'

# 找到重复数据
if not df_batch_add.empty and not df_batch_subtract.empty:
    # 使用“新雇员证件号码”和“投保方案”作为键进行合并
    merged_df = pd.merge(df_batch_add, df_batch_subtract, on=['新雇员证件号码', '投保方案'], how='inner',
                         suffixes=('_add', '_subtract'))

    # 标记在 df_batch_subtract 中与 merged_df 中相同的行
    merged_flag = pd.merge(df_batch_subtract, merged_df[['新雇员证件号码']], on='新雇员证件号码', how='left',
                           indicator=True)

    # 从 df_batch_subtract 中删除与 merged_df 中相同的行
    df_batch_subtract = merged_flag[merged_flag['_merge'] != 'both'].drop(columns='_merge')
else:
    merged_df = pd.DataFrame()
# 比对和处理最新雇员清单
replacement_data = []

if not latest_employee_list.empty:
    for index, row in df_batch_add.iterrows():
        employee_id = row['新雇员证件号码']
        job_name = row['新岗位名称']

        match = latest_employee_list[(latest_employee_list['证件号码'] == employee_id)]

        if not match.empty:
            if (match['岗位名称'] == job_name).any():
                df_batch_add.drop(index, inplace=True)
            else:
                combined_row = pd.concat([match.iloc[0], row], axis=0)
                replacement_data.append(combined_row)
                row_dropped = row.drop(labels=['投保方案'])
                df_batch_subtract = pd.concat([df_batch_subtract, row_dropped.to_frame().T], ignore_index=True)

replacement_df = pd.DataFrame(replacement_data)

# 创建一个新的Excel writer
output_file = os.path.join(today_directory, f'{today}.xlsx')
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_batch_add.to_excel(writer, sheet_name='批增', index=False)
    df_batch_subtract.to_excel(writer, sheet_name='批减', index=False)
    merged_df.to_excel(writer, sheet_name='重复数据', index=False)
    if not replacement_df.empty:
        replacement_df.to_excel(writer, sheet_name='替换数据', index=False)

print(f'数据已成功合并并保存为 {output_file}')
