import pandas as pd
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 读取Excel文件
file_path = r'D:\数据总表\1.xlsx'
df = pd.read_excel(file_path)

# 检查数据类型
print(df.dtypes)

# 数据预览
print(df.head())

# 清洗“客户赔付率”字段
# 假设“客户赔付率”字段包含百分号,需要移除百分号并转换为浮点数
if '客户赔付率' in df.columns:
    df['客户赔付率'] = df['客户赔付率'].str.replace('%', '', regex=False).astype(float) / 100
else:
    print("警告: '客户赔付率' 字段不存在于数据中。")

# 定义需要合并的字段
sum_fields = ['总保费', '已结保费', '未结保费', '投保人数', '预估赔付', '实际赔付', '综合赔付']
mean_fields = ['客户赔付率']

# 按“客户名称”分组并进行合并操作
grouped = df.groupby('客户名称', as_index=False).agg({
    **{field: 'sum' for field in sum_fields},
    **{field: 'mean' for field in mean_fields}
})

# 保留原始数据中的第一行记录
first_rows = df.drop_duplicates(subset='客户名称', keep='first')

# 更新第一行记录的合并数据
merged_df = pd.merge(first_rows, grouped, on='客户名称', suffixes=('', '_merged'))

# 更新原始数据中的字段
for field in sum_fields + mean_fields:
    merged_df[field] = merged_df[f'{field}_merged']
    del merged_df[f'{field}_merged']

# 保存结果到新的Excel文件
output_file_path = r'D:\数据总表\merged_1.xlsx'
merged_df.to_excel(output_file_path, index=False)

print(f'数据已成功合并并保存为 {output_file_path}')
