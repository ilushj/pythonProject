import pandas as pd
from openpyxl import load_workbook
import os
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta

# 步骤 1: 通过 input 获取文件路径
file_path = input("请输入完整文件路径及文件名（例如 D:\\对账\\数据总表.xlsx）：")

# 检查输入文件是否存在
if not os.path.exists(file_path):
    print(f"错误：输入的文件 {file_path} 不存在，请检查路径和文件名是否正确。")
    exit(1)

# 获取当前日期并计算上个月的目录名称
current_date = datetime.now()
last_month = current_date - relativedelta(months=1)
directory_name = last_month.strftime("%Y%m")  # 格式为 YYYYMM，例如 202502

# 构建新目录路径和输出文件路径
output_directory = r"D:\对账" + "\\" + directory_name
output_file_path = os.path.join(output_directory, os.path.basename(file_path))
baosi_guazhang_path = os.path.join(output_directory, "保司挂账.xlsx")

# 确保输出目录存在，如果不存在则创建
try:
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
except PermissionError:
    print(f"错误：无法创建目录 {output_directory}，请检查权限或以管理员身份运行程序。")
    exit(1)

# 如果输出文件不存在，复制输入文件到输出路径
if not os.path.exists(output_file_path):
    try:
        shutil.copyfile(file_path, output_file_path)
    except Exception as e:
        print(f"错误：无法复制文件到 {output_file_path}，原因：{str(e)}")
        exit(1)

# 读取原始数据
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"错误：无法读取文件 {file_path}，原因：{str(e)}")
    exit(1)

# 加载输出文件（现在已存在）
try:
    book = load_workbook(output_file_path)
except Exception as e:
    print(f"错误：无法加载文件 {output_file_path}，原因：{str(e)}")
    exit(1)

# 创建 ExcelWriter 对象，并绑定现有的 workbook
try:
    writer = pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
    writer.workbook = book
except Exception as e:
    print(f"错误：无法写入文件 {output_file_path}，原因：{str(e)}")
    exit(1)

# 步骤 2: 创建“全量保司结算”Sheet
df_settlement = df[df['结算类型'] == '保司结算']
df_summary = df_settlement.groupby(['发票抬头', '项目类型', '开票类型']).agg({
    '总保费': 'sum',
    '总成本': 'sum'
}).reset_index()

df_summary = df_summary.merge(df_settlement[['发票抬头', '项目类型', '客户名称', '税号']].drop_duplicates(),
                              on=['发票抬头', '项目类型'], how='left')

df_summary = df_summary[['发票抬头', '税号', '客户名称', '项目类型', '开票类型', '总保费', '总成本']]
df_summary.to_excel(writer, sheet_name='全量保司结算', index=False)

# 步骤 3: 创建“比对结果”Sheet
df_full_settlement = df_summary
df_null_invoice = df_full_settlement[df_full_settlement['发票抬头'].isna()]
df_not_null_invoice = df_full_settlement[df_full_settlement['发票抬头'].notna()]
duplicated_invoices = df_not_null_invoice['发票抬头'].duplicated(keep=False)
df_not_null_filtered = df_not_null_invoice[
    (~duplicated_invoices) |
    (duplicated_invoices & (df_not_null_invoice['项目类型'] != '易康项目'))
]
df_compare = pd.concat([df_null_invoice, df_not_null_filtered], ignore_index=True)
df_compare.to_excel(writer, sheet_name='比对结果', index=False)

# 步骤 4: 按“项目类型”拆分Sheet
project_types = df_compare['项目类型'].unique()
for project_type in project_types:
    if pd.notna(project_type):
        df_project = df_compare[df_compare['项目类型'] == project_type]
        df_project.to_excel(writer, sheet_name=project_type, index=False)

# 步骤 5: 创建“总账”Sheet，无数据
pd.DataFrame().to_excel(writer, sheet_name='总账', index=False)

# 步骤 6: 创建“保司挂账”Sheet，无数据
pd.DataFrame().to_excel(writer, sheet_name='保司挂账', index=False)

# 保存并关闭主文件
writer.close()

# 步骤 7: 提取“全量保司结算”中的“发票抬头”和“税号”，保存到新文件
df_baosi_guazhang = df_summary[['发票抬头', '税号']].drop_duplicates()
try:
    with pd.ExcelWriter(baosi_guazhang_path, engine='openpyxl') as writer:
        df_baosi_guazhang.to_excel(writer, sheet_name='Sheet1', index=False)
except Exception as e:
    print(f"错误：无法写入文件 {baosi_guazhang_path}，原因：{str(e)}")
    exit(1)

print(f"处理完成，文件已保存至：{output_file_path}")
print(f"保司挂账数据已保存至：{baosi_guazhang_path}")