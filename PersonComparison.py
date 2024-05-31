import pandas as pd
from datetime import datetime, timedelta
import os
import warnings

# 忽略特定警告
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")

# 获取今天和昨天的日期
today = datetime.today()
yesterday = today - timedelta(days=1)

# 构造文件名
filename_A = os.path.join('D:\\数据比对', yesterday.strftime('%Y%m%d') + '.xlsx')
filename_B = os.path.join('D:\\数据比对', today.strftime('%Y%m%d') + '.xlsx')

# 打印文件名，检查文件名是否正确
print(f"昨日文件名: {filename_A}")
print(f"今日文件名: {filename_B}")

# 检查文件是否存在
if not os.path.exists(filename_A):
    print(f"文件 {filename_A} 不存在。")
    raise FileNotFoundError(f"文件 {filename_A} 不存在。")

if not os.path.exists(filename_B):
    print(f"文件 {filename_B} 不存在。")
    raise FileNotFoundError(f"文件 {filename_B} 不存在。")

# 读取工作表A和B
df_A = pd.read_excel(filename_A, engine='openpyxl')
df_B = pd.read_excel(filename_B, engine='openpyxl')

# 获取身份证列名
id_column = '身份证'

# 1. 获取A表中存在，B表中不存在的数据条目
df_jianbao = df_A[~df_A[id_column].isin(df_B[id_column])]

# 2. 获取A表中不存在，B表中存在的数据条目
df_jiabao = df_B[~df_B[id_column].isin(df_A[id_column])]

# 3. 获取A表和B表中都存在，但其它列数据不同的数据条目（替换）
df_replace = pd.DataFrame(columns=df_B.columns)

# 指定要比对的列
compare_columns = ['雇主单位', '职业类别', '工种', '用工单位', '方案']

# 遍历B表中的每一行数据
for index, row_b in df_B.iterrows():
    # 在A表中查找相同身份证的数据行
    match_row_a = df_A[df_A[id_column] == row_b[id_column]]

    if not match_row_a.empty:
        # 检查指定的列是否相同
        diff_values = []
        for col in compare_columns:
            if match_row_a[col].iloc[0] != row_b[col]:
                diff_values.append((col, match_row_a[col].iloc[0], row_b[col]))

        # 如果有不同的列，则将该行数据保存到替换表中并打印差异
        if diff_values:
            df_replace = pd.concat([df_replace, row_b.to_frame().T], ignore_index=True)
            print(f"身份证号码为 {row_b[id_column]} 的数据存在差异:")
            for col, value_a, value_b in diff_values:
                print(f"  {col}: A表={value_a}, B表={value_b}")
            print()  # 添加空行作为分隔

# 构造保存结果的文件名
result_filename = os.path.join(os.path.dirname(filename_A), '比对结果.xlsx')

# 将结果保存到新的Excel文件中
with pd.ExcelWriter(result_filename) as writer:
    df_jianbao.to_excel(writer, sheet_name='减保', index=False)
    df_jiabao.to_excel(writer, sheet_name='加保', index=False)
    df_replace.to_excel(writer, sheet_name='替换', index=False)

print("比对结果已保存到比对结果.xlsx文件中。")


