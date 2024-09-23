import pandas as pd

# 指定要读取的Excel文件路径
data_file = r'D:\群发email\用户信息含PIN码1.xlsx'

# 读取数据
df = pd.read_excel(data_file)

# 将数据按姓名分组
grouped = df.groupby('姓名')  # 请确保列名“姓名”与文件中的列名一致

# 导出每个姓名的记录到不同的Excel文件
for name, group in grouped:
    # 为每个姓名创建Excel文件
    output_file = f"D:\\群发email\\{name}.xlsx"
    group.to_excel(output_file, index=False)  # 不存储索引

print("文件已成功导出！")