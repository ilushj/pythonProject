import pandas as pd
import os

# 指定要处理的 Excel 文件路径
input_file_path = r"d:\2023custom\merged_output.xlsx"   # 替换为你的文件路径

try:
    # 读取 Excel 文件
    df = pd.read_excel(input_file_path)

    # 检查是否存在“客户名称”、“客户赔付率”、“在保月份”和“业务员”列
    required_columns = ['客户名称', '客户赔付率', '在保月份', '业务员']
    if all(col in df.columns for col in required_columns):
        # 按“客户名称”和“业务员”以及“在保月份”升序排序
        df.sort_values(by=['客户名称', '业务员', '在保月份'], ascending=True, inplace=True)

        # 使用 pivot_table 将“客户名称”和“业务员”作为索引，“客户赔付率”横置
        pivot_df = df.pivot_table(index=['客户名称', '业务员'],
                                  columns=df.groupby(['客户名称', '业务员']).cumcount() + 1, values='客户赔付率',
                                  aggfunc='first')

        # 重置索引，使“客户名称”和“业务员”成为列
        pivot_df.reset_index(inplace=True)

        # 生成新的 Excel 文件路径
        output_file_path = os.path.join(os.path.dirname(input_file_path), 'pivoted_output.xlsx')

        # 保存横置后的数据到新的 Excel 文件
        pivot_df.to_excel(output_file_path, index=False)
        print(f"处理完成，结果已保存至 '{output_file_path}'")
    else:
        print("输入文件缺少必要的列。")

except Exception as e:
    print(f"处理文件时出错: {e}")