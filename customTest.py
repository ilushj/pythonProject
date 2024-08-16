import pandas as pd
import os

# 指定要扫描的目录
directory = r"d:\2023custom"  # 请将此处替换为你的目录

merged_data = []

# 遍历目录中的所有文件
for filename in os.listdir(directory):
    if filename.endswith('.xlsx') and filename != 'merged_output.xlsx':  # 排除已有的合并文件
        file_path = os.path.join(directory, filename)

        # 读取 Excel 文件的 Sheet1
        try:
            df = pd.read_excel(file_path, sheet_name='sheet1')

            # 选取需要的列
            if all(col in df.columns for col in ['业务员', '客户名称', '归属单位', '在保月份', '总保费', '客户赔付率', '归属赔付率']):
                selected_data = df[['业务员', '客户名称', '归属单位', '在保月份', '总保费', '客户赔付率', '归属赔付率']]
                merged_data.append(selected_data)
            else:
                print(f"文件 {filename} 缺少必要的列")

        except Exception as e:
            print(f"处理文件 {filename} 时出错: {e}")

        # 合并所有提取的数据
if merged_data:
    result_df = pd.concat(merged_data, ignore_index=True)

    # 保存合并后的数据到新的 Excel 文件（保存到原目录）
    output_file_path = os.path.join(directory, 'merged_output.xlsx')
    result_df.to_excel(output_file_path, index=False)
    print(f"合并完成，结果已保存至 '{output_file_path}'")
else:
    print("没有找到有效的数据。")