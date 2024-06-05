import os
import pandas as pd
from datetime import datetime
import re


def process_sheet(df, columns_needed, new_column_names, change_type, new_filename, sheet_name):
    # 检查是否存在所有指定的列
    if all(col in df.columns for col in columns_needed):
        # 选择并重命名列
        selected_data = df[columns_needed].copy()
        selected_data.columns = new_column_names

        # 处理生效日期数据
        selected_data['生效日期'] = selected_data['生效日期'].apply(lambda x: re.sub(r'零点|零时', '', str(x)))
        try:
            selected_data['生效日期'] = pd.to_datetime(selected_data['生效日期'], format='%Y%m%d')
        except Exception as e:
            print(f"生效日期转换错误: {e}")
            return

        # 添加新列及默认值
        if sheet_name == "减保":
            selected_data['工种'] = ""
        selected_data['护照'] = ""
        selected_data['变更类型'] = change_type
        selected_data['备注'] = ""
        # 重新排列列的顺序
        new_column_order = ['姓名', '身份证', '护照', '变更类型', '生效日期', '工种', '备注']
        selected_data = selected_data[new_column_order]

        # 如果文件已经存在，追加sheet，否则创建新文件
        if os.path.exists(new_file_path):
            with pd.ExcelWriter(new_file_path, mode='a', if_sheet_exists='replace') as writer:
                selected_data['身份证'] = selected_data['身份证'].astype(str)
                selected_data.columns = ['姓名', '身份证', '护照', '变更类型', '生效日期', '工种', '备注']
                selected_data.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(new_file_path) as writer:
                selected_data['身份证'] = selected_data['身份证'].astype(str)
                selected_data.columns = ['姓名', '身份证', '护照', '变更类型', '生效日期', '工种', '备注']
                selected_data.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        missing_cols = [col for col in columns_needed if col not in df.columns]
        print(f"{sheet_name} 中的 {new_filename} 数据框缺少指定的列名: {missing_cols}")


# 定义目录路径
directory = r"d:\保全模板test"

# 遍历目录中的文件
for filename in os.listdir(directory):
    if filename.endswith(".xlsx") and "众安" in filename:
        file_path = os.path.join(directory, filename)
        # 读取xlsx文件
        xls = pd.ExcelFile(file_path)

        # 生成新的文件名和路径
        new_filename = filename.replace('.xlsx', '_奇点导入.xlsx')
        new_file_path = os.path.join(directory, new_filename)

        # 处理每个sheet
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if "增加" in sheet_name:
                process_sheet(
                    df,
                    columns_needed=['姓名*', '证件号码*', '生效日期*', '职业(职业代码)*'],
                    new_column_names=['姓名', '身份证', '生效日期', '工种'],
                    change_type="加保",
                    new_filename=new_file_path,
                    sheet_name='加保'
                )
            elif "减少" in sheet_name:
                process_sheet(
                    df,
                    columns_needed=['姓名*', '证件号码*', '生效日期*'],
                    new_column_names=['姓名', '身份证', '生效日期'],
                    change_type="减保",
                    new_filename=new_file_path,
                    sheet_name='减保'
                )
