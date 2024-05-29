import pandas as pd
from datetime import datetime, timedelta
import os


def process_excel_files(directory):
    filenames = os.listdir(directory)

    for filename in filenames:
        if filename.endswith(".xlsx"):
            file_path = os.path.join(directory, filename)
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                if "增员" in sheet_names:
                    df_add = pd.read_excel(file_path, sheet_name="增员")

                    # 只取姓名列不为空的数据
                    df_add = df_add[df_add.filter(like='姓名').notnull().any(axis=1)]

                    new_df_add = pd.DataFrame(
                        columns=["姓名", "身份证", "护照", "变更类型", "生效日期", "工种", "有无社保", "备注"])

                    for col in df_add.columns:
                        if "姓名" in col:
                            new_df_add["姓名"] = df_add[col]
                        elif "证件号码" in col:
                            new_df_add["身份证"] = df_add[col]
                        elif "工种" in col:
                            new_df_add["工种"] = df_add[col]

                    new_df_add["生效日期"] = datetime.now() + timedelta(days=1)

                    # 删除姓名列为空的行
                    new_df_add = new_df_add.dropna(subset=["姓名"])

                    # 生成新的excel文件
                    new_filename_add = filename.split(".")[0] + "-奇点批增.xlsx"
                    new_file_path_add = os.path.join(directory, new_filename_add)
                    new_df_add.to_excel(new_file_path_add, index=False)

                if "减员" in sheet_names:
                    df_del = pd.read_excel(file_path, sheet_name="减员")

                    # 只取姓名列不为空的数据
                    df_del = df_del[df_del.filter(like='姓名').notnull().any(axis=1)]

                    new_df_del = pd.DataFrame(
                        columns=["姓名", "身份证", "护照", "变更类型", "生效日期", "工种", "有无社保", "备注"])

                    for col in df_del.columns:
                        if "姓名" in col:
                            new_df_del["姓名"] = df_del[col]
                        elif "证件号码" in col:
                            new_df_del["身份证"] = df_del[col]
                        elif "工种" in col:
                            new_df_del["工种"] = df_del[col]

                    new_df_del["生效日期"] = datetime.now() + timedelta(days=1)

                    # 删除姓名列为空的行
                    new_df_del = new_df_del.dropna(subset=["姓名"])

                    # 生成新的excel文件
                    new_filename_del = filename.split(".")[0] + "-奇点批减.xlsx"
                    new_file_path_del = os.path.join(directory, new_filename_del)
                    new_df_del.to_excel(new_file_path_del, index=False)

            except Exception as e:
                print(f"文件 {filename} 不符合条件，跳过处理。错误信息：{e}")


# 调用函数，传入目录路径
process_excel_files("d:/保全模板test")