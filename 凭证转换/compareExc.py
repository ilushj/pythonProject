import pandas as pd
import os
from datetime import datetime


def compare_excel_files(source_path, compare_path, source_cols, compare_cols, output_path):
    # 读取 Excel 文件
    try:
        source_df = pd.read_excel(source_path)
        if os.path.abspath(source_path) == os.path.abspath(compare_path):
            compare_df = source_df.copy()
        else:
            compare_df = pd.read_excel(compare_path)
    except Exception as e:
        print(f"读取 Excel 文件时出错: {e}")
        return

    # 处理组合列
    source_combine = []
    compare_combine = []

    # 处理源表组合列
    for col in source_cols:
        if col.startswith('"') and col.endswith('"'):
            fixed_value = col[1:-1].strip()
            data = pd.Series([fixed_value] * len(source_df))
            source_combine.append(data)
        else:
            if col not in source_df.columns:
                print(f"源表中不存在列: {col}")
                return
            data = source_df[col].astype(str).str.strip()
            source_combine.append(data)

    # 处理比对表组合列
    for col in compare_cols:
        if col.startswith('"') and col.endswith('"'):
            fixed_value = col[1:-1].strip()
            data = pd.Series([fixed_value] * len(compare_df))
            compare_combine.append(data)
        else:
            if col not in compare_df.columns:
                print(f"比对表中不存在列: {col}")
                return
            data = compare_df[col].astype(str).str.strip()
            compare_combine.append(data)

    # 创建组合列
    source_df['combined'] = pd.concat(source_combine, axis=1).agg(''.join, axis=1)
    compare_df['combined'] = pd.concat(compare_combine, axis=1).agg(''.join, axis=1)

    # 查找重复数据（源表有，比对表有）
    duplicate_data = source_df[source_df['combined'].isin(compare_df['combined'])]

    # 查找源表独有数据（源表有，比对表没有）
    unique_source_data = source_df[~source_df['combined'].isin(compare_df['combined'])]

    # 查找比对表独有数据（源表没有，比对表有）
    unique_compare_data = compare_df[~compare_df['combined'].isin(source_df['combined'])]

    # 创建输出文件名
    compare_cols_str = '_'.join([str(col) for col in compare_cols])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"{compare_cols_str}_比较结果_{timestamp}.xlsx"
    output_full_path = os.path.join(output_path, output_file)

    # 保存结果到一个 Excel 文件，包含三个 sheet
    with pd.ExcelWriter(output_full_path) as writer:
        duplicate_data.to_excel(writer, sheet_name='源表有_比对表有', index=False)
        unique_source_data.to_excel(writer, sheet_name='源表有_比对表没有', index=False)
        unique_compare_data.to_excel(writer, sheet_name='源表没有_比对表有', index=False)

    print(f"比较完成，结果已保存到: {output_full_path}")


def main():
    # 获取用户输入
    source_path = input("请输入源表文件路径: ")
    compare_path = input("请输入比对表文件路径: ")

    # 获取比较列
    source_cols_input = input("请输入源表比较列（用+分隔，固定值用双引号，例如: 身份证号+组别+\"批增\"): ")
    compare_cols_input = input("请输入比对表比较列（用+分隔，固定值用双引号，例如: 身份证+组别+变更类型）: ")

    # 分割输入的列名或字符串
    source_cols = [col.strip() for col in source_cols_input.split('+')]
    compare_cols = [col.strip() for col in compare_cols_input.split('+')]

    # 获取脚本所在目录作为输出路径
    output_path = os.path.dirname(os.path.abspath(__file__))

    # 执行比较
    compare_excel_files(source_path, compare_path, source_cols, compare_cols, output_path)


if __name__ == "__main__":
    main()