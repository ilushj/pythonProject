import pandas as pd
import os
from datetime import datetime
import sys


def compare_excel_files(source_path, compare_path, source_cols, compare_cols):
    """Excel文件比对核心函数"""
    # 初始化 application_path 用于后续结果文件存储路径
    application_path = ""

    # 读取 Excel 文件（支持源表和比对表为同一文件的情况）
    try:
        # 读取源表数据
        source_df = pd.read_excel(source_path)

        # 判断是否为同一文件路径（考虑不同格式的路径可能指向同一文件）
        if os.path.abspath(source_path) == os.path.abspath(compare_path):
            compare_df = source_df.copy()  # 避免重复读取相同文件
        else:
            compare_df = pd.read_excel(compare_path)
    except Exception as e:
        print(f"读取 Excel 文件时出错: {e}")
        return

    # 处理组合列（支持固定值和动态列的组合）
    source_combine = []  # 存储源表组合列数据
    compare_combine = []  # 存储比对表组合列数据

    # 处理源表组合列逻辑
    for col in source_cols:
        # 处理固定值情况（用双引号包裹的字符串）
        if col.startswith('"') and col.endswith('"'):
            fixed_value = col[1:-1].strip()  # 去除双引号和首尾空格
            data = pd.Series([fixed_value] * len(source_df))  # 创建等长固定值序列
            source_combine.append(data)
        # 处理动态列情况
        else:
            if col not in source_df.columns:
                print(f"源表中不存在列: {col}")
                return
            # 转换为字符串并去除首尾空格
            data = source_df[col].astype(str).str.strip()
            source_combine.append(data)

    # 处理比对表组合列逻辑（与源表处理逻辑相同）
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

    # 创建组合列（将多个列/固定值拼接为单一字符串）
    source_df['combined'] = pd.concat(source_combine, axis=1).agg(''.join, axis=1)
    compare_df['combined'] = pd.concat(compare_combine, axis=1).agg(''.join, axis=1)

    # 数据比对逻辑
    # 查找两表交集数据（源表有且比对表有）
    duplicate_data = source_df[source_df['combined'].isin(compare_df['combined'])]

    # 查找源表独有数据（源表有但比对表没有）
    unique_source_data = source_df[~source_df['combined'].isin(compare_df['combined'])]

    # 查找比对表独有数据（源表没有但比对表有）
    unique_compare_data = compare_df[~compare_df['combined'].isin(source_df['combined'])]

    # 生成结果文件名（包含时间戳防止重复）
    compare_cols_str = '_'.join([str(col) for col in compare_cols])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"{compare_cols_str}_比较结果_{timestamp}.xlsx"

    # 确定输出路径（优先使用exe所在目录，开发时使用脚本所在目录）
    if getattr(sys, 'frozen', False):  # 打包成exe时的情况
        application_path = os.path.dirname(sys.executable)
    elif __file__:  # 直接运行脚本时的情况
        application_path = os.path.dirname(__file__)

    output_full_path = os.path.join(application_path, output_file)

    # 将结果写入Excel的不同sheet页
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

    # 执行比较
    compare_excel_files(source_path, compare_path, source_cols, compare_cols)


if __name__ == "__main__":
    main()
