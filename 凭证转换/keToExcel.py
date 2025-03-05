import pandas as pd  # 导入 pandas 库用于数据处理
from datetime import datetime, timedelta  # 导入 datetime 和 timedelta 用于处理日期
import sys  # 导入 sys 模块用于获取程序相关信息
import os  # 导入 os 模块用于处理文件路径
import re

# 获取当前日期和次日日期，并将次日日期格式化为 'YYYY/MM/DD' 的形式
today = datetime.now()
tomorrow = today + timedelta(days=1)
effective_date = tomorrow.strftime('%Y/%m/%d')

# 动态获取 group_config.xlsx 的路径
if getattr(sys, 'frozen', False):
    # 打包后的环境，使用临时解压目录
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    output_base_path = os.path.dirname(sys.executable)
else:
    # 脚本运行环境，使用脚本所在目录
    base_path = os.path.dirname(os.path.abspath(__file__))
    output_base_path = base_path



config_file = os.path.join(base_path, "group_config.xlsx")  # 构建配置文件的完整路径

# 1. 通过 input 函数输入源文件路径和文件名
source_file = input("请输入源Excel文件路径和文件名：")
# 将单反斜杠替换为双反斜杠，或者直接用 os.path.normpath 规范化路径
# 去除 Unicode 控制字符
source_file = re.sub(r'[\u200b-\u200f\u202a-\u202e]', '', source_file)
source_file = source_file.replace("\\", "\\\\")
# 2. 读取输入的源 Excel 文件和配置文件
# 读取 Excel 文件，确保 '封面抬头' 为字符串类型
df = pd.read_excel(source_file, header=1, dtype={'封面抬头': str})
df = df.dropna(how='all')  # 删除所有列都是 NaN 的行




config_df = pd.read_excel(config_file)  # 读取配置文件

# 3. 获取特定列的数据，并跳过姓名为空的行
columns_needed = ['姓名', '身份证号', '派单类型', '岗位', '项目名称', '购买标准（元）', '身故或残疾额度（万元）',
                  '封面抬头', '备注']
df_selected = df[columns_needed]  # 选取指定的列
df_selected = df_selected[df_selected['姓名'].notna()]  # 跳过姓名为空的行
# 清理数据：去除空格，转换为小写
df_selected['封面抬头'] = df_selected['封面抬头'].str.strip().str.lower()

# 处理可能的空值（合并单元格导致）
df_selected['封面抬头'] = df_selected['封面抬头'].fillna(method='ffill')

# 定义公司列表（小写）
companies = ['安徽柯恩服务外包有限公司'.lower(), '安徽拓西人力资源管理有限公司'.lower(), '云南润才企业管理有限公司'.lower()]

# 过滤数据
df_kn = df_selected[df_selected['封面抬头'].isin(companies)]
df_wx = df_selected[~df_selected['封面抬头'].isin(companies)]

# 检查结果
# 4. 处理“派单类型”列，根据不同的条件进行分类转换
def process_change_type(row):
    change_type = row['派单类型']
    remark = str(row['备注']) if pd.notna(row['备注']) else ''

    if '增' in str(change_type):
        return '批增'
    elif '减' in str(change_type):
        return '批减'
    elif change_type == '替换':
        if '离' in remark:
            return '批减'
        elif '新' in remark:
            return '批增'
        else:
            return '未知'
    return change_type


df_selected['派单类型'] = df_selected.apply(process_change_type, axis=1)  # 应用处理函数

# 5. 增加新列“生效日期”
df_selected['生效日期'] = effective_date


# 定义函数获取组别号
def get_group_number(row, config_df):
    # 根据封面抬头判断公司类型
    if row['封面抬头'] in ['安徽柯恩服务外包有限公司', '安徽拓西人力资源管理有限公司', '云南润才企业管理有限公司']:
        company_type = '柯恩'
    else:
        company_type = '非柯恩'

    # 在配置文件中匹配相应的组别号
    match = config_df[
        (config_df['公司类型'] == company_type) &
        (config_df['购买标准'] == row['购买标准（元）']) &
        (config_df['身故或残疾额度（万元）'] == row['身故或残疾额度（万元）'])
        ]

    return match['组别号'].iloc[0] if not match.empty else None


df_selected['组别号'] = df_selected.apply(lambda row: get_group_number(row, config_df), axis=1)  # 应用获取组别号的函数

# 6. 根据封面抬头分割数据集为柯恩相关和非柯恩相关
df_kn = df_selected[df_selected['封面抬头'].isin(
    ['安徽柯恩服务外包有限公司', '安徽拓西人力资源管理有限公司', '云南润才企业管理有限公司'])]
df_wx = df_selected[~df_selected['封面抬头'].isin(
    ['安徽柯恩服务外包有限公司', '安徽拓西人力资源管理有限公司', '云南润才企业管理有限公司'])]

# 定义新的列名映射
new_columns = {
    '姓名': '姓名',
    '身份证号': '身份证',
    '派单类型': '变更类型',
    '生效日期': '生效日期',
    '岗位': '工种',
    '组别号': '组别号',
    '项目名称': '用工单位',
    '购买标准（元）': '购买标准',
    '身故或残疾额度（万元）': '保额'
}

# 对分割后的数据集重命名列并选取新的列
df_kn_new = df_kn.rename(columns=new_columns)[list(new_columns.values())]
df_wx_new = df_wx.rename(columns=new_columns)[list(new_columns.values())]


def clean_duplicates(df, output_batch_add_file):
    # 1. 获取所有身份证号及其出现次数
    id_counts = df['身份证'].value_counts()

    # 2. 初始化保留行索引和新提取数据的列表
    rows_to_keep = []
    batch_add_rows = []

    # 3. 遍历所有唯一的身份证号
    for id_num in df['身份证'].unique():
        # 获取该身份证号对应的所有行
        id_rows = df[df['身份证'] == id_num]

        # 4. 如果身份证号只出现一次（不重复），保留该行
        if id_counts[id_num] == 1:
            rows_to_keep.extend(id_rows.index.tolist())
            continue

        # 5. 检查“组别号”的唯一值数量
        unique_groups = id_rows['组别号'].nunique()

        # 6. 如果组别号相同，保留所有记录
        if unique_groups == 1:
            rows_to_keep.extend(id_rows.index.tolist())
            continue

        # 7. 身份证相同且组别号不相同的情况
        if unique_groups > 1:
            # 保留“批减”记录
            keep_rows = id_rows[id_rows['变更类型'] == '批减']
            rows_to_keep.extend(keep_rows.index.tolist())

            # 提取“批增”记录
            batch_add = id_rows[id_rows['变更类型'] == '批增']
            batch_add_rows.extend(batch_add.index.tolist())

    # 8. 生成清理后的主 DataFrame（保留“批减”等记录）
    df_cleaned = df.loc[rows_to_keep].reset_index(drop=True)

    # 9. 生成“批增”记录的 DataFrame 并保存为新文件
    df_batch_add = df.loc[batch_add_rows].reset_index(drop=True)
    if not df_batch_add.empty:
        df_batch_add.to_excel(output_batch_add_file, index=False)
        print(f"已生成批增记录文件：{output_batch_add_file}")
    else:
        print("没有批增记录需要提取")

    # 10. 返回清理后的主 DataFrame
    return df_cleaned


# 示例使用
# 假设 df 是您的原始 DataFrame
# df_cleaned = clean_duplicates(df, 'batch_add_records.xlsx')


# 对处理后的数据集进行查重和删除重复行操作
kn_batch_add_file = os.path.join(output_base_path, "柯恩提出重复新增数据.xlsx")
wx_batch_add_file = os.path.join(output_base_path, "皖信提出重复新增数据.xlsx")
df_kn_new = clean_duplicates(df_kn_new, kn_batch_add_file)
df_wx_new = clean_duplicates(df_wx_new, wx_batch_add_file)

# 8. 生成文件名并设置输出路径为 exe 所在目录
kn_filename = os.path.join(output_base_path, f"柯恩批改导入模板{today.strftime('%Y%m%d')}.xlsx")
wx_filename = os.path.join(output_base_path, f"皖信批改导入模板{today.strftime('%Y%m%d')}.xlsx")

# 9. 将处理后的数据集保存到新的 Excel 文件
df_kn_new.to_excel(kn_filename, index=False)
df_wx_new.to_excel(wx_filename, index=False)

print(f"文件已生成：{kn_filename} 和 {wx_filename}")
