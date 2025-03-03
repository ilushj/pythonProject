import pandas as pd  # 导入 pandas 库用于数据处理
from datetime import datetime, timedelta  # 导入 datetime 和 timedelta 用于处理日期
import sys  # 导入 sys 模块用于获取程序相关信息
import os  # 导入 os 模块用于处理文件路径

# 获取当前日期和次日日期，并将次日日期格式化为 'YYYY/MM/DD' 的形式
today = datetime.now()
tomorrow = today + timedelta(days=1)
effective_date = tomorrow.strftime('%Y/%m/%d')

# 动态获取程序所在目录和 group_config.xlsx 的路径
if getattr(sys, 'frozen', False):
    # 如果是打包后的 exe 文件，获取 exe 所在目录
    base_path = os.path.dirname(sys.executable)
else:
    # 如果是 Python 脚本运行，获取脚本所在目录
    base_path = os.path.dirname(os.path.abspath(__file__))

config_file = os.path.join(base_path, "group_config.xlsx")  # 构建配置文件的完整路径

# 1. 通过 input 函数输入源文件路径和文件名
source_file = input("请输入源Excel文件路径和文件名：")

# 2. 读取输入的源 Excel 文件和配置文件
df = pd.read_excel(source_file, header=1)  # 读取源文件，跳过第一行作为表头
config_df = pd.read_excel(config_file)  # 读取配置文件

# 3. 获取特定列的数据，并跳过姓名为空的行
columns_needed = ['姓名', '身份证号', '派单类型', '岗位', '项目名称', '购买标准（元）', '身故或残疾额度（万元）',
                  '封面抬头', '备注']
df_selected = df[columns_needed]  # 选取指定的列
df_selected = df_selected[df_selected['姓名'].notna()]  # 跳过姓名为空的行


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


# 7. 定义查重和删除重复行的函数
def clean_duplicates(df):
    duplicated_ids = df[df['身份证'].duplicated(keep=False)]['身份证'].unique()

    if not duplicated_ids.any():
        return df

    rows_to_keep = []
    for id_num in duplicated_ids:
        id_rows = df[df['身份证'] == id_num]
        if id_rows['组别号'].nunique() > 1:
            if '批减' in id_rows['变更类型'].values:
                keep_rows = id_rows[id_rows['变更类型'] != '批减']
                rows_to_keep.extend(keep_rows.index.tolist())
            else:
                rows_to_keep.extend(id_rows.index.tolist())
        else:
            rows_to_keep.extend(id_rows.index.tolist())

    return df.loc[rows_to_keep].reset_index(drop=True)


# 对处理后的数据集进行查重和删除重复行操作
df_kn_new = clean_duplicates(df_kn_new)
df_wx_new = clean_duplicates(df_wx_new)

# 8. 生成文件名并设置输出路径为 exe 所在目录
kn_filename = os.path.join(base_path, f"柯恩批改导入模板{today.strftime('%Y%m%d')}.xlsx")
wx_filename = os.path.join(base_path, f"皖信批改导入模板{today.strftime('%Y%m%d')}.xlsx")

# 9. 将处理后的数据集保存到新的 Excel 文件
df_kn_new.to_excel(kn_filename, index=False)
df_wx_new.to_excel(wx_filename, index=False)

print(f"文件已生成：{kn_filename} 和 {wx_filename}")
