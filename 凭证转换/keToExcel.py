import pandas as pd
from datetime import datetime, timedelta
import sys
import os

# 获取当前日期和次日日期
today = datetime.now()
tomorrow = today + timedelta(days=1)
effective_date = tomorrow.strftime('%Y/%m/%d')

# 动态获取 group_config.xlsx 的路径
if getattr(sys, 'frozen', False):
    # 如果是打包后的 exe，文件在临时解压目录（sys._MEIPASS）
    base_path = sys._MEIPASS
else:
    # 如果是 Python 脚本运行，文件在脚本所在目录
    base_path = os.path.dirname(os.path.abspath(__file__))

config_file = os.path.join(base_path, "group_config.xlsx")

# 1. 通过 input 输入源文件路径和文件名
source_file = input("请输入源Excel文件路径和文件名：")

# 2. 读取 Excel 文件和配置文件
df = pd.read_excel(source_file, header=1)
config_df = pd.read_excel(config_file)

# 3. 获取特定列数据并跳过姓名为空的行
columns_needed = ['姓名', '身份证号', '派单类型', '岗位', '项目名称', '购买标准（元）', '身故或残疾额度（万元）',
                  '封面抬头', '备注']
df_selected = df[columns_needed]
df_selected = df_selected[df_selected['姓名'].notna()]  # 跳过姓名为空的行


# 4. 处理“派单类型”列
def process_change_type(row):
    change_type = row['派单类型']
    remark = str(row['备注']) if pd.notna(row['备注']) else ''  # 将备注转换为字符串并处理空值

    # 规则1：包含“增”或“减”
    if '增' in str(change_type):
        return '批增'
    elif '减' in str(change_type):
        return '批减'
    # 规则2：值为“替换”，根据“备注”判断
    elif change_type == '替换':
        if '离' in remark:
            return '批减'
        elif '新' in remark:
            return '批增'
    # 如果不符合上述条件，返回原始值
    return change_type


# 应用处理函数到“派单类型”列
df_selected['派单类型'] = df_selected.apply(process_change_type, axis=1)

# 5. 增加新列
# 添加生效日期列
df_selected['生效日期'] = effective_date


# 定义组别号生成函数（基于外部配置文件）
def get_group_number(row, config_df):
    # 判断公司类型
    if row['封面抬头'] in ['安徽柯恩服务外包有限公司', '安徽拓西人力资源管理有限公司', '云南润才企业管理有限公司']:
        company_type = '柯恩'
    else:
        company_type = '非柯恩'

    # 从配置文件中查找匹配的组别号
    match = config_df[
        (config_df['公司类型'] == company_type) &
        (config_df['购买标准'] == row['购买标准（元）']) &
        (config_df['身故或残疾额度（万元）'] == row['身故或残疾额度（万元）'])
        ]

    # 如果找到匹配项，返回对应的组别号，否则返回 None
    if not match.empty:
        return match['组别号'].iloc[0]
    else:
        return None


# 添加组别号列
df_selected['组别号'] = df_selected.apply(lambda row: get_group_number(row, config_df), axis=1)

# 6. 根据封面抬头分割数据集
df_kn = df_selected[df_selected['封面抬头'].isin(
    ['安徽柯恩服务外包有限公司', '安徽拓西人力资源管理有限公司', '云南润才企业管理有限公司'])]
df_wx = df_selected[~df_selected['封面抬头'].isin(
    ['安徽柯恩服务外包有限公司', '安徽拓西人力资源管理有限公司', '云南润才企业管理有限公司'])]

# 定义新列名对应关系和顺序
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

# 重命名并按指定顺序排列列
df_kn_new = df_kn.rename(columns=new_columns)[list(new_columns.values())]
df_wx_new = df_wx.rename(columns=new_columns)[list(new_columns.values())]


# 7. 定义查重和删除函数
def clean_duplicates(df):
    # 查找身份证重复的行
    duplicated_ids = df[df['身份证'].duplicated(keep=False)]['身份证'].unique()

    # 如果没有重复，直接返回原数据
    if not duplicated_ids.any():
        return df

    # 存储需要保留的行索引
    rows_to_keep = []

    # 处理每一组重复的身份证
    for id_num in duplicated_ids:
        id_rows = df[df['身份证'] == id_num]
        # 如果组别号不同
        if id_rows['组别号'].nunique() > 1:
            # 检查是否有“批减”行
            if '批减' in id_rows['变更类型'].values:
                # 保留非“批减”的行
                keep_rows = id_rows[id_rows['变更类型'] != '批减']
                rows_to_keep.extend(keep_rows.index.tolist())
            else:
                # 如果没有“批减”，保留所有行
                rows_to_keep.extend(id_rows.index.tolist())
        else:
            # 如果组别号相同，保留所有行
            rows_to_keep.extend(id_rows.index.tolist())

    # 返回清洗后的数据
    return df.loc[rows_to_keep].reset_index(drop=True)


# 分别对 df_kn_new 和 df_wx_new 进行查重和清洗
df_kn_new = clean_duplicates(df_kn_new)
df_wx_new = clean_duplicates(df_wx_new)

# 生成文件名（包含当天日期）
kn_filename = f"柯恩批改导入模板{today.strftime('%Y%m%d')}.xlsx"
wx_filename = f"皖信批改导入模板{today.strftime('%Y%m%d')}.xlsx"

# 8. 保存到新的 Excel 文件
df_kn_new.to_excel(kn_filename, index=False)
df_wx_new.to_excel(wx_filename, index=False)

print(f"文件已生成：{kn_filename} 和 {wx_filename}")