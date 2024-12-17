import pandas as pd

# 获取用户输入的月份
month = input("请输入月份（如：1、2、3...）：")

# 1. 文件路径
rule_path = '易久保规则.xlsx'
year_path = f"{month}全年.xlsx"
current_month_path = f"{month}当月.xlsx"
output_path = f"{month}月佣金数据.xlsx"  # 结果文件路径
hulin_file = f"{month}胡林特殊.xlsx"  # 结果文件路径


# 2. 读取规则和全年表
rules = pd.read_excel(rule_path)
year_data = pd.read_excel(year_path)
current_month_data = pd.read_excel(current_month_path)
hulin_data = pd.read_excel(hulin_file)

# 提取胡林特殊.xlsx中的“业务员”和“客户名称”列
hulin_data = hulin_data[['业务员', '客户名称']].drop_duplicates()

# 3. 百分比转换函数
def percentage_to_float(series):
    """
    将百分比字符串列转换为浮点数（0~1）。
    若已是小数形式（如 0.2），直接返回。
    """
    series = series.astype(str)
    if series.str.contains('%').any():  # 检查是否包含 "%"
        return series.str.rstrip('%').astype(float) / 100
    return series.astype(float)


# 转换规则表中相关列（确保只处理百分比字段）
rules.iloc[:, :6] = rules.iloc[:, :6].apply(percentage_to_float)

# 转换全年表中的赔付率列
year_data['客户赔付率'] = percentage_to_float(year_data['客户赔付率'])
year_data['归属赔付率'] = percentage_to_float(year_data['归属赔付率'])
year_data['个人赔付率'] = percentage_to_float(year_data['个人赔付率'])

# 4. 客户赔付率计算
year_data['客户赔付率'] = year_data.apply(
    lambda row: row['归属赔付率'] if pd.notna(row['归属赔付率']) else row['客户赔付率'], axis=1
)


# 5. 比对函数
def match_rule(person_rate, client_rate):
    """
    根据个人赔付率和客户赔付率匹配规则，返回业绩比例和提奖比例。
    """
    for i in range(len(rules)):
        if (rules.iloc[i, 0] <= person_rate < rules.iloc[i, 1] and
                rules.iloc[i, 2] <= client_rate < rules.iloc[i, 3]):
            return rules.iloc[i, 4], rules.iloc[i, 5]
    return None, None


# 6. 应用规则
result = year_data.apply(
    lambda row: match_rule(row['个人赔付率'], row['客户赔付率']),
    axis=1
)

# 7. 分配结果
year_data['业绩比例'], year_data['提奖比例'] = zip(*result)

# 删除可能重复的列
current_month_data = current_month_data.drop(columns=['客户赔付率', '个人赔付率'], errors='ignore')


# 8. 数据匹配
# 当前月数据（业务员 + 客户名称）与全年数据匹配
merged_data = pd.merge(
    current_month_data,  # 当月.xlsx 的数据
    year_data[['业务员', '客户名称', '客户赔付率', '个人赔付率', '业绩比例', '提奖比例']],  # 匹配用字段
    on=['业务员', '客户名称'],  # 匹配条件
    how='left'  # 左连接，保留当月.xlsx 的所有数据
)

# 计算业绩和提奖
merged_data['业绩'] = (
    merged_data['总保费'] *
    merged_data['佣金折扣'] *
    merged_data['业绩比例']
).round(2)  # 保留两位小数

merged_data['提奖'] = (
    merged_data['总保费'] *
    merged_data['佣金折扣'] *
    merged_data['提奖比例']
).round(2)  # 保留两位小数

# 9. 保留字段
result_data = merged_data[
    ['业务员', '客户名称', '在保月份', '投保方案', '总保费', '佣金折扣', '项目类型',
     '客户赔付率', '个人赔付率', '业绩比例', '提奖比例', '业绩', '提奖']
]


# 匹配出与胡林特殊.xlsx相同的行
matched_data = result_data.merge(hulin_data, on=['业务员', '客户名称'], how='inner')

# 从result_data中移除这些匹配的行
filtered_data = result_data.merge(hulin_data, on=['业务员', '客户名称'], how='left', indicator=True)
filtered_data = filtered_data[filtered_data['_merge'] == 'left_only'].drop(columns=['_merge'])

# 保存匹配的数据到胡林当月.xlsx
matched_data.to_excel(f"{month}月胡林佣金.xlsx", index=False)

# 按照业务员名称降序排序
filtered_data = filtered_data.sort_values(by='业务员', ascending=False)

# 10. 保存结果
filtered_data.to_excel(output_path, index=False)  # 自动覆盖原文件

print(f"文件已保存并覆盖至 {output_path}")
