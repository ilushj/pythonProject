import pandas as pd

# 1. 文件路径
rule_path = '易久保规则.xlsx'
year_path = '全年.xlsx'
output_path = '当月赔付数据.xlsx'  # 结果文件路径

# 2. 读取规则和全年表
rules = pd.read_excel(rule_path)
year_data = pd.read_excel(year_path)

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

# 8. 保存结果
output_data = year_data[[
    '业务员', '客户名称', '客户赔付率', '归属赔付率', '个人赔付率', '业绩比例', '提奖比例'
]]
output_data.to_excel(output_path, index=False)  # 自动覆盖原文件

print(f"文件已保存并覆盖至 {output_path}")
