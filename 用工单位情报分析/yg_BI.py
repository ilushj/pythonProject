import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # 设置无GUI后端
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from sklearn.preprocessing import MinMaxScaler

# 设置中文支持
plt.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
plt.rcParams['axes.unicode_minus'] = False    # 修复负号显示

# 1. 数据加载与清洗
df = pd.read_excel('test1.xlsx')
df['最早投保'] = pd.to_datetime(df['最早投保'])
df['最晚投保'] = pd.to_datetime(df['最晚投保'])
date_error = df[df['最早投保'] > df['最晚投保']]
if not date_error.empty:
    print("发现投保时间逻辑错误：")
    print(date_error[['用工单位名称', '最早投保', '最晚投保']])

# 2. 特征工程
# 计算百人出险率
df['百人出险率'] = (df['出险次数'] / df['当前在保人数'].replace(0, 1)) * 100
# 新增骨折出险率
df['骨折出险率'] = (df['骨折次数'] / df['当前在保人数'].replace(0, 1)) * 100
df['人均月风险金额'] = df['出险金额'] / df['折算月人数'].replace(0, 1)

# 3. 风险评分模型
scaler = MinMaxScaler()
# 将骨折出险率纳入风险因素
risk_factors = df[['出险次数', '骨折次数', '出险金额', '骨折出险率']]
scaled_factors = scaler.fit_transform(risk_factors)

# 调整加权评分模型，新增骨折出险率权重
weights = np.array([0.35, 0.25, 0.25, 0.15])  # 出险次数，骨折次数，出险金额，骨折出险率
df['风险评分'] = np.dot(scaled_factors, weights)
df['风险等级'] = pd.qcut(df['风险评分'], q=3, labels=['低风险', '中风险', '高风险'])

# 4. 可视化分析
plt.figure(figsize=(15, 18))

# 高风险单位TOP10
plt.subplot(3, 2, 1)
top_high_risk = df.nlargest(10, '风险评分')
sns.barplot(x='风险评分', y='用工单位名称', data=top_high_risk, palette='Reds_r')
plt.title('高风险单位TOP10')

# 在保天数与出险次数的关系
plt.subplot(3, 2, 2)
sns.scatterplot(x='在保天数', y='出险次数', hue='风险等级', data=df, size='骨折出险率', sizes=(20, 200))
plt.title('在保天数与出险次数的关系（点大小：骨折出险率）')

# 行业分类
industries = {
    '制造业': ['制造', '机电', '机械', '电子', '科技'],
    '物流': ['物流', '供应链'],
    '食品': ['食品', '餐饮'],
    '物业': ['物业']
}

def classify_industry(name):
    for industry, keywords in industries.items():
        if any(keyword in name for keyword in keywords):
            return industry
    return '其他'

df['行业分类'] = df['用工单位名称'].apply(classify_industry)
plt.subplot(3, 2, 3)
sns.boxplot(x='行业分类', y='骨折出险率', data=df)
plt.xticks(rotation=45)
plt.title('分行业骨折出险率分布')

# 时间趋势分析（骨折次数）
df['投保年份'] = df['最早投保'].dt.year
plt.subplot(3, 2, 4)
sns.lineplot(x='投保年份', y='骨折次数', data=df, estimator='sum', ci=None)
plt.title('年度投保单位骨折次数趋势')

# 风险因素相关性热力图
plt.subplot(3, 2, 5)
corr_matrix = df[['风险评分', '出险次数', '骨折次数', '出险金额', '骨折出险率', '在保天数', '当前在保人数']].corr()
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm')
plt.title('风险因素相关性分析')

plt.tight_layout()
plt.savefig('风险分析图表.png', dpi=300, bbox_inches='tight')
plt.close()

# 5. 结果输出
high_risk_report = df[df['风险等级'] == '高风险'].sort_values('风险评分', ascending=False)
report_columns = [
    '用工单位名称', '行业分类', '风险评分', '风险等级',
    '出险次数', '骨折次数', '出险金额', '当前在保人数',
    '百人出险率', '骨折出险率'
]

print("\n高风险单位分析报告：")
print(high_risk_report[report_columns].to_string(index=False))

# 修改改进建议，考虑骨折次数标准
def generate_recommendation(row):
    recommendations = []
    if row['百人出险率'] > 10:
        recommendations.append("建议加强安全培训")
    # 新增骨折相关建议：投保1年内骨折>=4起
    if row['在保天数'] <= 365 and row['骨折次数'] >= 4:
        recommendations.append("建议加强骨折防护措施")
    elif row['骨折出险率'] > 5:  # 骨折出险率高于5%也需关注
        recommendations.append("建议关注骨折风险")
    if row['当前在保人数'] < 100 and row['出险次数'] > 5:
        recommendations.append("建议优化用工结构")
    if row['在保天数'] < 365:
        recommendations.append("建议延长保险周期")
    return "; ".join(recommendations) if recommendations else "暂无建议"

high_risk_report['改进建议'] = high_risk_report.apply(generate_recommendation, axis=1)

print("\n改进建议：")
print(high_risk_report[['用工单位名称', '改进建议']].to_string(index=False))

# 6. 数据保存
high_risk_report.to_excel('高风险单位分析报告.xlsx', index=False)