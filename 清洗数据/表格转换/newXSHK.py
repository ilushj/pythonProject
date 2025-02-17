import pandas as pd

# 读取原 Excel 文件
file_path = "业绩转换.xlsx"
with pd.ExcelFile(file_path) as excel:
    df = pd.read_excel(excel, sheet_name="CWYJ")

# 条件筛选：新老业绩列中只选择值为 "新业绩"、"新2"、"新结" 的数据行
filtered_df = df[df["新老业绩"].isin(["新业绩", "新2", "新结"])]

# 创建新的数据表 xshk
xshk = pd.DataFrame()

# 列映射
xshk["申请人"] = filtered_df["姓名"]
xshk["上报日期"] = filtered_df["合作日期"]
xshk["新总回款金额"] = filtered_df["回款金额"]
xshk["申请人区域"] = filtered_df["销售地区"]
xshk["销售经理"] = filtered_df["销售经理"]
xshk["销售总监"] = filtered_df["销售总监"]
xshk["销售副总"] = filtered_df["销售副总"]
xshk["总经理"] = filtered_df["总经理"]
xshk["总裁"] = filtered_df["总裁"]


# 添加空列
empty_columns = [
    "下店合作新业绩",
    "市场引流医院新业绩",
    "一个月内二次医院成交的新业绩",
    "回收欠新业绩",
    "职位",
    "在职情况",
    "上海经理排名周期",
    "销售9月份抽调机制",
    "培训店跟踪周期",
    "结算周期",
    "销售5月份抽调机制（总）",
    "新业绩回款（下店+回收欠款）",
    "市场引流回款（引流+2开）",
    "是否归属销售部"
]
for col in empty_columns:
    xshk[col] = None

# 调整列顺序
column_order = [
    "申请人",
    "上报日期",
    "下店合作新业绩",
    "市场引流医院新业绩",
    "一个月内二次医院成交的新业绩",
    "回收欠新业绩",
    "新总回款金额",
    "申请人区域",
    "销售经理",
    "销售总监",
    "销售副总",
    "总经理",
    "总裁",
    "职位",
    "在职情况",
    "上海经理排名周期",
    "新业绩回款（下店+回收欠款）",
    "市场引流回款（引流+2开）",
    "销售9月份抽调机制",
    "培训店跟踪周期",
    "结算周期",
    "销售5月份抽调机制（总）",
    "是否归属销售部",
]
xshk = xshk[column_order]

# 保存到新的 Excel 文件
# xshk.to_excel("xshk.xlsx", index=False)

# print("数据表 xshk 已生成并保存为 xshk.xlsx")

# 将新表 xshk 写入原 Excel 文件的新 Sheet 中
with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    xshk.to_excel(writer, sheet_name="XSHK", index=False)

print(f"新表 xshk 已成功写入原文件 {file_path} 中！")
input("按任意键退出程序...")
