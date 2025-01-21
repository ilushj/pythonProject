import pandas as pd

# 读取 Excel 文件
file_path = '11.xlsx'  # 替换为实际的文件路径
sheet_data = pd.read_excel(file_path)

# 定义行业分类函数
def classify_industry(name):
    if "玻璃" in name:
        return "玻璃制造业"
    elif "汽车" in name or "车" in name:
        return "汽车制造业"
    elif "电气" in name or "电" in name:
        return "电气设备制造"
    elif "生物" in name:
        return "生物技术"
    elif "科技" in name or "信息" in name or "智慧" in name or "新能源" in name or "新材料" in name or "智能" in name or "机器人" in name:
        return "科技研发"
    elif "股份" in name:
        return "综合企业"
    elif "办公" in name or "文化" in name:
        return "办公"
    elif "物业" in name:
        return "物业"
    elif "化妆" in name:
        return "化妆"
    elif "餐饮" in name or "酒" in name or "食" in name or "便利" in name or "餐" in name or "厨" in name:
        return "餐饮食品业"
    elif "供应链" in name or "物流" in name or "冷藏" in name or "储" in name or "运" in name or "仓" in name:
        return "物流业"
    elif "移动" in name or "通信" in name :
        return "电子通讯"
    elif "人力" in name or "劳务" in name or "服务" in name:
        return "人力资源"
    elif "机械" in name or "精工" in name or "设备" in name or "制造" in name or "制品" in name or "模具" in name or "元件" in name or "材料" in name or "加工" in name or "精密" in name:
        return "机械设备制造"

    else:
        return "未分类"

# 应用分类函数
sheet_data['所属行业'] = sheet_data['用工单位'].apply(classify_industry)

# 保存结果到新的 Excel 文件
output_path = '11_with_industry.xlsx'  # 替换为想要保存的路径
sheet_data.to_excel(output_path, index=False)
print(f"结果已保存到 {output_path}")
