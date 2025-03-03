import os
import re
import PyPDF2
from openpyxl import Workbook


def extract_refund_total(text):
    # 首先尝试提取“退费合计：”
    refund_total = re.search(r'退费合计：\s*([^，]+)', text)

    if refund_total:
        return refund_total.group(1)  # 提取第一个捕获组

    # 如果没有找到“退费合计：”，则检查其他两种情况
    refund_total = re.search(r'(加收价税合计金额=|退还价税合计金额=|总保额变化:)\s*(\S+)', text)
    if refund_total:
        return refund_total.group(2)  # 提取第二个捕获组

    return None  # 如果都没有找到，返回 None

def extract_information(text):
    # 使用正则表达式提取指定字段的内容
    product_name = re.search(r'(产品名称|险种)：\s*(\S+)', text)
    correction_number = re.search(r'(批改序号|批单号码|批 单 号)：\s*(\S+)', text)
    policy_holder = re.search(r'(投保人|被保险人)：\s*(\S+)', text)
    refund_total = extract_refund_total(text)


    return {
        "产品名称": product_name.group(2) if product_name else None,
        "批改序号": correction_number.group(2) if correction_number else None,
        "投保人": policy_holder.group(2) if policy_holder else None,
        "退费合计": refund_total if refund_total else None
    }


def extract_text_from_pdf(pdf_path):
    # 提取PDF文件中的所有文本
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        return text


def process_pdfs_in_directory(directory):
    # 创建一个Excel文件
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["文件名", "产品名称", "批改序号", "投保人", "退费合计"])  # 添加表头

    # 遍历目录中的所有PDF文件
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            print(f"处理文件: {filename}")

            # 从PDF中提取文本
            text = extract_text_from_pdf(pdf_path)

            # 提取所需的信息
            info = extract_information(text)

            # 将文件名和提取的信息写入Excel文件
            sheet.append([filename, info["产品名称"], info["批改序号"], info["投保人"], info["退费合计"]])

    # 将Excel文件保存为与目录中的PDF文件相同名称
    output_path = os.path.join(directory, "提取结果.xlsx")
    workbook.save(output_path)
    print(f"数据已保存到 {output_path}")


# 指定包含PDF文件的目录
pdf_directory = r"D:/提取测试"  # 替换为你实际的PDF目录路径
process_pdfs_in_directory(pdf_directory)
