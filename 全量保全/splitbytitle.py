import pdfplumber
from PyPDF2 import PdfWriter, PdfReader
import re


def extract_name(text):
    # 使用正则表达式提取“被保险人姓名：”后的名字
    match = re.search(r'被保险人姓名：(\S+)', text)
    if match:
        return match.group(1)
    return None


def split_pdf_by_title(input_pdf, output_prefix, title="个人保险凭证"):
    with pdfplumber.open(input_pdf) as pdf:
        current_pdf_writer = PdfWriter()
        document_count = 1
        reader = PdfReader(input_pdf)
        name = None

        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()

            # 检查每个条件
            title_check = title in text.splitlines()[0].strip()
            pages_check = len(current_pdf_writer.pages) > 0

            # 打印每个条件的结果
            print(f"文本内容存在: {bool(text)}")
            print(f"标题检查: {title_check}")
            print(f"当前文档页面数大于0: {pages_check}")

            # 检查是否在页面的第一行出现了“个人保险凭证”
            if text and title_check and len(current_pdf_writer.pages) > 0:
                # 在遇到新的关键词前保存当前文档
                print(f"保存名称：{name}")
                output_filename = f"{output_prefix}_{document_count}_{name}.pdf"
                with open(output_filename, "wb") as out_f:
                    current_pdf_writer.write(out_f)
                document_count += 1

                # 初始化下一个文档的 writer
                current_pdf_writer = PdfWriter()
            if title_check:
                name = extract_name(text)
                if name:
                    print(f"提取到的姓名：{name}")
            # 无论是否匹配关键词，当前页面都应添加到当前的 writer 中
            current_pdf_writer.add_page(reader.pages[page_num])

        # 保存最后一个文档
        if len(current_pdf_writer.pages) > 0:
            output_filename = f"{output_prefix}_{document_count}_{name}.pdf"
            with open(output_filename, "wb") as out_f:
                current_pdf_writer.write(out_f)

    print(f"PDF 文件已根据 '{title}' 分割为 {document_count} 个文件。")


# 调用示例
input_pdf = 'd:/splitpdf/家属个人凭证测试pdf.pdf'
output_prefix = 'd:/splitpdf/11/'
split_pdf_by_title(input_pdf, output_prefix)
