import os
import re
import pdfplumber


def extract_info_from_pdf(pdf_path):
    """
    使用 pdfplumber 从 PDF 中提取“购 名称”和“价税合计(大写)”中的信息。
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text

        # 打印调试信息，查看提取的文本
        print(f"Extracted text from {os.path.basename(pdf_path)}:\n{text}\n{'-' * 50}")

        # 清理文本，去除多余的换行和空格
        text = text.replace('\n', '').replace(' ', '')

        # 提取“购 名称”
        name_match = re.search(r"购名称[:：]\s*([\u4e00-\u9fa5A-Za-z0-9]+)", text)
        name = name_match.group(1).strip() if name_match else None

        # 提取“价税合计(大写)”后的金额
        amount_match = re.search(r"价税合计\(大写\).*?（小写）¥([\d.]+)", text)
        amount = amount_match.group(1).strip() if amount_match else None

        return name, amount
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return None, None


def rename_pdf_files_in_directory(directory):
    """
    遍历目录中的所有 PDF 文件，提取“购 名称”和“价税合计(大写)”的金额，并修改文件名。
    """
    for filename in os.listdir(directory):
        if filename.lower().endswith('.pdf'):
            file_path = os.path.join(directory, filename)

            # 提取购方名称和金额
            buyer_name, amount = extract_info_from_pdf(file_path)

            if buyer_name and amount:
                # 构造新的文件名（确保名称合法）
                sanitized_name = re.sub(r'[\\/*?:"<>|]', "", buyer_name)
                new_filename = f"{sanitized_name}_{amount}.pdf"
                new_file_path = os.path.join(directory, new_filename)

                # 重命名文件
                try:
                    os.rename(file_path, new_file_path)
                    print(f"Renamed: {filename} -> {new_filename}")
                except Exception as e:
                    print(f"Error renaming {filename}: {e}")
            else:
                print(f"Could not extract name or amount from {filename}")


# 运行函数
directory_path = r"D:\发票"
rename_pdf_files_in_directory(directory_path)
