import os
import re
import shutil
import PyPDF2

# 根据PDF内容重命名文件
def rename_and_copy_pdf_files(directory):
    # 确保目标目录 "ex" 存在，如果不存在则创建它
    target_directory = os.path.join(directory, 'ex')
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    # 遍历目录中的所有文件
    for filename in os.listdir(directory):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(directory, filename)

            # 打开 PDF 文件并读取其内容
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""

                # 提取所有页面的文本
                for page in reader.pages:
                    text += page.extract_text()
                # 打印提取的文字（去除空格后的内容）
                # print(f"Extracted text from {filename} (cleaned):")
                # print(text)
                # 检查是否包含“被保险人变动批单”或“被保险人变动清单”
                if "被保险人变动批单" in text or "同期增减人批单" in text:
                    voucher_number = extract_voucher_number(text)
                    application_date = extract_date(text, "业务申请日期")
                    if voucher_number:
                        new_name = f"{voucher_number}-被保险人变动批单-{application_date}.pdf"
                        new_path = get_unique_filepath(target_directory, new_name)
                        copy_file(pdf_path, new_path)

                elif "被保险人变动清单" in text or "同期增减人清单" in text:
                    voucher_number = extract_voucher_number(text)
                    reception_date = extract_date(text, "受理日期")
                    if voucher_number:
                        new_name = f"{voucher_number}-被保险人变动清单-{reception_date}.pdf"
                        new_path = get_unique_filepath(target_directory, new_name)
                        copy_file(pdf_path, new_path)


def extract_voucher_number(text):
    # 找到“业务凭证号”后面的内容
    match = re.search(r"业务凭证号：\s*([A-Za-z0-9]+)", text)
    if match:
        # 提取业务凭证号
        voucher_number = match.group(1)
        # 获取最后9个字符
        return voucher_number[-9:]
    return None

def extract_date(text, date_label):
    # 根据标签提取日期（格式：yyyy年mm月dd日）
    match = re.search(rf"{date_label}?：\s*(\d{{4}}年\d{{2}}月\d{{2}}日)", text)
    if match:
        # 提取日期并转换成 YYYYMMDD 格式
        date_str = match.group(1)
        return date_str.replace("年", "").replace("月", "").replace("日", "")
    return None


def get_unique_filepath(directory, filename):
    # 确保文件名唯一，如果存在重复文件，加入流水号
    base, ext = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base}-{counter}{ext}"
        counter += 1
    return os.path.join(directory, new_filename)


def copy_file(src, dst):
    try:
        # 复制文件到目标目录
        shutil.copy2(src, dst)
        print(f"Copied: {src} to {dst}")
    except Exception as e:
        print(f"Failed to copy {src} to {dst}. Error: {e}")


# 从用户输入获取目录路径
directory = input("请输入PDF文件所在目录路径: ").strip()

# 检查目录是否有效
if os.path.isdir(directory):
    rename_and_copy_pdf_files(directory)
else:
    print("提供的目录路径无效，请检查并重新输入。")
