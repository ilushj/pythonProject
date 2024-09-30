import os
import smtplib
import time
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
import openpyxl

# 读取配置文件
config_path = os.path.join(os.getcwd(), 'config.ini')
with open(config_path, 'r', encoding='utf-8') as config_file:
    try:
        config = json.load(config_file)
    except json.JSONDecodeError as e:
        print(f"JSON 解析错误: {e}")
    except FileNotFoundError as e:
        print(f"找不到文件: {e}")
    except Exception as e:
        print(f"发生未知错误: {e}")

# 从配置文件中获取
file_path = config.get('file_path')
from_email = config.get('from_email')
password = config.get('password')
Subject = config.get('subject')
text = config.get('text')
directory_path = config.get('directory_path')
cc_emails = config.get('cc_email')

# file_path = r'D:\佣金明细\email\通讯录.xlsx'
# 读取Excel文件
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# 设置发件人邮箱和密码
# from_email = 'customer13@qidianbx.com'
# password = 'Qd6657bx'

# 连接SMTP服务器
server = smtplib.SMTP_SSL('smtp.qiye.aliyun.com', 465)
server.login(from_email, password)

# 遍历Excel文件中的每一行
for row in sheet.iter_rows(min_row=2, values_only=True):
    to_email = row[0]
    # 获取关键词并拆分为列表
    keywords = row[1].split(';')  # 根据分号分割并去除前后空格
    keywords = [keyword.strip() for keyword in keywords]

    # 查找包含关键词的所有文件
    attachment_paths = []
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if any(keyword in file for keyword in keywords):  # 检查文件名是否包含任一关键词
                attachment_paths.append(os.path.join(root, file))

    if not attachment_paths:
        print(f"没有找到包含关键词 '{'; '.join(keywords)}' 的文件，跳过发送邮件。")
        continue

    # 构建邮件
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['CC'] = cc_emails   # 抄送地址，可以是单个地址或以逗号分隔的多个地址
    msg['Subject'] = f"{Subject} - {', '.join(keywords)}"

    # 添加文本消息
    text = text
    msg.attach(MIMEText(text, 'plain'))

    # 添加所有找到的附件
    for attachment_path in attachment_paths:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            encoded_filename = Header(os.path.basename(attachment_path), 'utf-8').encode()
            part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
            msg.attach(part)

    try:
        to_and_cc_emails = [msg['To']] + [addr.strip() for addr in msg['CC'].split(',')]
        server.sendmail(from_email, to_and_cc_emails, msg.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException as e:
        print(f"邮件发送失败：{e}")
    else:
        print("邮件发送成功")

        # 删除所有附件文件
    for attachment_path in attachment_paths:
        os.remove(attachment_path)
        print(f"已成功删除文件 '{os.path.basename(attachment_path)}'.")


# 关闭连接
server.quit()
