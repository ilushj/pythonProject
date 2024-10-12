import os
import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
import configparser
import ctypes

output_log = "执行完成！ \n"

# 创建一个ConfigParser对象并读取配置文件
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

# 从配置文件中获取必要的设置
file_path = config.get('Settings', 'file_path').replace('\\', '\\\\')
from_email = config.get('Settings', 'from_email').replace('\\', '\\\\')
password = config.get('Settings', 'password')
subject = config.get('Settings', 'subject')
email_text = config.get('Settings', 'text')
directory_path = config.get('Settings', 'directory_path')
cc_emails = config.get('Settings', 'cc_email')
smtp_host = config.get('Settings', 'smtp_host')
smtp_port = config.getint('Settings', 'smtp_port')  # 假设端口是整数

# 读取Excel文件
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active


# 发送邮件的函数
def send_email(to_email, subject, body, attachment_paths):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['CC'] = cc_emails
    msg['Subject'] = subject

    # 添加邮件正文
    msg.attach(MIMEText(body, 'plain'))

    # 添加附件
    for attachment_path in attachment_paths:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{Header(os.path.basename(attachment_path), "utf-8").encode()}"'
            )
            msg.attach(part)

    to_and_cc_emails = [to_email] + [addr.strip() for addr in cc_emails.split(',')]
    server.sendmail(from_email, to_and_cc_emails, msg.as_string())


# 查找附件文件的函数
def find_attachments(keywords):
    attachment_paths = []
    for root, _, files in os.walk(directory_path):
        for file in files:
            if any(keyword in file for keyword in keywords):
                file_path = os.path.join(root, file)
                if file_path not in attachment_paths:
                    attachment_paths.append(file_path)
    return attachment_paths


# 连接SMTP服务器并发送邮件
with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
    server.login(from_email, password)

    # 遍历Excel文件中的每一行
    for row in sheet.iter_rows(min_row=2, values_only=True):
        to_email, keywords_str = row[0], row[1]
        keywords = [k.strip() for k in keywords_str.split(';')]

        # 查找符合条件的附件
        attachment_paths = find_attachments(keywords)
        if not attachment_paths:
            print(f"没有找到包含关键词 '{'; '.join(keywords)}' 的文件，跳过发送邮件。")
            continue

        try:
            email_subject = f"{subject} - {', '.join(keywords)}"
            send_email(to_email, email_subject, email_text, attachment_paths)
            output_log += "邮件发送成功 \n"
            print("邮件发送成功")

            # 删除附件文件
            for attachment_path in attachment_paths:
                os.remove(attachment_path)
                output_log += f"已成功删除文件 '{os.path.basename(attachment_path)}' \n"
                print(f"已成功删除文件 '{os.path.basename(attachment_path)}'")
        except smtplib.SMTPException as e:
            output_log += f"邮件发送失败：{e} \n"
            print(f"邮件发送失败：{e}")

# Your main code
ctypes.windll.user32.MessageBoxW(0, output_log, "Notification", 1)
