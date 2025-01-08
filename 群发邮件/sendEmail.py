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

output_log = "执行完成！ \n"  # 初始化日志字符串

# 创建一个ConfigParser对象并读取配置文件
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')  # 读取配置文件，指定UTF-8编码

# 从配置文件中获取必要的设置
file_path = config.get('Settings', 'file_path').replace('\\', '\\\\')  # Excel文件路径，替换反斜杠
from_email = config.get('Settings', 'from_email').replace('\\', '\\\\')  # 发件人邮箱，替换反斜杠
password = config.get('Settings', 'password')  # 发件人邮箱密码
subject = config.get('Settings', 'subject')  # 邮件主题
email_text = config.get('Settings', 'text')  # 邮件正文
directory_path = config.get('Settings', 'directory_path')  # 附件目录
cc_emails = config.get('Settings', 'cc_email')  # 抄送邮箱，多个邮箱用逗号分隔
smtp_host = config.get('Settings', 'smtp_host')  # SMTP服务器地址
smtp_port = config.getint('Settings', 'smtp_port')  # SMTP服务器端口，转换为整数

# 读取Excel文件
workbook = openpyxl.load_workbook(file_path)  # 打开Excel文件
sheet = workbook.active  # 获取当前工作表


# 发送邮件的函数
def send_email(to_email, subject, body, attachment_paths):
    global output_log
    msg = MIMEMultipart()  # 创建一个带附件的邮件对象
    msg['From'] = from_email  # 设置发件人
    msg['To'] = to_email  # 设置收件人
    msg['CC'] = cc_emails  # 设置抄送人
    msg['Subject'] = subject  # 设置邮件主题

    # 添加邮件正文
    msg.attach(MIMEText(body, 'plain'))  # 将邮件正文添加到邮件对象中

    # 添加附件
    for attachment_path in attachment_paths:  # 遍历附件路径
        try:
            with open(attachment_path, 'rb') as attachment:  # 以二进制只读模式打开附件
                part = MIMEBase('application', 'octet-stream')  # 创建一个MIMEBase对象
                part.set_payload(attachment.read())  # 读取附件内容
                encoders.encode_base64(part)  # 对附件进行base64编码
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{Header(os.path.basename(attachment_path), "utf-8").encode()}"'
                    # 设置附件的header，解决中文文件名乱码问题
                )
                msg.attach(part)  # 将附件添加到邮件对象中
        except FileNotFoundError:
            print(f"警告：找不到附件文件：{attachment_path}")
            output_log += f"警告：找不到附件文件：{attachment_path}\n"

    to_and_cc_emails = [to_email] + [addr.strip() for addr in cc_emails.split(',')]  # 将收件人和抄送人合并为一个列表
    try:
        server.sendmail(from_email, to_and_cc_emails, msg.as_string())  # 发送邮件
    except smtplib.SMTPSenderRefused as e:
        print(f"发件人地址被拒绝：{e}")
        output_log += f"发件人地址被拒绝：{e}\n"
    except smtplib.SMTPRecipientsRefused as e:
        print(f"收件人地址被拒绝：{e}")
        output_log += f"收件人地址被拒绝：{e}\n"
    except Exception as e:
        print(f"发送邮件时发生其他错误：{e}")
        output_log += f"发送邮件时发生其他错误：{e}\n"


# 查找附件文件的函数
def find_attachments(keywords):
    attachment_paths = []  # 初始化附件路径列表
    for root, _, files in os.walk(directory_path):  # 遍历指定目录及其子目录
        for file in files:  # 遍历每个文件
            if any(keyword in file for keyword in keywords):  # 如果文件名包含任意一个关键词
                file_path = os.path.join(root, file)  # 构建完整的文件路径
                if file_path not in attachment_paths:  # 如果文件路径不在列表中
                    attachment_paths.append(file_path)  # 将文件路径添加到列表中
    return attachment_paths  # 返回附件路径列表


# 连接SMTP服务器并发送邮件
try:
    with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:  # 使用SSL连接SMTP服务器
        server.login(from_email, password)  # 登录邮箱服务器

        # 遍历Excel文件中的每一行，从第二行开始
        for row in sheet.iter_rows(min_row=2, values_only=True):
            to_email, keywords_str = row[0], row[1]  # 获取收件人邮箱和关键词字符串
            if to_email is None or keywords_str is None:
                print("Excel文件中存在空单元格，跳过该行。")
                output_log += "Excel文件中存在空单元格，跳过该行。\n"
                continue
            keywords = [k.strip() for k in keywords_str.split(';')]  # 将关键词字符串按分号分割成列表

            # 查找符合条件的附件
            attachment_paths = find_attachments(keywords)
            if not attachment_paths:  # 如果没有找到附件
                print(f"没有找到包含关键词 '{'; '.join(keywords)}' 的文件，跳过发送邮件。")
                output_log += f"没有找到包含关键词 '{'; '.join(keywords)}' 的文件，跳过发送邮件。\n"
                continue

            try:
                email_subject = f"{subject} - {', '.join(keywords)}"  # 构建邮件主题
                send_email(to_email, email_subject, email_text, attachment_paths)  # 发送邮件
                output_log += f"邮件发送到 {to_email} 成功，主题：{email_subject}\n"
                print(f"邮件发送到 {to_email} 成功，主题：{email_subject}")

                # 删除附件文件
                for attachment_path in attachment_paths:
                    try:
                        os.remove(attachment_path)
                        output_log += f"已成功删除文件 '{os.path.basename(attachment_path)}' \n"
                        print(f"已成功删除文件 '{os.path.basename(attachment_path)}'")
                    except OSError as e:
                        print(f"删除文件 '{os.path.basename(attachment_path)}' 失败：{e}")
                        output_log += f"删除文件 '{os.path.basename(attachment_path)}' 失败：{e}\n"

            except Exception as e:  # 捕获发送邮件过程中的其他异常
                output_log += f"发送邮件到 {to_email} 失败：{e} \n"
                print(f"发送邮件到 {to_email} 失败：{e}")
except smtplib.SMTPAuthenticationError as e:
    print(f"SMTP 认证失败：{e}")
    output_log += f"SMTP 认证失败：{e}\n"
except Exception as e:
    print(f"连接或登录 SMTP 服务器失败：{e}")
    output_log += f"连接或登录 SMTP 服务器失败：{e}\n"

# 显示执行结果的消息框
ctypes.windll.user32.MessageBoxW(0, output_log, "Notification", 1)
