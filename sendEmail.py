import os
import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
import openpyxl

# 读取配置文件
config_path = 'D:\\群发email\\config.ini'
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
    file_name = row[1]+".xlsx"

    # 检查文件是否存在
    attachment_path = r'D:\群发email\{}'.format(file_name)  # 绝对路径
    if not os.path.exists(attachment_path):
        print(f"文件 '{file_name}' 不存在，跳过发送邮件。")
        continue

    # 构建邮件
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = Subject

    # 添加文本消息
    text = text
    msg.attach(MIMEText(text, 'plain'))

    # 添加附件
    with open(attachment_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        encoded_filename = Header(file_name, 'utf-8').encode()
        part.add_header('Content-Disposition', f'attachment; filename="{encoded_filename}"')
        msg.attach(part)

    # 发送邮件
    server.sendmail(from_email, to_email, msg.as_string())

    # 删除附件文件
    os.remove(attachment_path)
    print(f"已成功发送邮件并删除文件 '{file_name}'.")

# 关闭连接
server.quit()
