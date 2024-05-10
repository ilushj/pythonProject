import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import openpyxl
import os

file_path = r'D:\佣金确认\email\通讯录.xlsx'
# 读取Excel文件
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# 设置发件人邮箱和密码
from_email = 'shihj@qidianbx.com'
password = 'Ilushj771119'

# 连接SMTP服务器
server = smtplib.SMTP_SSL('smtp.qiye.aliyun.com', 465)
server.login(from_email, password)

# 遍历Excel文件中的每一行
for row in sheet.iter_rows(min_row=2, values_only=True):
    to_email = row[0]
    file_name = row[1]+".xlsx"

    # 检查文件是否存在
    attachment_path = r'D:\佣金确认\email\{}'.format(file_name)  # 绝对路径
    if not os.path.exists(attachment_path):
        print(f"文件 '{file_name}' 不存在，跳过发送邮件。")
        continue

    # 构建邮件
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = file_name + '佣金确认'

    # 添加文本消息
    text = "您好，附件中是您的佣金及业绩，请查收。（此邮件不要回复！）"
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
