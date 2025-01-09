import email
import imaplib
import os
import subprocess
import sys
import time
from datetime import datetime
from email.header import decode_header, make_header
import configparser  # 导入 configparser 用于读取配置文件

import schedule

# 读取配置文件
config = configparser.ConfigParser()
config.read('checkmail.ini')

# 邮箱配置
EMAIL = config.get('mail', 'email')  # 从配置文件读取邮箱地址
PASSWORD = config.get('mail', 'password')  # 从配置文件读取邮箱密码
IMAP_SERVER = config.get('mail', 'imap_server')  # 从配置文件读取IMAP服务器地址
IMAP_PORT = config.getint('mail', 'imap_port')  # 从配置文件读取IMAP端口（转换为整数）
SAVE_DIRECTORY = config.get('mail', 'save_directory')  # 从配置文件读取保存目录

# 邮件过滤配置
FROM_ADDRESS = config.get('filter', 'from_address')  # 从配置文件读取发件人地址
SUBJECT_KEYWORD = config.get('filter', 'subject_keyword')  # 从配置文件读取主题关键词


# 监控邮件并保存附件的函数
def check_email():
    try:
        # 登录到邮箱，使用IMAP协议
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL, PASSWORD)
        print("登录成功")
    except Exception as e:
        # 登录失败时捕获异常并打印错误信息
        print(f"登录失败: {e}")
        return

    # 选择“收件箱”文件夹
    mail.select("inbox")

    # 获取今天的日期并格式化为IMAP搜索要求的格式
    today = datetime.today().strftime("%d-%b-%Y")  # IMAP搜索格式为 "day-month-year"，例如 "09-Jan-2025"

    try:
        # 搜索从今天开始的所有邮件
        status, messages = mail.search(None, 'SINCE', today)
    except Exception as e:
        # 搜索邮件时出现问题，捕获异常并打印错误信息
        print(f"搜索邮件时出现问题: {e}")
        return

    # 如果搜索状态不是"OK"，说明出现了问题
    if status != "OK":
        print("搜索邮件时出现问题")
        return

    # 获取所有邮件的ID
    email_ids = messages[0].split()

    # 如果找到了邮件，打印提示信息
    if email_ids:
        print("找到邮件了")

    # 遍历所有匹配的邮件ID
    for email_id in email_ids:
        # 获取邮件内容，使用RFC822格式
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        # 解码邮件的主题
        subject = str(make_header(decode_header(msg["Subject"])))
        try:
            # 获取并解析发件人地址
            from_address_raw = msg.get("From")
            if from_address_raw:
                from_address = email.utils.parseaddr(from_address_raw)[1]
                print(f"From address: {from_address}")
            else:
                print("发件人字段为空")
        except Exception as e:
            # 如果解析发件人地址时发生错误，捕获并打印错误信息
            print(f"解析 '发件人' 地址时出错: {e}")

        # 如果发件人是配置中的指定地址且主题包含配置中的关键词，则继续处理附件
        if from_address == FROM_ADDRESS and SUBJECT_KEYWORD in subject:
            print("找到匹配的邮件")

            # 遍历邮件的所有部分（可能包含多个附件）
            for part in msg.walk():
                # 如果该部分是多部分邮件（例如包含文本和附件），跳过
                if part.get_content_maintype() == "multipart":
                    continue
                # 如果该部分没有附件信息（没有Content-Disposition字段），跳过
                if part.get("Content-Disposition") is None:
                    continue
                # 获取附件文件名
                file_name = part.get_filename()
                if bool(file_name):  # 如果附件有文件名
                    # 处理文件名中的非ASCII字符
                    file_name = str(make_header(decode_header(file_name)))
                    # 生成保存文件的路径和文件名，使用当前日期作为文件名
                    today = datetime.today().strftime("%y%m%d")
                    file_path = os.path.join(SAVE_DIRECTORY, f"{today}.xlsx")

                    # 保存附件到本地指定目录
                    with open(file_path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print(f"附件已保存: {file_path}")

                    # 执行备份程序 fullbackup.exe
                    backup_exe = os.path.join(SAVE_DIRECTORY, "fullbackup.exe")
                    if os.path.exists(backup_exe):  # 如果备份程序存在
                        subprocess.run([backup_exe])  # 执行备份程序
                        print("执行 fullbackup.exe 完成")

                    # 保存附件后退出程序
                    mail.logout()  # 登出邮箱
                    sys.exit()  # 退出程序

    # 如果没有找到任何符合条件的邮件，登出邮箱
    mail.logout()


# 立即执行一次 check_email 函数
check_email()

# 定时任务：每30秒执行一次 check_email 函数
schedule.every(30).seconds.do(check_email)

# 运行定时任务，保持程序持续运行
while True:
    schedule.run_pending()  # 执行已安排的任务
    time.sleep(1)  # 每秒钟检查一次任务
