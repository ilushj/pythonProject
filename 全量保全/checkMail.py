import email
import imaplib
import os
import subprocess
import sys
import time
from datetime import datetime
from email.header import decode_header, make_header

import schedule

# 邮箱配置
EMAIL = "wangxy@pagzb.com"
PASSWORD = "Wxy1799170"
IMAP_SERVER = "imap.exmail.qq.com"
IMAP_PORT = 993
SAVE_DIRECTORY = r"d:\全量投保"


# 监控邮件并保存附件
def check_email():
    try:
        # 登录到邮箱
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL, PASSWORD)
        print("登录成功")
    except Exception as e:
        print(f"登录失败: {e}")
        return

    mail.select("inbox")

    # 获取昨天的日期
    today = datetime.today().strftime("%d-%b-%Y")  # 修改为适合IMAP搜索的日期格式

    try:
        # 搜索昨天以来的所有邮件
        status, messages = mail.search(None, 'SINCE', today)
    except Exception as e:
        print(f"搜索邮件时出现问题: {e}")
        return

    if status != "OK":
        print("搜索邮件时出现问题")
        return

    email_ids = messages[0].split()

    if email_ids:
        print("找到邮件了")

    # 遍历所有匹配的邮件
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        # 解码邮件主题
        subject = str(make_header(decode_header(msg["Subject"])))
        try:
            # 检查发件人和主题
            from_address_raw = msg.get("From")  # 或者 msg.get("From")
            if from_address_raw:
                from_address = email.utils.parseaddr(from_address_raw)[1]
                print(f"From address: {from_address}")
            else:
                print("发件人字段为空")
        except Exception as e:
            print(f"解析 '发件人' 地址时出错: {e}")
        if from_address == "wangxy0127@foxmail.com" and "名单报备" in subject:
            print("找到匹配的邮件")

            # 检查邮件是否有附件
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue
                file_name = part.get_filename()
                if bool(file_name):
                    # 处理文件名中的非ASCII字符
                    file_name = str(make_header(decode_header(file_name)))
                    # 生成保存文件的路径和文件名
                    today = datetime.today().strftime("%y%m%d")
                    file_path = os.path.join(SAVE_DIRECTORY, f"{today}.xlsx")
                    with open(file_path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print(f"附件已保存: {file_path}")

                    # 执行 fullbackup.exe
                    backup_exe = os.path.join(SAVE_DIRECTORY, "fullbackup.exe")
                    if os.path.exists(backup_exe):
                        subprocess.run([backup_exe])
                        print("执行 fullbackup.exe 完成")

                    # 保存成功后退出程序
                    mail.logout()
                    sys.exit()

    mail.logout()


# 立即执行一次
check_email()

# 定时任务
schedule.every(30).seconds.do(check_email)

# 运行定时任务
while True:
    schedule.run_pending()
    time.sleep(1)
