import imaplib
import email
from email.header import decode_header, make_header
import os
import schedule
import time
from datetime import datetime, timedelta
import sys
import subprocess

# 邮箱配置
EMAIL = "ilushj@hotmail.com"
PASSWORD = "771119+_"
IMAP_SERVER = "imap-mail.outlook.com"
IMAP_PORT = 993
SAVE_DIRECTORY = r"d:\千伏"


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
    yesterday = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%Y")

    try:
        # 搜索昨天以来的所有邮件
        status, messages = mail.search(None, 'SINCE', yesterday)
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

        # 检查发件人和主题
        from_address = email.utils.parseaddr(msg.get("From"))[1]
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
