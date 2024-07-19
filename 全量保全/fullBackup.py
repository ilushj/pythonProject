import os
import requests
import hashlib
import json
import pandas as pd
import time
from openpyxl import Workbook
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from datetime import datetime


def send_error_email(filepath, subject, body):
    # 邮件配置
    sender_email = "ilushj@hotmail.com"
    sender_name = "易久保系统"
    receiver_email = "ilushj@hotmail.com"

    # 创建邮件
    msg = EmailMessage()
    msg['From'] = formataddr((sender_name, sender_email))
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.set_content(body)

    # 检查文件是否存在，如果存在则添加附件
    if os.path.exists(filepath):
        with open(filepath, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(filepath)
            msg.add_attachment(file_data, maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    # 发送邮件
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender_email, '771119+_')
            server.send_message(msg)
        print("邮件发送成功")
    except Exception as e:
        print(f"邮件发送失败: {e}")


# 登录函数
def login():
    url = 'https://test.ekangonline.com/kaimai/user/login'
    username = 'test21'
    password = hashlib.md5("123456".encode()).hexdigest()
    client_type = 4

    # 构造请求参数
    params = {
        'clientType': client_type,
        'username': username,
        'password': password
    }
    # 尝试登录的次数
    attempts = 3

    for attempt in range(attempts):
        # 发送POST请求
        response = requests.post(url, data=params)

        # 解析JSON响应
        if response.status_code == 200:
            data = response.json()
            message = data.get('message', '')
            if message == '请求成功':
                data_value = data.get('data', '')
                if data_value:
                    user_token = data_value.split(',')[0]
                    return user_token
            else:
                print(f"登录失败: {message}")
        else:
            print(f"请求失败: {response.status_code}")

        # 如果登录失败且还没有达到最大尝试次数，等待5秒后重试
        if attempt < attempts - 1:
            time.sleep(5)

    return None


# 格式化每行数据为规定的json格式
def format_row_data(row):
    formatted_data = {
        'colCount': 7,
        'col1': row['姓名'],
        'col2': row['身份证'],
        'col3': row['职业类别'],
        'col4': row['工种'],
        'col5': row['用工单位'],
        'col6': row['雇主单位'],
        'col7': "验证通过"
        # Add more fields as needed
    }
    return formatted_data


def callback_func(token, filepath, savepath):
    callback_url = 'https://test.ekangonline.com/kaimai/importrecord/uploadAllCheck'

    # 读取excel文件数据
    df = pd.read_excel(filepath)

    headers = {
        'Authorization': token,
    }

    employee_list = []

    for index, row in df.iterrows():
        data = format_row_data(row)
        employee_list.append(data)

    payload = {
        "employeeList": json.dumps(employee_list, ensure_ascii=False)
    }

    files = {
        'file': (file_path, open(file_path, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    }

    # 打印整理好的数据
    print(f"Payload: {payload}")

    response = requests.post(callback_url, headers=headers, data=payload, files=files)
    # 打印返回值
    print(f"Response: {response.text}")

    # 解析响应并检查 message 字段
    response_data = response.json()
    if response_data.get('message') == '请求成功':
        return '成功'

    if response_data.get('message') == '请求失败':
        error_records = response_data.get('data', [])

        # 生成一个"错误.xlsx"存放错误记录，并保存在d:\全量投保\目录中
        wb = Workbook()
        ws = wb.active
        ws.append(["idcard", "name", "insuredCustomer", "importMessage"])  # 假设我们只关心这两个字段

        for record in error_records:
            ws.append(
                [record.get('idcard'), record.get('name'), record.get('insuredCustomer'), record.get('importMessage')])

        wb.save(savepath)
        print(f"错误记录已保存至: {error_filepath}")

        # 从employee_list中删除 'col2' = “idcard”的记录
        error_idcards = {record.get('idcard') for record in error_records}
        initial_length = len(employee_list)
        employee_list = [employee for employee in employee_list if employee.get('col2') not in error_idcards]
        final_length = len(employee_list)

        print(f"Removed {initial_length - final_length} erroneous records.")

        # 更新payload
        payload = {
            "employeeList": json.dumps(employee_list, ensure_ascii=False)
        }
        print(f"Payload: {payload}")

        # 再次post
        response = requests.post(callback_url, headers=headers, data=payload, files=files)
        print(f"Response: {response.text}")

        response_data = response.json()
        if response_data.get('message') == '请求成功':
            return '成功'

    return '处理失败'


def upload_all(auth_value):
    url = 'https://test.ekangonline.com/kaimai/importrecord/uploadAll'
    headers = {
        'Authorization': auth_value
    }

    response = requests.post(url, headers=headers)

    # 打印返回值
    print(f"Upload All Response: {response.text}")

    # 解析响应并检查是否成功
    response_data = response.json()
    if response_data.get('message') == '请求成功':
        return '成功'

    return '处理失败'


# 执行登录函数
usertoken = login()
if usertoken:
    print(f"登录成功，usertoken: {usertoken}")
else:
    print("登录失败")
current_date = datetime.now().strftime("%y%m%d")
file_path = r'd:\全量投保'
source_filename = f'{current_date}.xlsx'
source_filepath = os.path.join(file_path, source_filename)
save_path = r'd:\全量投保\错误数据'
error_filename = f'错误_{current_date}.xlsx'
error_filepath = os.path.join(save_path, error_filename)
result = callback_func(usertoken, source_filepath, error_filepath)
print(result)

if result == '成功':
    final_result = upload_all(usertoken)
    send_error_email(error_filepath, "错误记录", "请查看附件中的错误记录。")
    print(final_result)
    if final_result != "成功":
        send_error_email(error_filepath, "全员投保失败", "全员投保失败，请手动操作！")
else:
    send_error_email(error_filepath, "全员投保失败", "全员投保失败，请手动操作！")
