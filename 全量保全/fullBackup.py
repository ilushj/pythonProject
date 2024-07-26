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


# 登录函数
def login(plan):
    url = 'https://ekangonline.com/kaimai/user/login'

    # 方案与账户信息映射
    accounts = {
        '花名册3010': ('shmj3010', '0551'),
        '花名册6010': ('shmjqy6010', '114003'),
        '花名册8010': ('shmj8010', '0551'),
        '花名册10010': ('shmj10010', '0551')
    }

    if plan in accounts:
        username, raw_password = accounts[plan]
        password = hashlib.md5(raw_password.encode()).hexdigest()
    else:
        print(f"未找到方案 {plan} 的账户信息")
        return None

    client_type = 4
    params = {
        'clientType': client_type,
        'username': username,
        'password': password
    }
    attempts = 3

    for attempt in range(attempts):
        response = requests.post(url, data=params)
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
        if attempt < attempts - 1:
            time.sleep(5)

    return None


# 数据上传及处理函数
def callback_func(token, filepath, savepath):
    callback_url = 'https://ekangonline.com/kaimai/importrecord/uploadAllCheck'
    df = pd.read_excel(filepath)
    headers = {'Authorization': token}
    employee_list = [format_row_data(row) for index, row in df.iterrows()]

    payload = {"employeeList": json.dumps(employee_list, ensure_ascii=False)}
    files = {
        'file': (filepath, open(filepath, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}

    response = requests.post(callback_url, headers=headers, data=payload, files=files)
    print(f"Response: {response.text}")
    response_data = response.json()

    if response_data.get('message') == '请求成功':
        return '成功'

    if response_data.get('message') == '请求失败':
        error_records = response_data.get('data', [])
        wb = Workbook()
        ws = wb.active
        ws.append(["idcard", "name", "insuredCustomer", "importMessage"])

        for record in error_records:
            ws.append(
                [record.get('idcard'), record.get('name'), record.get('insuredCustomer'), record.get('importMessage')])

        wb.save(savepath)
        print(f"错误记录已保存至: {savepath}")

        error_idcards = {record.get('idcard') for record in error_records}
        initial_length = len(employee_list)
        employee_list = [employee for employee in employee_list if employee.get('col2') not in error_idcards]
        final_length = len(employee_list)

        print(f"Removed {initial_length - final_length} erroneous records.")
        payload = {"employeeList": json.dumps(employee_list, ensure_ascii=False)}
        # print(f"Payload: {payload}")

        response = requests.post(callback_url, headers=headers, data=payload, files=files)
        print(f"Response: {response.text}")
        response_data = response.json()
        if response_data.get('message') == '请求成功':
            return '成功附带错误'

    return '检测处理失败'


# 上传所有数据函数
def upload_all(auth_value):
    url = 'https://ekangonline.com/kaimai/importrecord/uploadAll'
    headers = {'Authorization': auth_value}
    response = requests.post(url, headers=headers)
    print(f"Upload All Response: {response.text}")
    response_data = response.json()
    if response_data.get('message') == '请求成功':
        return '成功'
    return '投保处理失败'


# 格式化每行数据为规定的JSON格式
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
    }
    return formatted_data


# 邮件发送函数
def send_error_email(filepath, subject, body):
    sender_email = "wangxy@pagzb.com"
    sender_name = "易久保系统"
    receiver_email = "wangxy@pagzb.com"
    msg = EmailMessage()
    msg['From'] = formataddr((sender_name, sender_email))
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.set_content(body)

    if os.path.exists(filepath):
        with open(filepath, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(filepath)
            msg.add_attachment(file_data, maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    try:
        # 使用 SMTP_SSL 进行 SSL 连接
        with smtplib.SMTP_SSL('smtp.exmail.qq.com', 465) as server:
            server.login(sender_email, 'Wxy1799170')  # 确保密码正确
            server.send_message(msg)  
        print("邮件发送成功")
    except Exception as e:
        print(f"邮件发送失败: {e}")


# 主程序部分
# 获取当天日期 格式为YYMMDD

current_date = datetime.now().strftime("%y%m%d")
file_path = r'd:\全量投保'
source_filename = f'{current_date}.xlsx'
source_filepath = os.path.join(file_path, source_filename)
save_path = r'd:\全量投保\错误数据'

# 读取Excel文件
df = pd.read_excel(source_filepath)
plans = df['方案'].unique()
results = []

for plan in plans:
    print(f"正在处理方案: {plan}")
    usertoken = login(plan)
    if usertoken:
        print(f"登录成功，usertoken: {usertoken}")
    else:
        print(f"登录失败，方案: {plan}")
        continue

    # 筛选当前方案的数据
    df_plan = df[df['方案'] == plan]

    # 临时保存当前方案的数据到一个新的Excel文件
    temp_filename = f'temp_{plan}_{current_date}.xlsx'
    temp_filepath = os.path.join(file_path, temp_filename)
    df_plan.to_excel(temp_filepath, index=False)

    error_filename = f'错误_{plan}_{current_date}.xlsx'
    error_filepath = os.path.join(save_path, error_filename)

    result = callback_func(usertoken, temp_filepath, error_filepath)
    print(result)
    results.append((plan, result))

    if result == '成功':
        final_result = upload_all(usertoken)
        print(final_result)
    elif result == '成功附带错误':
        final_result = upload_all(usertoken)
        send_error_email(error_filepath, f'{current_date}_{plan}_错误记录', "请查看附件中的错误记录。")
        print(final_result)
    else:
        send_error_email(error_filepath, "全员投保失败", "全员投保失败，请手动操作！")

    if final_result == "投保处理失败":
        send_error_email(error_filepath, "投保意外失败", "全员投保失败，请手动操作！")
    else:
        message = f'{current_date}_{plan}_投保成功'
        send_error_email(error_filepath, message, "投保成功！")

for plan, result in results:
    print(f"方案: {plan}, 结果: {result}")
