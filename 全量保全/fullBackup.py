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
    # 登录API的URL
    url = 'https://ekangonline.com/kaimai/user/login'

    # 方案与账户信息的映射
    accounts = {
        '花名册3010': ('shmj3010', '0551'),
        '花名册6010': ('shmjqy6010', '114003'),
        '花名册8010': ('shmj8010', '0551'),
        '花名册10010': ('shmj10010', '0551')
    }

    # 根据传入的plan选择对应的账户信息
    if plan in accounts:
        username, raw_password = accounts[plan]
        # 密码使用MD5加密
        password = hashlib.md5(raw_password.encode()).hexdigest()
    else:
        print(f"未找到方案 {plan} 的账户信息")
        return None

    client_type = 4  # 客户端类型
    # 请求参数，包括用户名、密码和客户端类型
    params = {
        'clientType': client_type,
        'username': username,
        'password': password
    }
    attempts = 3  # 登录重试次数

    # 尝试登录，最多尝试3次
    for attempt in range(attempts):
        response = requests.post(url, data=params)
        if response.status_code == 200:
            data = response.json()
            message = data.get('message', '')
            if message == '请求成功':
                # 获取返回的token
                data_value = data.get('data', '')
                if data_value:
                    user_token = data_value.split(',')[0]
                    return user_token  # 返回token
            else:
                print(f"登录失败: {message}")
        else:
            print(f"请求失败: {response.status_code}")

        # 如果未成功且未到最大重试次数，等待5秒后重试
        if attempt < attempts - 1:
            time.sleep(5)

    return None  # 登录失败，返回None


# 数据上传及处理函数
def callback_func(token, filepath, savepath):
    # 数据上传的API URL
    callback_url = 'https://ekangonline.com/kaimai/importrecord/uploadAllCheck'

    # 读取Excel文件中的数据
    df = pd.read_excel(filepath)
    headers = {'Authorization': token}  # 使用token设置认证头
    # 遍历DataFrame中的每一行，格式化为指定的JSON格式
    employee_list = [format_row_data(row) for index, row in df.iterrows()]

    # 构造请求的payload
    payload = {"employeeList": json.dumps(employee_list, ensure_ascii=False)}
    # 构造文件请求
    files = {
        'file': (filepath, open(filepath, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    }

    # 发送POST请求上传数据
    response = requests.post(callback_url, headers=headers, data=payload, files=files)
    print(f"Response: {response.text}")
    response_data = response.json()

    if response_data.get('message') == '请求成功':
        return '成功'  # 数据上传成功

    # 如果请求失败，处理错误记录并保存为Excel
    if response_data.get('message') == '请求失败':
        error_records = response_data.get('data', [])
        wb = Workbook()
        ws = wb.active
        ws.append(["idcard", "name", "insuredCustomer", "importMessage"])

        # 将错误记录写入Excel文件
        for record in error_records:
            ws.append(
                [record.get('idcard'), record.get('name'), record.get('insuredCustomer'), record.get('importMessage')])

        wb.save(savepath)  # 保存错误记录Excel文件
        print(f"错误记录已保存至: {savepath}")

        # 获取所有错误记录的身份证号码
        error_idcards = {record.get('idcard') for record in error_records}
        initial_length = len(employee_list)
        # 删除包含错误ID的记录
        employee_list = [employee for employee in employee_list if employee.get('col2') not in error_idcards]
        final_length = len(employee_list)

        print(f"Removed {initial_length - final_length} erroneous records.")

        # 重新上传去除错误记录的数据
        payload = {"employeeList": json.dumps(employee_list, ensure_ascii=False)}
        response = requests.post(callback_url, headers=headers, data=payload, files=files)
        print(f"Response: {response.text}")
        response_data = response.json()

        if response_data.get('message') == '请求成功':
            return '成功附带错误'  # 上传成功，但附带错误记录

    return '检测处理失败'  # 数据上传失败


# 上传所有数据的函数
def upload_all(auth_value):
    # 上传所有数据的API URL
    url = 'https://ekangonline.com/kaimai/importrecord/uploadAll'
    headers = {'Authorization': auth_value}

    # 发送请求上传所有数据
    response = requests.post(url, headers=headers)
    print(f"Upload All Response: {response.text}")
    response_data = response.json()

    if response_data.get('message') == '请求成功':
        return '成功'  # 上传成功
    return '投保处理失败'  # 上传失败


# 格式化每行数据为规定的JSON格式
def format_row_data(row):
    formatted_data = {
        'colCount': 6,
        'col1': row['姓名'],
        'col2': row['身份证'],
        'col3': row['用工单位'],
        'col4': row['工种'],
        'col5': row['职业类别'],
        'col6': row['雇主单位']
    }
    return formatted_data


# 邮件发送函数
def send_error_email(filepath, subject, body):
    sender_email = "wangxy@pagzb.com"  # 发送者邮箱地址
    sender_name = "易久保系统"  # 发送者名称
    receiver_email = "wangxy@pagzb.com"  # 收件者邮箱地址
    msg = EmailMessage()
    msg['From'] = formataddr((sender_name, sender_email))  # 设置发送者地址
    msg['To'] = receiver_email  # 设置收件者地址
    msg['Subject'] = subject  # 设置邮件主题
    msg.set_content(body)  # 设置邮件正文内容

    # 如果附件文件存在，读取并添加为附件
    if os.path.exists(filepath):
        with open(filepath, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(filepath)
            msg.add_attachment(file_data, maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    try:
        # 使用SMTP_SSL进行SSL加密连接，发送邮件
        with smtplib.SMTP_SSL('smtp.exmail.qq.com', 465) as server:
            server.login(sender_email, 'Wxy1799170')  # 登录SMTP服务器
            server.send_message(msg)  # 发送邮件
        print("邮件发送成功")
    except Exception as e:
        print(f"邮件发送失败: {e}")  # 捕获并打印邮件发送异常


# 主程序部分
# 获取当天日期 格式为YYMMDD
current_date = datetime.now().strftime("%y%m%d")
file_path = r'd:\全量投保'  # 文件存储路径
source_filename = f'{current_date}.xlsx'  # 源文件名，使用当天日期命名
source_filepath = os.path.join(file_path, source_filename)  # 源文件路径
save_path = r'd:\全量投保\错误数据'  # 错误数据存储路径

# 读取Excel文件中的数据
df = pd.read_excel(source_filepath)
plans = df['方案'].unique()  # 获取所有方案的唯一值
results = []  # 用于存储各个方案的处理结果

# 遍历每个方案
for plan in plans:
    print(f"正在处理方案: {plan}")
    usertoken = login(plan)  # 登录并获取token
    if usertoken:
        print(f"登录成功，usertoken: {usertoken}")
    else:
        print(f"登录失败，方案: {plan}")
        continue  # 登录失败则跳过当前方案

    # 筛选当前方案的数据
    df_plan = df[df['方案'] == plan]

    # 临时保存当前方案的数据到一个新的Excel文件
    temp_filename = f'temp_{plan}_{current_date}.xlsx'
    temp_filepath = os.path.join(file_path, temp_filename)
    df_plan.to_excel(temp_filepath, index=False)

    # 错误数据保存路径
    error_filename = f'错误_{plan}_{current_date}.xlsx'
    error_filepath = os.path.join(save_path, error_filename)

    # 上传数据并处理返回结果
    result = callback_func(usertoken, temp_filepath, error_filepath)
    print(result)
    results.append((plan, result))

    # 根据处理结果进行后续操作
    if result == '成功':
        final_result = upload_all(usertoken)  # 上传所有数据
        print(final_result)
    elif result == '成功附带错误':
        final_result = upload_all(usertoken)  # 上传所有数据
        send_error_email(error_filepath, f'{current_date}_{plan}_错误记录', "请查看附件中的错误记录。")  # 发送错误邮件
        print(final_result)
    else:
        send_error_email(error_filepath, f'{current_date}_{plan}_全员投保失败', "全员投保失败，请手动操作！")  # 发送投保失败邮件

    # 如果最终结果是上传失败，发送投保失败邮件
    if final_result == "投保处理失败":
        send_error_email(error_filepath, "投保意外失败", "全员投保失败，请手动操作！")
    else:
        message = f'{current_date}_{plan}_投保成功'
        send_error_email(error_filepath, message, "投保成功！")  # 发送成功邮件

# 输出每个方案的最终处理结果
for plan, result in results:
    print(f"方案: {plan}, 结果: {result}")
