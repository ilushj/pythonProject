import requests
import hashlib
import json
import pandas as pd


# 定义登录函数
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

    # 发送POST请求
    response = requests.post(url, data=params)

    # 解析JSON响应
    if response.status_code == 200:
        data = response.json()
        message = data.get('message', '')
        if message == '请求成功':
            data_value = data.get('data', '')
            if data_value:
                usertoken = data_value.split(',')[0]
                return usertoken
        else:
            print(f"登录失败: {message}")
    else:
        print(f"请求失败: {response.status_code}")

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


def callback_func(token, filepath):
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
file_path = r'd:\全量投保\A.xlsx'
result = callback_func(usertoken, file_path)
print(result)

if result == '成功':
    final_result = upload_all(usertoken)
    print(final_result)
