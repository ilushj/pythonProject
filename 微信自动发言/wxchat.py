import requests

# 企业微信群机器人 Webhook URL
webhook_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_WEBHOOK_KEY"

# 发送消息函数
def send_message_to_wechat(message):
    headers = {"Content-Type": "application/json"}
    data = {
        "msgtype": "text",
        "text": {
            "content": message,
        }
    }
    response = requests.post(webhook_url, json=data, headers=headers)
    if response.status_code == 200:
        print("消息发送成功！")
    else:
        print(f"消息发送失败：{response.status_code}, {response.text}")

# 调用函数发送消息
send_message_to_wechat("吃货们，今天有什么推荐的美食吗？😋")
