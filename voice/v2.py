import imaplib
import email
from email.header import decode_header
import time
import json
import requests
import configparser
from datetime import datetime
import pytz

# 读取配置文件
config = configparser.ConfigParser()
config.read('config.ini')

# 从配置文件中获取IMAP服务器和登录信息
imap_host = config['imap']['host']
imap_user = config['imap']['user']
imap_pass = config['imap']['password']

# 从配置文件中获取微信企业应用的信息
corpid = config['wechat_enterprise']['corpid']
corpsecret = config['wechat_enterprise']['corpsecret']
agentid = config['wechat_enterprise']['agentid']
touser = config['wechat_enterprise']['touser']

# 获取微信企业应用的access_token
def get_wechat_access_token():
    url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpid}&corpsecret={corpsecret}"
    response = requests.get(url)
    data = response.json()
    return data.get('access_token')

# 发送微信消息
def send_wechat_message(access_token, content):
    url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
    headers = {"Content-Type": "application/json"}
    data = {
        "touser": touser,
        "msgtype": "text",
        "agentid": agentid,
        "text": {
            "content": content
        },
        "safe": 0
    }
    response = requests.post(url, headers=headers, data=json.dumps(data))
    return response.json()

# 检查并处理新邮件
def check_new_emails():
    # 连接到IMAP服务器
    with imaplib.IMAP4_SSL(imap_host) as mail:
        mail.login(imap_user, imap_pass)

        # 选择收件箱
        mail.select('inbox')

        # 搜索新邮件（假设“UNSEEN”表示新邮件）
        result, data = mail.search(None, 'UNSEEN')

        # 遍历所有新邮件的ID
        for num in data[0].split():
            # 获取邮件内容
            _, msg_data = mail.fetch(num, '(RFC822)')
            raw_email = msg_data[0][1].decode('utf-8')
            email_message = email.message_from_string(raw_email)

            # 解析邮件发件人
            sender_name = email.utils.parseaddr(email_message['From'])[0]

            # 解析邮件时间（UTC）
            utc_time = email.utils.parsedate_to_datetime(email_message['Date'])
            # 将UTC时间转换为北京时间（东八区）
            beijing_time = utc_time.astimezone(pytz.timezone('Asia/Shanghai'))
            # 格式化时间，去掉时区信息
            formatted_time = beijing_time.strftime('%Y-%m-%d %H:%M:%S')

            # 解析邮件正文
            body = ''
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        body = part.get_payload(decode=True).decode('utf-8')
                        break
            else:
                body = email_message.get_payload(decode=True).decode('utf-8')

            # 删除前两行
            body_lines = body.splitlines()[2:]

            # 遍历修改后的行列表，找到以“您的账号”或“要回复此短信”开头的行并跳过
            # 创建一个新列表，用于存储需要保留的行
            filtered_lines = []
            skip_lines = False
            for line in body_lines:
                # 如果找到了以“您的账号”开头的行，则设置 skip_lines 为 True
                if line.strip().startswith('您的账号'):
                    skip_lines = True
                # 如果 skip_lines 为 True，则跳过当前行及后续所有行
                if skip_lines:
                    continue
                # 如果行以“要回复此短信”开头，则跳过该行
                if line.strip().startswith('要回复此短信'):
                    continue
                # 否则，将行添加到 filtered_lines 中
                filtered_lines.append(line)

            # 将保留的行合并回一个字符串
            filtered_body = '\n'.join(filtered_lines)

            # 将内容拼接保存到content变量中
            content = f"{sender_name}\n{filtered_body}\n\n{formatted_time}"

            # 通过企业微信应用推送
            access_token = get_wechat_access_token()
            send_wechat_message(access_token, content)

        # 关闭连接
        mail.close()
        mail.logout()

# 主循环，每10秒执行一次check_new_emails函数
while True:
    check_new_emails()
    time.sleep(10)  # 等待10秒
