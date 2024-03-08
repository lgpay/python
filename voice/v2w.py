import imaplib  
import email  
from email.header import decode_header  
import time  
import json  
import requests
import configparser

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
        "touser": WECHAT_USERID,  
        "msgtype": "text",  
        "agentid": WECHAT_AGENTID,  
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
    mail = imaplib.IMAP4_SSL(imap_host)  
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
  
        # 解析邮件标题  
        subject = decode_header(email_message['Subject'])[0][0]  
        if isinstance(subject, bytes):  
            subject = subject.decode()  
  
        # 删除subject中的“收到”和“新”字符  
        cleaned_subject = subject.replace('收到', '').replace('新', '')  
  
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
  
        # 遍历修改后的行列表，找到以“您的账号”开头的行及其后所有行并删除  
        # 创建一个新列表，用于存储需要保留的行  
        filtered_lines = []  
        delete_lines = False  
        for line in body_lines:  
            if line.strip().startswith('您的账号'):  
                delete_lines = True  
            if not delete_lines:  
                filtered_lines.append(line)  
  
        # 将保留的行合并回一个字符串  
        filtered_body = '\n'.join(filtered_lines)    
  
        # 将正文和清理后的标题拼接保存到content变量中  
        content = f"{filtered_body}\n{cleaned_subject}"  
  
        # 打印内容或进行其他处理  
        access_token = get_wechat_access_token()  
        send_wechat_message(access_token, content) 
  
    # 关闭连接  
    mail.close()  
    mail.logout()  
  
# 主循环，每10秒执行一次check_new_emails函数  
while True:  
    check_new_emails()  
    time.sleep(10)  # 等待10秒
