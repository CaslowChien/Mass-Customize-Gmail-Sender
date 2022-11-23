from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

def send(send_subject, send_from, send_smtp, send_to, send_content):
    content = MIMEMultipart()  #建立MIMEMultipart物件
    content["subject"] = send_subject  #郵件標題
    content["from"] = send_from  #寄件者
    content["to"] = send_to #收件者
    content.attach(MIMEText(send_content))  #郵件內容

    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
        try:
            smtp.ehlo()  # 驗證SMTP伺服器
            smtp.starttls()  # 建立加密傳輸
            smtp.login(send_from, send_smtp)  # 登入寄件者gmail
            smtp.send_message(content)  # 寄送郵件
            return True
        except Exception as e: # mail格式不對
            return False