import smtplib
import os
import html
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
from datetime import datetime

# 加载环境变量（使用脚本所在目录的.env文件）
script_dir = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(script_dir, '.env'))

def send_password_email(receiver_email, new_password, sam_account='', display_name='', department=''):
    # 邮件配置（从环境变量读取）
    sender_email = os.getenv("EMAIL_SENDER")
    sender_password = os.getenv("EMAIL_PASSWORD")
    smtp_server = os.getenv("EMAIL_SMTP_SERVER")
    smtp_port = int(os.getenv("EMAIL_SMTP_PORT", "465"))
    bcc_emails = os.getenv("EMAIL_BCC", "").strip()
    additional_content = os.getenv("EMAIL_ADDITIONAL_CONTENT", "").strip()
    
    # 提取登录用的纯邮箱地址（如果 sender_email 包含显示名称）
    if '<' in sender_email and '>' in sender_email:
        login_email = sender_email.split('<')[1].split('>')[0].strip()
    else:
        login_email = sender_email

    # 检查配置
    if not all([sender_email, sender_password, smtp_server]):
        return False, "邮件配置不完整，请检查 .env 文件"

    # 获取当前日期
    current_date = datetime.now().strftime("%Y年%m月%d日")
    
    # 处理空部门
    if not department or department.strip() == '':
        department = '（未分配部门）'

    # HTML 邮件内容
    html_body = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, 'Microsoft YaHei', sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 100vw; margin: 20px auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }}
            .header {{ background-color: #4CAF50; color: white; padding: 5px; border-radius: 5px 5px 0 0; text-align: center; }}
            .content {{ padding: 20px; background-color: #f9f9f9; }}
            .info-box {{ background-color: white; padding: 15px; margin: 15px 0; border-left: 4px solid #4CAF50; }}
            .footer {{ margin-top: 20px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 12px; color: #666; }}
            table {{ width: 100%; border-collapse: collapse; }}
            td {{ padding: 8px; }}
            .label {{ font-weight: bold; width: 100px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>域账号密码通知</h2>
            </div>
            <div class="content">
                <p>您好，<strong>{display_name}</strong>：</p>
                <p>您的域账号已创建，以下是您的登录信息：</p>
                
                <div class="info-box">
                    <table>
                        <tr>
                            <td class="label">账号：</td>
                            <td>{sam_account}</td>
                        </tr>
                        <tr>
                            <td class="label">姓名：</td>
                            <td>{display_name}</td>
                        </tr>
                        <tr>
                            <td class="label">邮箱：</td>
                            <td>{receiver_email}</td>
                        </tr>
                        <tr>
                            <td class="label">部门：</td>
                            <td>{department}</td>
                        </tr>
                        <tr>
                            <td class="label">密码：</td>
                            <td>{html.escape(new_password)}</td>
                        </tr>
                    </table>
                </div>
                
                <p><strong>重要提示：</strong></p>
                <ul>
                    <li>请妥善保管您的密码，不要与他人分享</li>
                </ul>
                
                {additional_content}
                
                <div class="footer">
                    <p>此邮件由系统自动发送，请勿回复。</p>
                    <p>发送时间：{current_date}</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """

    try:
        # 创建邮件对象
        message = MIMEMultipart("alternative")
        message["Subject"] = "域账号密码通知"
        message["From"] = sender_email
        message["To"] = receiver_email
        
        # 添加密送
        if bcc_emails:
            message["Bcc"] = bcc_emails
        
        # 添加HTML内容
        html_part = MIMEText(html_body, "html", "utf-8")
        message.attach(html_part)
        
        # 发送邮件
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(login_email, sender_password)
            server.send_message(message)
        
        return True, "发送成功"
    
    except Exception as e:
        return False, str(e)
