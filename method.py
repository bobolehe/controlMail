import smtplib
import time
import imaplib
import email
import json
import os
import requests
import re
import pytz

from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.header import Header
from bs4 import BeautifulSoup


def load_processed_emails(username):
    """加载已处理的邮件ID"""
    if os.path.exists('processed_emails.json'):
        with open('processed_emails.json', 'r') as f:
            # 修改这里：直接使用json.load而不是f.json()
            try:
                return list(set(json.load(f).get(username, [])))
            except json.JSONDecodeError:
                return []
    return []


def save_processed_emails(username, processed_ids):
    """保存已处理的邮件ID"""
    data = {}
    if os.path.exists('processed_emails.json'):
        with open('processed_emails.json', 'r') as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                pass

    with open('processed_emails.json', 'w') as f:
        data[username] = list(processed_ids)
        json.dump(data, f)


def check_email_inbox(subject_keyword, save_dir='attachments', imap_server='imap.qq.com', username='3066@qq.com', password='ptgmuusmzpuddfdd', folder='INBOX', today=None):
    """
    检查邮箱收件箱是否存在包含指定关键词的邮件，并下载附件
    """
    pattern = r"购方名称：([^<\n]+)(?:<br>|\n|<br>\n).*?金额合计：([\d\.]+)元(?:<br>|\n|<br>\n).*?发票号码：(\d+)(?:<br>|\n|<br>\n)"

    try:
        # 创建附件保存目录
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        if folder == '收件箱':
            folder = 'INBOX'
        elif folder == '已发送':
            folder = 'Sent Messages'
        elif folder == '草稿箱':
            folder = 'Drafts'
        elif folder == '已删除':
            folder = 'Deleted Messages'
        elif folder == '垃圾箱':
            folder = 'Junk'

        imap.select(folder)
        # 设置搜索条件为最近的邮件
        # 获取今天的日期
        if today is None:
            today = datetime.now() - timedelta(days=1)
            time_str = '2024-12-05 09:51:49'
            today = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')

        timezone = pytz.timezone("Asia/Shanghai")
        today = timezone.localize(today)
        print(f"搜索时间: {today}")
        _, messages = imap.search(None, "ALL")
        message_ids = messages[0].split()
        processed_ids = load_processed_emails(username)
        new_messages = []
        file_list = []
        error_file_list = []

        for msg_id in message_ids:
            if msg_id.decode() not in processed_ids:
                _, msg_data = imap.fetch(msg_id, '(RFC822)')
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                # 改进的主题解码逻辑
                subject = ''
                subject_header = email.header.decode_header(email_message['Subject'])[0]

                if isinstance(subject_header[0], bytes):
                    # 尝试多种编码方式
                    encodings = ['utf-8', 'gb18030', 'gb2312', 'gbk', 'iso-8859-1']
                    for encoding in encodings:
                        try:
                            subject = subject_header[0].decode(encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                else:
                    subject = subject_header[0]

                if subject_keyword in subject:
                    print(f"\n找到匹配邮件{msg_id}:")
                    print(f"主题: {subject}")
                    print(f"发件人: {email_message['From']}")
                    email_message_date_str = email_message['Date'].split(' (')[0]
                    email_message_date = datetime.strptime(email_message_date_str, "%a, %d %b %Y %H:%M:%S %z")
                    print(f"日期: {email_message_date}")
                    # 输出邮件内容
                    if email_message_date < today:
                        processed_ids.append(msg_id.decode())
                        continue
                    # 获取邮件正文内容
                    body = ""
                    html_content = ""

                    if email_message.is_multipart():
                        for part in email_message.walk():
                            content_type = part.get_content_type()
                            try:
                                # 获取编码信息
                                charset = part.get_content_charset() or 'utf-8'
                                payload = part.get_payload(decode=True)

                                if payload:
                                    decoded_content = payload.decode(charset)
                                    if content_type == "text/plain":
                                        body = decoded_content
                                    elif content_type == "text/html":
                                        html_content = decoded_content
                                    # print(f"使用编码 {charset} 成功解码 {content_type}")
                            except Exception as e:
                                print(f"解码失败 {content_type}: {str(e)}")
                                continue
                    else:
                        try:
                            charset = email_message.get_content_charset() or 'utf-8'
                            payload = email_message.get_payload(decode=True)
                            if payload:
                                body = payload.decode(charset)
                        except Exception as e:
                            print(f"解码失败: {str(e)}")

                    # 优先使用HTML内容，因为它包含链接
                    final_content = html_content if html_content else body
                    if not final_content:
                        final_content = "未找到邮件正文内容"
                    # 使用re.search()来查找匹配的内容
                    cleaned_text = re.sub(r"<br\s*/?>\n", "\n", final_content)  # 替换 HTML 换行符为 \n
                    cleaned_text = re.sub(r"<br\s*/?>", "", cleaned_text)  # 替换 HTML 换行符为 \n
                    pattern = r"购方名称：(.+?)金额合计：([\d\.]+)元.*?发票号码：(\d+)"

                    match = re.search(pattern, cleaned_text, re.DOTALL)
                    if match:
                        purchaser_name, total_amount, invoice_number = match.groups()
                        purchaser_name = purchaser_name.strip()
                        total_amount = total_amount.strip()
                        invoice_number = invoice_number.strip()
                        # 创建购方名称文件夹
                        if not os.path.exists(os.path.join(save_dir, purchaser_name)):
                            os.makedirs(os.path.join(save_dir, purchaser_name))
                        filename = f"{purchaser_name}_{total_amount}元.pdf"
                    else:
                        print("未匹配到相关信息")

                    if html_content and match:
                        try:
                            soup = BeautifulSoup(html_content, 'html.parser')
                            links = soup.find_all('a')
                            for link in links:
                                download_url = link.get('href')
                                if download_url and ('download' in download_url.lower() or 'pdf' in download_url.lower()) and 'pdf' in download_url.lower():
                                    try:
                                        headers = {
                                            "Host": "msauth.norincogroup-ebuy.com",
                                            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:52.0) Gecko/20100101 Firefox/52.0",
                                            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                                            "Accept-Language": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
                                            "Accept-Encoding": "gzip, deflate, br",
                                            "DNT": "1",
                                            "Connection": "keep-alive",
                                            "Upgrade-Insecure-Requests": "1",
                                        }
                                        try:
                                            response = requests.get(download_url, headers=headers)  # verify=False 处理https证书问题
                                        except Exception as e:
                                            response = requests.get(download_url, headers=headers, proxies={'http': 'http://127.0.0.1:7890', 'https': 'http://127.0.0.1:7890'})

                                        if response.status_code == 200:
                                            # 生成文件名
                                            filepath = os.path.join(os.path.join(save_dir, purchaser_name), filename)

                                            with open(filepath, 'wb') as f:
                                                f.write(response.content)
                                            print(f"已下载发票文件: {filepath}")
                                            file_list.append(filepath)
                                            imap.store(msg_id, '+FLAGS', '\\Seen')
                                        else:
                                            error_file_list.append(download_url)
                                            print(f"下载失败，状态码: {response.status_code}")
                                    except Exception as e:
                                        error_file_list.append(download_url)
                                        print(f"下载文件时出错: {str(e)}")
                        except Exception as e:
                            print(f"解析HTML内容时出错: {str(e)}")

                    # 处理附件
                    for part in email_message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        # 改进的文件名解码逻辑
                        filename = part.get_filename()
                        if filename:
                            # 解码文件名
                            filename_header = email.header.decode_header(filename)[0]
                            if isinstance(filename_header[0], bytes):
                                # 尝试多种编码方式
                                for encoding in encodings:
                                    try:
                                        filename = filename_header[0].decode(encoding or 'utf-8')
                                        break
                                    except UnicodeDecodeError:
                                        continue
                            else:
                                filename = filename_header[0]

                            # 清理文件名中的非法字符
                            filename = "".join(c for c in filename if c.isprintable() and c not in r'<>:"/\|?*')

                            filepath = os.path.join(save_dir, filename)
                            if os.path.exists(filepath):
                                filename = f"{int(time.time())}_{filename}"
                                filepath = os.path.join(save_dir, filename)

                            try:
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"已下载附件: {filename}")
                            except Exception as e:
                                print(f"保存附件失败: {filename}, 错误: {str(e)}")

                    new_messages.append(msg_id)
                    processed_ids.append(msg_id.decode())
        save_processed_emails(username, processed_ids)
        imap.close()
        imap.logout()

        return len(new_messages), file_list
    except Exception as e:
        return 0, [f"检查邮箱失败: {str(e)}"]


def monitor_inbox(subject_keyword, interval=300):
    """
    定期监控邮箱
    """
    print(f"开始监控邮箱，查找主题包含 '{subject_keyword}' 的邮件...")
    while True:
        count = check_email_inbox(
            subject_keyword,
            imap_server=EMAIL_SERVER,
            username=EMAIL_USERNAME,
            password=EMAIL_PASSWORD,
            folder=FOLDER
        )
        print(f"找到 {count} 封相关邮件")
        time.sleep(interval)


# 获取邮件可以查看文件夹
def get_email_folders(imap_server, username, password):
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(username, password)
    status, folders = imap.list()
    if status == 'OK':
        folder_list = []
        for folder in [folder.decode().split(' "/" ')[-1].strip('"') for folder in folders]:
            if folder == 'INBOX':
                folder_list.append('收件箱')
            elif folder == 'Sent Messages':
                folder_list.append('已发送')
            elif folder == 'Drafts':
                folder_list.append('草稿箱')
            elif folder == 'Deleted Messages':
                folder_list.append('已删除')
            elif folder == 'Junk':
                folder_list.append('垃圾箱')
            else:
                folder_list.append(folder)
        return folder_list
    else:
        print(f"获取邮箱文件夹失败: {status}")
        return []


if __name__ == '__main__':
    pass
    # 配置邮箱参数
    # EMAIL_SERVER = 'imap.qq.com'
    # EMAIL_USERNAME = '6958@qq.com'
    # EMAIL_USERNAME = '306608@qq.com'
    # EMAIL_PASSWORD = 'orgtfmjbqqfnbfia'
    # EMAIL_PASSWORD = 'ptgmuusmzpuddfdd'
    # SUBJECT_KEYWORD = '发票'
    # CHECK_INTERVAL = 300  # 5分钟检查一次
    # ATTACHMENT_DIR = 'attachments'  # 附件保存目录
    # FOLDER = '收件箱'  # 收件箱
    #
    # # 启动监控
    # monitor_inbox(
    #     subject_keyword=SUBJECT_KEYWORD,
    #     interval=CHECK_INTERVAL
    # )

    # folders = get_email_folders(EMAIL_SERVER, EMAIL_USERNAME, EMAIL_PASSWORD)
    # print(folders)

