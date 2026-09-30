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
from email.utils import parsedate_to_datetime
from bs4 import BeautifulSoup


def sanitize_path_name(name):
    """清理 Windows 文件/目录名中的非法字符。"""
    name = re.sub(r'[<>:"/\\|?*]+', '_', str(name or '').strip())
    name = re.sub(r'\s+', ' ', name)
    return name.strip(' .') or '未命名'


def get_unique_filepath(directory, filename):
    """如果目标文件已存在，自动追加序号，避免覆盖已有发票。"""
    base, ext = os.path.splitext(filename)
    filepath = os.path.join(directory, filename)
    index = 1
    while os.path.exists(filepath):
        filepath = os.path.join(directory, f"{base}_{index}{ext}")
        index += 1
    return filepath


def normalize_folder_name(folder):
    """把界面显示的常用中文文件夹名转换为 IMAP 通用/服务端文件夹名。"""
    folder_map = {
        '收件箱': 'INBOX',
        'INBOX': 'INBOX',
        # QQ 邮箱历史映射，保持原行为不变。
        '已发送': 'Sent Messages',
        '草稿箱': 'Drafts',
        '已删除': 'Deleted Messages',
        '垃圾箱': 'Junk',
    }
    return folder_map.get(folder, folder)


def parse_imap_folder_name(folder_line):
    """从 IMAP LIST 响应中提取文件夹名，兼容 QQ/163/Gmail 常见格式。"""
    if isinstance(folder_line, bytes):
        text = folder_line.decode('utf-8', errors='replace')
    else:
        text = str(folder_line)

    quoted_names = re.findall(r'"((?:[^"\\]|\\.)*)"', text)
    if quoted_names:
        return quoted_names[-1].replace(r'\"', '"')

    parts = text.split()
    return parts[-1] if parts else ''


def decode_mime_header(value):
    """解码 Subject / 附件文件名这类 MIME Header。"""
    if not value:
        return ''
    encodings = ['utf-8', 'gb18030', 'gb2312', 'gbk', 'iso-8859-1']
    decoded_parts = []
    for part, encoding in email.header.decode_header(value):
        if isinstance(part, bytes):
            candidates = [encoding] if encoding else []
            candidates.extend(encodings)
            decoded = None
            for enc in candidates:
                if not enc:
                    continue
                try:
                    decoded = part.decode(enc)
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            decoded_parts.append(decoded if decoded is not None else part.decode('utf-8', errors='replace'))
        else:
            decoded_parts.append(part)
    return ''.join(decoded_parts)


def decode_payload(payload, charset):
    """按邮件声明编码解码正文，失败时尝试常见中文编码。"""
    if not payload:
        return ''
    candidates = [charset] if charset else []
    candidates.extend(['utf-8', 'gb18030', 'gb2312', 'gbk', 'iso-8859-1'])
    for enc in candidates:
        if not enc:
            continue
        try:
            return payload.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return payload.decode('utf-8', errors='replace')


def extract_email_content(email_message):
    """提取 text/plain 和 text/html 正文。"""
    body = ''
    html_content = ''
    if email_message.is_multipart():
        for part in email_message.walk():
            content_type = part.get_content_type()
            if content_type not in ('text/plain', 'text/html'):
                continue
            payload = part.get_payload(decode=True)
            decoded_content = decode_payload(payload, part.get_content_charset())
            if content_type == 'text/plain' and decoded_content:
                body = decoded_content
            elif content_type == 'text/html' and decoded_content:
                html_content = decoded_content
    else:
        payload = email_message.get_payload(decode=True)
        body = decode_payload(payload, email_message.get_content_charset())
    return body, html_content


def content_to_text(content):
    """把 HTML/纯文本正文整理成便于正则匹配的文本。"""
    if not content:
        return ''
    text = re.sub(r'<br\s*/?>', '\n', content, flags=re.I)
    text = BeautifulSoup(text, 'html.parser').get_text('\n')
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def parse_invoice_info(content):
    """从邮件正文中提取购方名称、金额、发票号码；兼容少量字段别名。"""
    text = content_to_text(content)
    patterns = [
        r"(?:购方名称|购买方名称)[:：]\s*(.+?)\s*(?:金额合计|价税合计)[:：]\s*￥?\s*([\d.]+)\s*元?.*?发票号码[:：]\s*(\d+)",
        r"(?:购方名称|购买方名称)[:：]\s*(.+?)\s*发票号码[:：]\s*(\d+).*?(?:金额合计|价税合计)[:：]\s*￥?\s*([\d.]+)\s*元?",
    ]
    for index, pattern in enumerate(patterns):
        match = re.search(pattern, text, re.DOTALL)
        if not match:
            continue
        groups = match.groups()
        if index == 0:
            purchaser_name, total_amount, invoice_number = groups
        else:
            purchaser_name, invoice_number, total_amount = groups
        return {
            'purchaser_name': sanitize_path_name(purchaser_name),
            'total_amount': sanitize_path_name(total_amount),
            'invoice_number': sanitize_path_name(invoice_number),
        }
    return None


def extract_links(html_content, text_content=''):
    """从 HTML 和纯文本中提取候选下载链接。"""
    links = []
    if html_content:
        soup = BeautifulSoup(html_content, 'html.parser')
        for link in soup.find_all('a'):
            href = link.get('href')
            if href:
                links.append(href.strip())
    for url in re.findall(r'https?://[^\s<>"\']+', text_content or ''):
        links.append(url.rstrip('.,;，；)）'))

    deduped = []
    seen = set()
    for link in links:
        if link and link not in seen:
            seen.add(link)
            deduped.append(link)
    return deduped


def is_probable_pdf_response(response, url):
    """判断响应是否适合作为 PDF 保存。"""
    url_lower = (url or '').lower()
    content_type = response.headers.get('Content-Type', '').lower()
    content_disposition = response.headers.get('Content-Disposition', '').lower()
    content_start = response.content[:16]
    return (
        response.status_code == 200 and (
            b'%PDF' in content_start or
            'pdf' in content_type or
            'pdf' in content_disposition or
            'pdf' in url_lower or
            'download' in url_lower
        )
    )


def download_pdf_url(download_url, filepath):
    """下载 PDF；目标存在时直接跳过并视为成功。"""
    if os.path.exists(filepath):
        print(f"发票文件已存在，跳过重复下载: {filepath}")
        return True, filepath

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64; rv:52.0) Gecko/20100101 Firefox/52.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
        "Accept-Encoding": "gzip, deflate, br",
        "DNT": "1",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }
    try:
        response = requests.get(download_url, headers=headers, timeout=20)
    except Exception:
        response = requests.get(
            download_url,
            headers=headers,
            proxies={'http': 'http://127.0.0.1:7890', 'https': 'http://127.0.0.1:7890'},
            timeout=20
        )

    if not is_probable_pdf_response(response, download_url):
        return False, f"非PDF响应或下载失败: {download_url} 状态码={response.status_code} Content-Type={response.headers.get('Content-Type', '')}"

    with open(filepath, 'wb') as f:
        f.write(response.content)
    print(f"已下载发票文件: {filepath}")
    return True, filepath


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
    检查邮箱收件箱是否存在包含指定关键词的邮件，并下载发票 PDF 或附件。
    使用 IMAP UID 作为已处理缓存键，避免普通消息序号变化导致漏处理/误跳过。
    """
    imap = None
    try:
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        folder = normalize_folder_name(folder)

        status, _ = imap.select(folder)
        if status != 'OK':
            status, _ = imap.select(f'"{folder}"')
        if status != 'OK':
            return 0, [f"选择邮箱文件夹失败: {folder}"]

        if today is None:
            today = datetime.now()
        timezone = pytz.timezone("Asia/Shanghai")
        if today.tzinfo is None:
            today = timezone.localize(today)
        else:
            today = today.astimezone(timezone)
        print(f"搜索时间: {today}")

        since_date = today.strftime('%d-%b-%Y')
        status, messages = imap.uid('search', None, 'SINCE', since_date)
        if status != 'OK':
            return 0, [f"搜索邮件失败: {status}"]

        message_uids = list(reversed(messages[0].split()))
        processed_ids = set(load_processed_emails(username))
        new_messages = []
        file_list = []
        error_file_list = []

        for uid in message_uids:
            uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
            if uid_str in processed_ids:
                continue

            status, msg_data = imap.uid('fetch', uid, '(RFC822)')
            if status != 'OK' or not msg_data:
                error_file_list.append(f"UID {uid_str} 拉取失败: {status}")
                continue

            email_body = None
            for item in msg_data:
                if isinstance(item, tuple) and item[1]:
                    email_body = item[1]
                    break
            if not email_body:
                error_file_list.append(f"UID {uid_str} 邮件内容为空")
                continue

            email_message = email.message_from_bytes(email_body)
            subject = decode_mime_header(email_message.get('Subject', ''))
            if subject_keyword and subject_keyword not in subject:
                continue

            print(f"\n找到匹配邮件 UID {uid_str}:")
            print(f"主题: {subject}")
            print(f"发件人: {email_message.get('From', '')}")

            try:
                email_message_date = parsedate_to_datetime(email_message.get('Date'))
                if email_message_date.tzinfo is None:
                    email_message_date = timezone.localize(email_message_date)
                else:
                    email_message_date = email_message_date.astimezone(timezone)
            except Exception as e:
                error_file_list.append(f"{subject} 解析邮件日期失败: {email_message.get('Date')} {str(e)}")
                continue
            print(f"日期: {email_message_date}")
            if email_message_date < today:
                continue

            message_downloaded = False
            body, html_content = extract_email_content(email_message)
            final_content = html_content if html_content else body
            invoice_info = parse_invoice_info(final_content)

            if invoice_info:
                purchaser_name = invoice_info['purchaser_name']
                total_amount = invoice_info['total_amount']
                invoice_number = invoice_info['invoice_number']
                purchaser_dir = os.path.join(save_dir, purchaser_name)
                filename = f"{purchaser_name}_{total_amount}元_{invoice_number}.pdf"
            else:
                print("未匹配到发票字段，进入 PDF 链接/附件兜底下载")
                purchaser_dir = os.path.join(save_dir, '未匹配发票信息')
                filename = f"{sanitize_path_name(subject) or '邮件'}_UID{uid_str}.pdf"

            if not os.path.exists(purchaser_dir):
                os.makedirs(purchaser_dir)

            candidate_links = extract_links(html_content, body)
            if not candidate_links:
                error_file_list.append(f"{subject} 未找到下载链接")

            for download_url in candidate_links:
                filepath = os.path.join(purchaser_dir, filename)
                ok, result = download_pdf_url(download_url, filepath)
                if ok:
                    file_list.append(result)
                    message_downloaded = True
                    break
                error_file_list.append(str(result))

            # 处理附件：如果正文已识别出发票信息，PDF 附件也使用发票唯一文件名。
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                attachment_name = decode_mime_header(part.get_filename())
                if not attachment_name:
                    continue
                attachment_name = sanitize_path_name(attachment_name)
                payload = part.get_payload(decode=True)
                if not payload:
                    continue

                if invoice_info and attachment_name.lower().endswith('.pdf'):
                    filepath = os.path.join(purchaser_dir, filename)
                    if os.path.exists(filepath):
                        print(f"发票附件已存在，跳过重复保存: {filepath}")
                    else:
                        with open(filepath, 'wb') as f:
                            f.write(payload)
                        print(f"已保存发票附件: {filepath}")
                else:
                    filepath = get_unique_filepath(save_dir, attachment_name)
                    with open(filepath, 'wb') as f:
                        f.write(payload)
                    print(f"已下载附件: {filepath}")

                file_list.append(filepath)
                message_downloaded = True

            if message_downloaded:
                new_messages.append(uid_str)
                processed_ids.add(uid_str)
                imap.uid('store', uid, '+FLAGS', '\\Seen')
            else:
                error_file_list.append(f"{subject} 未成功下载 PDF 或附件")

        save_processed_emails(username, processed_ids)

        if not new_messages and error_file_list:
            return 0, [f"未成功下载文件，错误/链接: {'; '.join(error_file_list[:3])}"]
        return len(new_messages), file_list
    except Exception as e:
        return 0, [f"检查邮箱失败: {str(e)}"]
    finally:
        if imap is not None:
            try:
                imap.close()
            except Exception:
                pass
            try:
                imap.logout()
            except Exception:
                pass

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
    try:
        imap.login(username, password)
        status, folders = imap.list()
        if status == 'OK':
            folder_list = []
            for folder in [parse_imap_folder_name(folder) for folder in folders]:
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
            # 收件箱是三类邮箱最常用监控入口，放到第一项；不改变 QQ 的 INBOX 映射。
            if '收件箱' in folder_list:
                folder_list.insert(0, folder_list.pop(folder_list.index('收件箱')))
            return folder_list
        else:
            print(f"获取邮箱文件夹失败: {status}")
            return []
    finally:
        try:
            imap.logout()
        except Exception:
            pass


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

