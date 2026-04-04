import imaplib
import email
import re
import time
from email.header import decode_header
from email import utils as email_utils
import requests

# --- 常量配置 ---
IMAP_SERVER = 'outlook.live.com'
IMAP_PORT = 993
TOKEN_URL = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'
INBOX_FOLDER_NAME = "INBOX"
JUNK_FOLDER_NAME = "Junk"
UID_PATTERN = re.compile(rb"UID (?P<uid>\d+)")


def _looks_like_html(content: str) -> bool:
    if not content:
        return False
    text = content.lstrip().lower()
    return (
        text.startswith("<!doctype html")
        or text.startswith("<html")
        or ("<body" in text and "</body>" in text)
    )


def decode_header_value(header_value):
    """辅助函数：解码邮件头中的中文字符等"""
    if header_value is None: return ""
    decoded_string = ""
    try:
        parts = decode_header(str(header_value))
        for part, charset in parts:
            if isinstance(part, bytes):
                try:
                    decoded_string += part.decode(charset if charset else 'utf-8', 'replace')
                except LookupError:
                    decoded_string += part.decode('utf-8', 'replace')
            else:
                decoded_string += str(part)
    except Exception:
        return str(header_value)
    return decoded_string


# =======================================================
# 专属工具：手动刷新并获取新的 Refresh Token (你自己定期调)
# =======================================================
def refresh_oauth_token_manually(client_id, current_refresh_token):
    """
    专门用来刷新 Token 的工具。
    返回包含新 refresh_token 的字典，拿到后你自己保存到本地或数据库。
    """
    result = {
        "success": False,
        "new_refresh_token": "",
        "new_access_token": "",
        "error_msg": ""
    }
    try:
        response = requests.post(TOKEN_URL, data={
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': current_refresh_token,
            'scope': 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access'
        }, timeout=15)
        response.raise_for_status()
        token_data = response.json()

        result["new_access_token"] = token_data.get('access_token', "")
        result["new_refresh_token"] = token_data.get('refresh_token', "")
        result["expires_in"] = token_data.get('expires_in', 0)

        if result["new_access_token"]:
            result["success"] = True
        else:
            result["error_msg"] = "微软接口未返回 access_token"

    except Exception as e:
        result["error_msg"] = f"刷新 Token 失败: {e}"

    return result


def _new_timing_dict():
    return {
        "imap_auth_ms": 0.0,
        "select_ms": 0.0,
        "search_ms": 0.0,
        "fetch_ms": 0.0,
        "total_ms": 0.0,
    }


def _finalize_timing(result: dict, started_at: float) -> dict:
    result["timings"]["total_ms"] = (time.perf_counter() - started_at) * 1000
    return result


def _extract_header_bytes(msg_data):
    if isinstance(msg_data[0], tuple) and len(msg_data[0]) == 2:
        return msg_data[0][1]
    if isinstance(msg_data, list) and len(msg_data) > 1:
        return msg_data[1]
    return None


def _extract_uid_from_fetch_meta(fetch_meta) -> str | None:
    if isinstance(fetch_meta, bytes):
        match = UID_PATTERN.search(fetch_meta)
        if match:
            return match.group("uid").decode("utf-8", "replace")
    return None


def _parse_mail_header(uid_str: str, header_content_bytes, target_folder: str) -> dict:
    subject_str = "(No Subject)"
    formatted_date_str = "(No Date)"
    from_name = "(Unknown)"
    from_email = ""

    if header_content_bytes:
        header_message = email.message_from_bytes(header_content_bytes)
        subject_str = decode_header_value(header_message.get('Subject', '(No Subject)'))
        from_str = decode_header_value(header_message.get('From', '(Unknown Sender)'))

        if '<' in from_str and '>' in from_str:
            from_name = from_str.split('<')[0].strip().strip('"')
            from_email = from_str.split('<')[1].split('>')[0].strip()
        else:
            from_email = from_str.strip()
            if '@' in from_email:
                from_name = from_email.split('@')[0]

        date_header_str = header_message.get('Date')
        if date_header_str:
            try:
                dt_obj = email_utils.parsedate_to_datetime(date_header_str)
                if dt_obj:
                    formatted_date_str = dt_obj.strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                pass

    return {
        'uid': uid_str,
        'subject': subject_str,
        'from_name': from_name,
        'from_email': from_email,
        'date': formatted_date_str,
        'folder': target_folder,
    }


def _fetch_mail_headers_batch(imap_conn, page_uids: list[bytes], target_folder: str) -> list[dict]:
    if not page_uids:
        return []

    uid_sequence = b",".join(page_uids)
    typ, msg_data = imap_conn.uid(
        'fetch',
        uid_sequence,
        '(UID BODY.PEEK[HEADER.FIELDS (SUBJECT DATE FROM)])',
    )
    if typ != 'OK' or not msg_data:
        return []

    by_uid: dict[str, dict] = {}
    for item in msg_data:
        if not isinstance(item, tuple) or len(item) != 2:
            continue

        fetch_meta, header_content_bytes = item
        uid_str = _extract_uid_from_fetch_meta(fetch_meta)
        if not uid_str:
            continue

        by_uid[uid_str] = _parse_mail_header(uid_str, header_content_bytes, target_folder)

    ordered_headers: list[dict] = []
    for uid_bytes in page_uids:
        uid_str = uid_bytes.decode('utf-8', 'replace')
        ordered_headers.append(
            by_uid.get(
                uid_str,
                _parse_mail_header(uid_str, None, target_folder),
            )
        )

    return ordered_headers


def get_emails_by_folder_paginated(email_address, access_token, target_folder=INBOX_FOLDER_NAME,
                                   page_number=0, emails_per_page=10):
    """
    分页获取 Outlook 指定文件夹的邮件列表。
    只返回干净的邮件数据，不包含任何 token 刷新的杂质。
    """
    result = {
        "success": False,
        "error_msg": "",
        "total_emails": 0,
        "emails": [],
        "auth_failed": False,
        "timings": _new_timing_dict(),
    }

    imap_conn = None
    started_at = time.perf_counter()
    try:
        auth_started = time.perf_counter()
        imap_conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        auth_string = f"user={email_address}\1auth=Bearer {access_token}\1\1"
        typ, _ = imap_conn.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))
        result["timings"]["imap_auth_ms"] = (time.perf_counter() - auth_started) * 1000

        if typ != 'OK':
            result["error_msg"] = "IMAP 认证失败，请确认凭证有效"
            result["auth_failed"] = True
            return _finalize_timing(result, started_at)

        select_started = time.perf_counter()
        typ, _ = imap_conn.select(target_folder, readonly=True)
        result["timings"]["select_ms"] = (time.perf_counter() - select_started) * 1000
        if typ != 'OK':
            result["error_msg"] = f"选择文件夹 '{target_folder}' 失败"
            return _finalize_timing(result, started_at)

        search_started = time.perf_counter()
        typ, uid_data = imap_conn.uid('search', None, "ALL")
        result["timings"]["search_ms"] = (time.perf_counter() - search_started) * 1000
        if typ != 'OK' or not uid_data[0]:
            result["success"] = True
            return _finalize_timing(result, started_at)

        uids = uid_data[0].split()
        result["total_emails"] = len(uids)
        uids.reverse()

        start_index = page_number * emails_per_page
        end_index = start_index + emails_per_page
        page_uids = uids[start_index:end_index]

        emails_list = []
        fetch_started = time.perf_counter()
        emails_list = _fetch_mail_headers_batch(imap_conn, page_uids, target_folder)
        result["timings"]["fetch_ms"] = (time.perf_counter() - fetch_started) * 1000
        result["emails"] = emails_list
        result["success"] = True
        return _finalize_timing(result, started_at)

    except Exception as e:
        result["error_msg"] = f"发生异常: {e}"
        return _finalize_timing(result, started_at)
    finally:
        if imap_conn:
            try:
                imap_conn.close()
                imap_conn.logout()
            except:
                pass


def get_email_detail_by_uid(email_address, access_token, target_uid, target_folder=INBOX_FOLDER_NAME):
    """
    根据 UID 获取特定邮件的完整内容。
    只返回纯净的详情字典，不含 token 更新逻辑。
    """
    result = {
        "success": False,
        "error_msg": "",
        "auth_failed": False,
        "timings": _new_timing_dict(),
        "detail": {
            "subject": "",
            "from": "",
            "to": "",
            "date": "",
            "body_text": "",
            "body_html": ""
        }
    }

    imap_conn = None
    started_at = time.perf_counter()
    try:
        auth_started = time.perf_counter()
        imap_conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        auth_string = f"user={email_address}\1auth=Bearer {access_token}\1\1"
        typ, _ = imap_conn.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))
        result["timings"]["imap_auth_ms"] = (time.perf_counter() - auth_started) * 1000

        if typ != 'OK':
            result["error_msg"] = "IMAP 认证失败"
            result["auth_failed"] = True
            return _finalize_timing(result, started_at)

        select_started = time.perf_counter()
        typ, _ = imap_conn.select(target_folder, readonly=True)
        result["timings"]["select_ms"] = (time.perf_counter() - select_started) * 1000
        if typ != 'OK':
            result["error_msg"] = f"选择文件夹 '{target_folder}' 失败"
            return _finalize_timing(result, started_at)

        uid_bytes = target_uid.encode('utf-8') if isinstance(target_uid, str) else target_uid
        fetch_started = time.perf_counter()
        typ, msg_data = imap_conn.uid('fetch', uid_bytes, '(RFC822)')
        result["timings"]["fetch_ms"] = (time.perf_counter() - fetch_started) * 1000

        if typ != 'OK' or not msg_data or msg_data[0] is None:
            result["error_msg"] = f"未在 {target_folder} 找到 UID 为 {target_uid} 的邮件"
            return _finalize_timing(result, started_at)

        raw_email_bytes = None
        if isinstance(msg_data[0], tuple) and len(msg_data[0]) == 2:
            raw_email_bytes = msg_data[0][1]
        elif isinstance(msg_data, list):
            for item in msg_data:
                if isinstance(item, tuple) and len(item) == 2:
                    raw_email_bytes = item[1];
                    break

        if not raw_email_bytes:
            result["error_msg"] = "解析邮件数据结构失败"
            return _finalize_timing(result, started_at)

        email_message = email.message_from_bytes(raw_email_bytes)

        result["detail"]["subject"] = decode_header_value(email_message.get('Subject', '(No Subject)'))
        result["detail"]["from"] = decode_header_value(email_message.get('From', '(Unknown Sender)'))
        result["detail"]["to"] = decode_header_value(email_message.get('To', '(Unknown Recipient)'))
        result["detail"]["date"] = email_message.get('Date', '(Unknown Date)')

        body_text = ""
        body_html = ""

        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" not in content_disposition:
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        payload = part.get_payload(decode=True)
                        if payload:
                            decoded_str = payload.decode(charset, errors='replace')
                            if content_type == "text/plain":
                                body_text += decoded_str
                            elif content_type == "text/html":
                                body_html += decoded_str
                    except Exception:
                        pass
        else:
            try:
                charset = email_message.get_content_charset() or 'utf-8'
                payload = email_message.get_payload(decode=True)
                if payload:
                    decoded_str = payload.decode(charset, errors='replace')
                    content_type = email_message.get_content_type()
                    if content_type == "text/html":
                        body_html = decoded_str
                    else:
                        body_text = decoded_str
            except Exception:
                pass

        body_text = body_text.strip()
        body_html = body_html.strip()

        # 兜底：部分邮件服务会把 HTML 正文错误标为 text/plain。
        # 遇到这种情况，把看起来像 HTML 的正文转到 body_html，供前端渲染。
        if not body_html and _looks_like_html(body_text):
            body_html = body_text
            body_text = ""

        result["detail"]["body_text"] = body_text
        result["detail"]["body_html"] = body_html
        result["success"] = True
        return _finalize_timing(result, started_at)

    except Exception as e:
        result["error_msg"] = f"解析异常: {e}"
        return _finalize_timing(result, started_at)
    finally:
        if imap_conn:
            try:
                imap_conn.close()
                imap_conn.logout()
            except:
                pass


# --- 用法示例 ---
if __name__ == "__main__":
    TEST_CLIENT_ID = '9e5f94bc-e8a4-4e73-b8be-63364c29d753'
    TEST_EMAIL = 'AdrianJones7591@outlook.com'
    TEST_REFRESH_TOKEN = 'M.C536_SN1.0.U.-CsWcHph3Kdy2aP9mEIHeES4HDzQj7Fi1WYFq!PZQ6dR!nKznaRGG2!V6SuZFfyIddv9U7ohjp9X4iUu2G978J84tQXM4KFduPV!lGvVClMUehH44yVN*hrIEZl5PqEnKaMWvkvVXZXk8dgCEeSJXLgfgnMGRBmtLg9OvEewOTKFE6l8mI38SSvaIbLD!Z7fnMVzecZJHVFeO1qUekXwUEB7iHdRNPtmI3*CoRQ46OtYhWNDD9j*4*w7Gnpjwvao*55q!ekBGAK7CjdaouLWvXGBvl3MAEhy8gX687P1KSqNRAtnMbPVY0cHSeUFMDhlCGSZ!U!MqG6WuKjpVo4RlUvBCzA29MfS!eFEPUWcrYlPbGtVTDGefRZUrV9lGMhesX8AJeSmb4hFPTN7RvpjcjfmwgOcuPycbJwfFWDLOQrZBaZ7g26h8ruVOmHORDlbDWA$$'

    # 场景 1：日常读取邮件，不再烦恼 token 更新的返回值，干净清爽！
    print(">>> 测试：静默获取邮件列表")
    token_res = refresh_oauth_token_manually(TEST_CLIENT_ID, TEST_REFRESH_TOKEN)
    test_access_token = token_res.get("new_access_token", "")
    list_res = get_emails_by_folder_paginated(
        TEST_EMAIL, test_access_token,
        target_folder=INBOX_FOLDER_NAME, page_number=0, emails_per_page=3
    )
    if list_res["success"]:
        print(f"成功获取 {len(list_res['emails'])} 封邮件。返回字典里再也没有 token 相关的杂乱字段了！")
    else:
        print(f"获取失败: {list_res['error_msg']}")

    # 场景 2：假设过了一个月，你想手动刷新并持久化 token 了
    print("\n>>> 测试：手动调用专职刷新工具")
    if token_res["success"]:
        print(f"刷新成功！新的 Refresh Token: {token_res['new_refresh_token'][:20]}...")
        # TODO: 这里写你保存到本地 txt 或数据库的代码
    else:
        print(f"刷新失败: {token_res['error_msg']}")
