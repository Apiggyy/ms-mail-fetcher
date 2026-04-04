import logging
import time

from fastapi import HTTPException
from sqlalchemy.orm import Session

from app.models.models import Account
from app.services.mail_cache import (
    clear_list_cache,
    get_detail_cache,
    get_list_cache,
    set_detail_cache,
    set_list_cache,
)
from app.services.tokens import get_access_token_for_account
from app.schemas.schemas import MailDetailResponse, MailListResponse
from app.utils.outlook_imap_client import (
    INBOX_FOLDER_NAME,
    JUNK_FOLDER_NAME,
    get_email_detail_by_uid,
    get_emails_by_folder_paginated,
)

logger = logging.getLogger("ms_mail_fetcher")


def resolve_folder(folder: str) -> str:
    folder_lower = folder.lower()
    if folder_lower == "inbox":
        return INBOX_FOLDER_NAME
    if folder_lower == "spam":
        return JUNK_FOLDER_NAME
    raise HTTPException(status_code=400, detail="folder 仅支持 inbox 或 spam")


def get_account_or_404(db: Session, account_id: int) -> Account:
    account = db.query(Account).filter(Account.id == account_id).first()
    if not account:
        raise HTTPException(status_code=404, detail="账号不存在")
    return account


def _log_mail_operation(operation: str, account: Account, folder: str, token_result: dict, mail_result: dict, total_ms: float,
                        *, page: int | None = None, page_size: int | None = None, retry: bool = False) -> None:
    timings = mail_result.get("timings", {})
    message_parts = [
        f"mail.{operation}",
        f"account={account.email}",
        f"folder={folder}",
        f"retry={retry}",
        f"token_source={token_result.get('token_source', 'unknown')}",
        f"token_ms={token_result.get('duration_ms', 0.0):.2f}",
        f"refresh_token_rotated={token_result.get('refresh_token_rotated', False)}",
        f"imap_auth_ms={timings.get('imap_auth_ms', 0.0):.2f}",
        f"select_ms={timings.get('select_ms', 0.0):.2f}",
        f"search_ms={timings.get('search_ms', 0.0):.2f}",
        f"fetch_ms={timings.get('fetch_ms', 0.0):.2f}",
        f"total_ms={total_ms:.2f}",
    ]
    if page is not None:
        message_parts.append(f"page={page}")
    if page_size is not None:
        message_parts.append(f"page_size={page_size}")
    if not mail_result.get("success"):
        message_parts.append(f"error={mail_result.get('error_msg', 'unknown')}")

    logger.info(" | ".join(message_parts))


def _log_mail_cache_hit(operation: str, account: Account, folder: str, total_ms: float,
                        *, page: int | None = None, page_size: int | None = None, message_id: str | None = None) -> None:
    message_parts = [
        f"mail.{operation}",
        f"account={account.email}",
        f"folder={folder}",
        "cache_hit=True",
        "token_source=skipped",
        "token_ms=0.00",
        "imap_auth_ms=0.00",
        "select_ms=0.00",
        "search_ms=0.00",
        "fetch_ms=0.00",
        f"total_ms={total_ms:.2f}",
    ]
    if page is not None:
        message_parts.append(f"page={page}")
    if page_size is not None:
        message_parts.append(f"page_size={page_size}")
    if message_id is not None:
        message_parts.append(f"message_id={message_id}")

    logger.info(" | ".join(message_parts))


def _get_mail_result_with_retry(
    db: Session,
    account: Account,
    fetch_func,
    operation: str,
    folder: str,
    *,
    page: int | None = None,
    page_size: int | None = None,
) -> dict:
    started_at = time.perf_counter()
    token_result = get_access_token_for_account(db, account)
    if not token_result.get("success"):
        total_ms = (time.perf_counter() - started_at) * 1000
        failed_result = {
            "success": False,
            "error_msg": token_result.get("error_msg", "获取 access_token 失败"),
            "timings": {
                "imap_auth_ms": 0.0,
                "select_ms": 0.0,
                "search_ms": 0.0,
                "fetch_ms": 0.0,
            },
        }
        _log_mail_operation(operation, account, folder, token_result, failed_result, total_ms, page=page, page_size=page_size)
        return failed_result

    mail_result = fetch_func(token_result["access_token"])
    retry = False
    if not mail_result.get("success") and mail_result.get("auth_failed"):
        retry = True
        retry_token_result = get_access_token_for_account(db, account, force_refresh=True)
        if retry_token_result.get("success"):
            token_result = retry_token_result
            mail_result = fetch_func(token_result["access_token"])
        else:
            token_result = retry_token_result
            mail_result = {
                "success": False,
                "error_msg": retry_token_result.get("error_msg", "刷新 access_token 失败"),
                "timings": mail_result.get("timings", {}),
            }

    total_ms = (time.perf_counter() - started_at) * 1000
    _log_mail_operation(operation, account, folder, token_result, mail_result, total_ms, page=page, page_size=page_size, retry=retry)
    return mail_result


def list_mails(db: Session, account_id: int, folder: str, page: int, page_size: int, force_refresh: bool = False) -> MailListResponse:
    started_at = time.perf_counter()
    account = get_account_or_404(db, account_id)
    target_folder = resolve_folder(folder)
    page_number = page - 1
    normalized_folder = folder.lower()

    if force_refresh:
        clear_list_cache(account.id, normalized_folder, page, page_size)
    else:
        cached_result = get_list_cache(account.id, normalized_folder, page, page_size)
        if cached_result is not None:
            total_ms = (time.perf_counter() - started_at) * 1000
            _log_mail_cache_hit("list", account, normalized_folder, total_ms, page=page, page_size=page_size)
            return MailListResponse(**cached_result)

    mail_result = _get_mail_result_with_retry(
        db,
        account,
        lambda access_token: get_emails_by_folder_paginated(
            email_address=account.email,
            access_token=access_token,
            target_folder=target_folder,
            page_number=page_number,
            emails_per_page=page_size,
        ),
        "list",
        normalized_folder,
        page=page,
        page_size=page_size,
    )

    if not mail_result.get("success"):
        raise HTTPException(status_code=400, detail=mail_result.get("error_msg", "读取邮件失败"))

    response_payload = {
        "account_id": account.id,
        "email": account.email,
        "folder": normalized_folder,
        "page": page,
        "page_size": page_size,
        "total": mail_result.get("total_emails", 0),
        "items": mail_result.get("emails", []),
    }
    set_list_cache(account.id, normalized_folder, page, page_size, response_payload)
    return MailListResponse(**response_payload)


def get_mail_detail(db: Session, account_id: int, folder: str, message_id: str) -> MailDetailResponse:
    started_at = time.perf_counter()
    account = get_account_or_404(db, account_id)
    target_folder = resolve_folder(folder)
    normalized_folder = folder.lower()

    cached_detail = get_detail_cache(account.id, normalized_folder, message_id)
    if cached_detail is not None:
        total_ms = (time.perf_counter() - started_at) * 1000
        _log_mail_cache_hit("detail", account, normalized_folder, total_ms, message_id=message_id)
        return MailDetailResponse(**cached_detail)

    detail_result = _get_mail_result_with_retry(
        db,
        account,
        lambda access_token: get_email_detail_by_uid(
            email_address=account.email,
            access_token=access_token,
            target_uid=message_id,
            target_folder=target_folder,
        ),
        "detail",
        normalized_folder,
    )

    if not detail_result.get("success"):
        raise HTTPException(status_code=400, detail=detail_result.get("error_msg", "读取邮件详情失败"))

    response_payload = {
        "account_id": account.id,
        "email": account.email,
        "folder": normalized_folder,
        "message_id": message_id,
        "detail": detail_result.get("detail", {}),
    }
    set_detail_cache(account.id, normalized_folder, message_id, response_payload)
    return MailDetailResponse(**response_payload)
