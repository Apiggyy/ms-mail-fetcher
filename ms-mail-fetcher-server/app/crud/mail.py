from datetime import datetime

from fastapi import HTTPException
from sqlalchemy.orm import Session

from app.models.models import Account
from app.schemas.schemas import MailDetailResponse, MailListResponse
from app.utils.outlook_imap_client import (
    INBOX_FOLDER_NAME,
    JUNK_FOLDER_NAME,
    get_email_detail_by_uid,
    get_emails_by_folder_paginated,
    refresh_oauth_token_manually,
)


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


def try_refresh_account_token(db: Session, account: Account) -> bool:
    refresh_result = refresh_oauth_token_manually(account.client_id, account.refresh_token)
    if not refresh_result.get("success"):
        return False

    new_refresh_token = refresh_result.get("new_refresh_token")
    if not new_refresh_token:
        return False

    account.refresh_token = new_refresh_token
    account.last_refresh_time = datetime.utcnow()
    db.add(account)
    db.commit()
    db.refresh(account)
    return True


def list_mails(db: Session, account_id: int, folder: str, page: int, page_size: int) -> MailListResponse:
    account = get_account_or_404(db, account_id)
    target_folder = resolve_folder(folder)
    page_number = page - 1

    mail_result = get_emails_by_folder_paginated(
        email_address=account.email,
        refresh_token=account.refresh_token,
        client_id=account.client_id,
        target_folder=target_folder,
        page_number=page_number,
        emails_per_page=page_size,
    )

    if not mail_result.get("success"):
        refreshed = try_refresh_account_token(db, account)
        if refreshed:
            mail_result = get_emails_by_folder_paginated(
                email_address=account.email,
                refresh_token=account.refresh_token,
                client_id=account.client_id,
                target_folder=target_folder,
                page_number=page_number,
                emails_per_page=page_size,
            )

    if not mail_result.get("success"):
        raise HTTPException(status_code=400, detail=mail_result.get("error_msg", "读取邮件失败"))

    return MailListResponse(
        account_id=account.id,
        email=account.email,
        folder=folder.lower(),
        page=page,
        page_size=page_size,
        total=mail_result.get("total_emails", 0),
        items=mail_result.get("emails", []),
    )


def get_mail_detail(db: Session, account_id: int, folder: str, message_id: str) -> MailDetailResponse:
    account = get_account_or_404(db, account_id)
    target_folder = resolve_folder(folder)

    detail_result = get_email_detail_by_uid(
        email_address=account.email,
        refresh_token=account.refresh_token,
        client_id=account.client_id,
        target_uid=message_id,
        target_folder=target_folder,
    )

    if not detail_result.get("success"):
        refreshed = try_refresh_account_token(db, account)
        if refreshed:
            detail_result = get_email_detail_by_uid(
                email_address=account.email,
                refresh_token=account.refresh_token,
                client_id=account.client_id,
                target_uid=message_id,
                target_folder=target_folder,
            )

    if not detail_result.get("success"):
        raise HTTPException(status_code=400, detail=detail_result.get("error_msg", "读取邮件详情失败"))

    return MailDetailResponse(
        account_id=account.id,
        email=account.email,
        folder=folder.lower(),
        message_id=message_id,
        detail=detail_result.get("detail", {}),
    )
