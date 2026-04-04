import logging
import time
from datetime import datetime, timedelta

from sqlalchemy.orm import Session

from app.models.models import Account
from app.utils.outlook_imap_client import refresh_oauth_token_manually


logger = logging.getLogger("ms_mail_fetcher")

TOKEN_REFRESH_BUFFER_SECONDS = 120


def _token_is_usable(account: Account) -> bool:
    if not account.access_token or not account.access_token_expires_at:
        return False

    safe_deadline = account.access_token_expires_at - timedelta(seconds=TOKEN_REFRESH_BUFFER_SECONDS)
    return datetime.utcnow() < safe_deadline


def _persist_refreshed_tokens(
    db: Session,
    account: Account,
    refresh_result: dict,
) -> dict:
    access_token = refresh_result.get("new_access_token")
    if not access_token:
        return {
            "success": False,
            "error_msg": refresh_result.get("error_msg", "微软接口未返回 access_token"),
        }

    expires_in = int(refresh_result.get("expires_in") or 0)
    if expires_in <= 0:
        expires_in = 3600

    rotated_refresh_token = refresh_result.get("new_refresh_token") or ""
    account.access_token = access_token
    account.access_token_expires_at = datetime.utcnow() + timedelta(seconds=expires_in)
    account.last_refresh_time = datetime.utcnow()

    refresh_token_rotated = False
    if rotated_refresh_token:
        refresh_token_rotated = rotated_refresh_token != account.refresh_token
        account.refresh_token = rotated_refresh_token

    db.add(account)
    db.commit()
    db.refresh(account)

    return {
        "success": True,
        "access_token": account.access_token,
        "expires_at": account.access_token_expires_at,
        "refresh_token_rotated": refresh_token_rotated,
        "token_source": "refresh",
    }


def get_access_token_for_account(
    db: Session,
    account: Account,
    *,
    force_refresh: bool = False,
) -> dict:
    start = time.perf_counter()

    if not force_refresh and _token_is_usable(account):
        duration_ms = (time.perf_counter() - start) * 1000
        return {
            "success": True,
            "access_token": account.access_token,
            "expires_at": account.access_token_expires_at,
            "refresh_token_rotated": False,
            "token_source": "cache",
            "duration_ms": duration_ms,
        }

    refresh_result = refresh_oauth_token_manually(account.client_id, account.refresh_token)
    persisted = _persist_refreshed_tokens(db, account, refresh_result)
    duration_ms = (time.perf_counter() - start) * 1000

    if not persisted.get("success"):
        logger.warning(
            "token.refresh_failed | account=%s | force=%s | duration_ms=%.2f | error=%s",
            account.email,
            force_refresh,
            duration_ms,
            persisted.get("error_msg", refresh_result.get("error_msg", "未知错误")),
        )
        return {
            "success": False,
            "error_msg": persisted.get("error_msg", refresh_result.get("error_msg", "刷新 Token 失败")),
            "token_source": "refresh",
            "duration_ms": duration_ms,
            "refresh_token_rotated": False,
        }

    persisted["duration_ms"] = duration_ms
    return persisted


def refresh_account_token_now(db: Session, account: Account) -> dict:
    return get_access_token_for_account(db, account, force_refresh=True)
