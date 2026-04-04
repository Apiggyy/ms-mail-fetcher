import imaplib
import logging
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass

from app.utils.outlook_imap_client import IMAP_PORT, IMAP_SERVER


logger = logging.getLogger("ms_mail_fetcher")

SESSION_IDLE_TIMEOUT_SECONDS = 300
SESSION_MAX_LIFETIME_SECONDS = 600


@dataclass
class ImapSessionEntry:
    imap_conn: imaplib.IMAP4_SSL
    email: str
    access_token: str
    selected_folder: str | None
    created_at: float
    last_used_at: float


_sessions: dict[int, ImapSessionEntry] = {}
_session_lock = threading.RLock()


def _close_entry(entry: ImapSessionEntry) -> None:
    try:
        entry.imap_conn.close()
    except Exception:
        pass

    try:
        entry.imap_conn.logout()
    except Exception:
        pass


def _build_auth_string(email: str, access_token: str) -> bytes:
    return f"user={email}\1auth=Bearer {access_token}\1\1".encode("utf-8")


def _create_authenticated_session(email: str, access_token: str) -> tuple[ImapSessionEntry, float]:
    started_at = time.perf_counter()
    imap_conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    typ, _ = imap_conn.authenticate("XOAUTH2", lambda _: _build_auth_string(email, access_token))
    auth_duration_ms = (time.perf_counter() - started_at) * 1000
    if typ != "OK":
        try:
            imap_conn.logout()
        except Exception:
            pass
        raise RuntimeError("IMAP 认证失败")

    now = time.monotonic()
    return ImapSessionEntry(
        imap_conn=imap_conn,
        email=email,
        access_token=access_token,
        selected_folder=None,
        created_at=now,
        last_used_at=now,
    ), auth_duration_ms


def _session_expired(entry: ImapSessionEntry, now: float) -> bool:
    if now - entry.last_used_at > SESSION_IDLE_TIMEOUT_SECONDS:
        return True
    if now - entry.created_at > SESSION_MAX_LIFETIME_SECONDS:
        return True
    return False


def _session_alive(entry: ImapSessionEntry) -> bool:
    try:
        typ, _ = entry.imap_conn.noop()
        return typ == "OK"
    except Exception:
        return False


def _cleanup_expired_sessions_locked() -> None:
    now = time.monotonic()
    expired_ids = [account_id for account_id, entry in _sessions.items() if _session_expired(entry, now)]
    for account_id in expired_ids:
        entry = _sessions.pop(account_id, None)
        if entry is not None:
            _close_entry(entry)


@contextmanager
def acquire_imap_session(account_id: int, email: str, access_token: str):
    with _session_lock:
        _cleanup_expired_sessions_locked()

        existing_entry = _sessions.get(account_id)
        session_source = "reused"
        auth_duration_ms = 0.0
        recreated = False

        reusable = (
            existing_entry is not None
            and existing_entry.email == email
            and existing_entry.access_token == access_token
            and _session_alive(existing_entry)
        )

        if not reusable and existing_entry is not None:
            _sessions.pop(account_id, None)
            _close_entry(existing_entry)
            recreated = True
            existing_entry = None

        if existing_entry is None:
            existing_entry, auth_duration_ms = _create_authenticated_session(email, access_token)
            _sessions[account_id] = existing_entry
            session_source = "recreated" if recreated else "new"

        existing_entry.last_used_at = time.monotonic()
        try:
            yield {
                "entry": existing_entry,
                "session_source": session_source,
                "imap_auth_ms": auth_duration_ms,
            }
        finally:
            existing_entry.last_used_at = time.monotonic()


def release_broken_session(account_id: int) -> None:
    with _session_lock:
        entry = _sessions.pop(account_id, None)

    if entry is not None:
        _close_entry(entry)


def close_all_sessions() -> None:
    with _session_lock:
        entries = list(_sessions.values())
        _sessions.clear()

    for entry in entries:
        _close_entry(entry)
