import copy
import threading
import time


LIST_CACHE_TTL_SECONDS = 30
DETAIL_CACHE_TTL_SECONDS = 300


class _TimedCache:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._store: dict[tuple, tuple[float, object]] = {}

    def get(self, key: tuple):
        with self._lock:
            cached = self._store.get(key)
            if not cached:
                return None

            expires_at, payload = cached
            if expires_at <= time.monotonic():
                self._store.pop(key, None)
                return None

            return copy.deepcopy(payload)

    def set(self, key: tuple, value, ttl_seconds: int) -> None:
        with self._lock:
            self._store[key] = (time.monotonic() + ttl_seconds, copy.deepcopy(value))

    def delete(self, key: tuple) -> None:
        with self._lock:
            self._store.pop(key, None)


_list_cache = _TimedCache()
_detail_cache = _TimedCache()


def build_list_cache_key(account_id: int, folder: str, page: int, page_size: int) -> tuple:
    return ("mail-list", account_id, folder.lower(), page, page_size)


def build_detail_cache_key(account_id: int, folder: str, message_id: str) -> tuple:
    return ("mail-detail", account_id, folder.lower(), message_id)


def get_list_cache(account_id: int, folder: str, page: int, page_size: int):
    return _list_cache.get(build_list_cache_key(account_id, folder, page, page_size))


def set_list_cache(account_id: int, folder: str, page: int, page_size: int, payload) -> None:
    _list_cache.set(build_list_cache_key(account_id, folder, page, page_size), payload, LIST_CACHE_TTL_SECONDS)


def clear_list_cache(account_id: int, folder: str, page: int, page_size: int) -> None:
    _list_cache.delete(build_list_cache_key(account_id, folder, page, page_size))


def get_detail_cache(account_id: int, folder: str, message_id: str):
    return _detail_cache.get(build_detail_cache_key(account_id, folder, message_id))


def set_detail_cache(account_id: int, folder: str, message_id: str, payload) -> None:
    _detail_cache.set(build_detail_cache_key(account_id, folder, message_id), payload, DETAIL_CACHE_TTL_SECONDS)
