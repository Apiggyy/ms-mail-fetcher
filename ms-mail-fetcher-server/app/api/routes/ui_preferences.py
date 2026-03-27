import json
import os
from pathlib import Path

from fastapi import APIRouter

from app.schemas.schemas import UiPreferences, UiPreferencesUpdate

router = APIRouter(prefix="/api/ui", tags=["ui-preferences"])


def _resolve_preferences_file() -> Path:
    appdata = os.getenv("LOCALAPPDATA")
    if appdata:
        base_dir = Path(appdata) / "ms-mail-fetcher"
    else:
        base_dir = Path.home() / ".ms-mail-fetcher"

    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir / "ui_preferences.json"


def _read_preferences() -> UiPreferences:
    target = _resolve_preferences_file()
    if not target.exists():
        return UiPreferences()

    try:
        with target.open("r", encoding="utf-8") as f:
            raw = json.load(f)
        if isinstance(raw, dict):
            return UiPreferences(**raw)
    except Exception:
        return UiPreferences()

    return UiPreferences()


def _write_preferences(payload: UiPreferences) -> None:
    target = _resolve_preferences_file()
    with target.open("w", encoding="utf-8") as f:
        json.dump(payload.model_dump(), f, ensure_ascii=False, indent=2)


@router.get("/preferences", response_model=UiPreferences)
def get_ui_preferences():
    return _read_preferences()


@router.put("/preferences", response_model=UiPreferences)
def update_ui_preferences(payload: UiPreferencesUpdate):
    current = _read_preferences()
    merged = UiPreferences(sidebar_collapsed=payload.sidebar_collapsed if payload.sidebar_collapsed is not None else current.sidebar_collapsed)
    _write_preferences(merged)
    return merged
