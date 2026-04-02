import json

from fastapi import APIRouter

from app.paths import resolve_data_file
from app.schemas.schemas import UiPreferences, UiPreferencesUpdate

router = APIRouter(prefix="/api/ui", tags=["ui-preferences"])


def _resolve_preferences_file():
    return resolve_data_file("ui_preferences.json")


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
    merged = UiPreferences(
        sidebar_collapsed=(
            payload.sidebar_collapsed
            if payload.sidebar_collapsed is not None
            else current.sidebar_collapsed
        ),
        window_width=(
            payload.window_width
            if payload.window_width is not None
            else current.window_width
        ),
        window_height=(
            payload.window_height
            if payload.window_height is not None
            else current.window_height
        ),
    )
    _write_preferences(merged)
    return merged
