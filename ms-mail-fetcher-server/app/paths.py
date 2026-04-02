import os
from pathlib import Path


APP_DATA_DIR_NAME = "ms-mail-fetcher"
LEGACY_HOME_DIR_NAME = ".ms-mail-fetcher"


def resolve_data_dir() -> Path:
    explicit_dir = os.getenv("DATA_DIR")
    if explicit_dir:
        base_dir = Path(explicit_dir).expanduser()
    else:
        appdata = os.getenv("LOCALAPPDATA")
        if appdata:
            base_dir = Path(appdata) / APP_DATA_DIR_NAME
        else:
            base_dir = Path.home() / LEGACY_HOME_DIR_NAME

    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir


def resolve_data_file(file_name: str) -> Path:
    return resolve_data_dir() / file_name
