import os
from pathlib import Path

from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker


def _resolve_db_file() -> Path:
    appdata = os.getenv("LOCALAPPDATA")
    if appdata:
        base_dir = Path(appdata) / "ms-mail-fetcher"
    else:
        base_dir = Path.home() / ".ms-mail-fetcher"

    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir / "ms_mail_fetcher.db"


DATABASE_URL = f"sqlite:///{_resolve_db_file().as_posix()}"

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
