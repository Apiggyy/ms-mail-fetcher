import logging

from sqlalchemy import text

from app.db.database import engine


logger = logging.getLogger("ms_mail_fetcher")


def ensure_sqlite_schema_compatibility() -> None:
    if engine.url.get_backend_name() != "sqlite":
        return

    with engine.begin() as connection:
        rows = connection.execute(text("PRAGMA table_info(accounts)")).mappings().all()
        if not rows:
            return

        existing_columns = {row["name"] for row in rows}
        if "access_token" not in existing_columns:
            connection.execute(text("ALTER TABLE accounts ADD COLUMN access_token VARCHAR"))
            logger.info("Schema migration applied: accounts.access_token")

        if "access_token_expires_at" not in existing_columns:
            connection.execute(text("ALTER TABLE accounts ADD COLUMN access_token_expires_at DATETIME"))
            logger.info("Schema migration applied: accounts.access_token_expires_at")
