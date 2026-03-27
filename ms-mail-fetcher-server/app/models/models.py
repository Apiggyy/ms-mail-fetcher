from datetime import datetime

from sqlalchemy import Boolean, Column, DateTime, Integer, String

from app.db.database import Base


class Account(Base):
    __tablename__ = "accounts"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String, nullable=False, unique=True, index=True)
    password = Column(String, nullable=False)
    client_id = Column(String, nullable=False)
    refresh_token = Column(String, nullable=False)
    last_refresh_time = Column(DateTime, nullable=False, default=datetime.utcnow)
    account_type = Column(String, nullable=True)
    remark = Column(String, nullable=True)
    is_active = Column(Boolean, nullable=False, default=True, index=True)


class AccountType(Base):
    __tablename__ = "account_types"

    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, nullable=False, unique=True, index=True)
    label = Column(String, nullable=False)
    color = Column(String, nullable=False, default="#409EFF")
