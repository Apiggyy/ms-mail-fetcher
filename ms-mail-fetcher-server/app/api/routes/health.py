from datetime import datetime

from fastapi import APIRouter, Depends
from sqlalchemy import func
from sqlalchemy.orm import Session

from app.db.database import get_db
from app.models.models import Account

router = APIRouter(tags=["health"])


@router.get("/api/health")
def health_check(db: Session = Depends(get_db)):
    total_accounts = db.query(func.count(Account.id)).scalar() or 0
    return {
        "ok": True,
        "total_accounts": total_accounts,
        "server_time": datetime.utcnow().isoformat(),
    }
