from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from app.crud.account_types import (
    create_account_type,
    delete_account_type,
    list_account_types,
    update_account_type,
)
from app.db.database import get_db
from app.schemas.schemas import AccountTypeCreate, AccountTypeOut, AccountTypeUpdate

router = APIRouter(prefix="/api/account-types", tags=["account-types"])


@router.get("", response_model=list[AccountTypeOut])
def list_account_types_route(db: Session = Depends(get_db)):
    return list_account_types(db=db)


@router.post("", response_model=AccountTypeOut)
def create_account_type_route(payload: AccountTypeCreate, db: Session = Depends(get_db)):
    return create_account_type(db=db, payload=payload)


@router.put("/{account_type_id}", response_model=AccountTypeOut)
def update_account_type_route(account_type_id: int, payload: AccountTypeUpdate, db: Session = Depends(get_db)):
    return update_account_type(db=db, account_type_id=account_type_id, payload=payload)


@router.delete("/{account_type_id}")
def delete_account_type_route(account_type_id: int, db: Session = Depends(get_db)):
    return delete_account_type(db=db, account_type_id=account_type_id)
