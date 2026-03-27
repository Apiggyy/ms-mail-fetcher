# SQLite 数据库设计文档（MS-Mail GPT Manager）

> 更新时间：2026-03-26  
> 适用后端：`ms-mail-fetcher-server`

## 1. 数据库说明

- 数据库类型：SQLite
- 默认连接串：`sqlite:///./ms_mail_fetcher.db`
- 数据库文件：运行后自动在后端根目录生成 `ms_mail_fetcher.db`
- 建表策略：服务启动时通过 SQLAlchemy 自动建表（`Base.metadata.create_all(bind=engine)`）

---

## 2. 核心表：`accounts`

| 字段名 | 类型 | 约束 | 默认值 | 说明 |
|---|---|---|---|---|
| `id` | INTEGER | PK, AUTOINCREMENT | - | 主键 |
| `email` | VARCHAR | NOT NULL, UNIQUE, INDEX | - | 邮箱地址（核心） |
| `password` | VARCHAR | NOT NULL | - | 邮箱密码（核心） |
| `client_id` | VARCHAR | NOT NULL | - | Microsoft Client ID |
| `refresh_token` | VARCHAR | NOT NULL | - | Microsoft Refresh Token |
| `last_refresh_time` | DATETIME | NOT NULL | `datetime.utcnow()` | 最后刷新时间 |
| `account_type` | VARCHAR | NULL | NULL | 账号类型编码（关联 `account_types.code`） |
| `remark` | VARCHAR | NULL | NULL | 备注 |
| `is_active` | BOOLEAN | NOT NULL, INDEX | `True` | 状态位：`1` 活跃池，`0` 归档池 |

---

## 3. 类型字典表：`account_types`

| 字段名 | 类型 | 约束 | 默认值 | 说明 |
|---|---|---|---|---|
| `id` | INTEGER | PK, AUTOINCREMENT | - | 主键 |
| `code` | VARCHAR | NOT NULL, UNIQUE, INDEX | - | 类型编码（如 `team/member/plus/idle`） |
| `label` | VARCHAR | NOT NULL | - | 前端展示文案 |
| `color` | VARCHAR | NOT NULL | `#409EFF` | 标签颜色（Hex，如 `#67C23A`） |

---

## 4. 对应 ORM 模型

文件：`domain/models.py`  
模型：`Account`

```python
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
```

---

## 5. SQLite 建表 SQL（参考）

```sql
CREATE TABLE accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email VARCHAR NOT NULL UNIQUE,
    password VARCHAR NOT NULL,
    client_id VARCHAR NOT NULL,
    refresh_token VARCHAR NOT NULL,
    last_refresh_time DATETIME NOT NULL,
    account_type VARCHAR NULL,
    remark VARCHAR NULL,
    is_active BOOLEAN NOT NULL DEFAULT 1
);

CREATE TABLE account_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code VARCHAR NOT NULL UNIQUE,
    label VARCHAR NOT NULL,
    color VARCHAR NOT NULL DEFAULT '#409EFF'
);

CREATE INDEX ix_accounts_id ON accounts (id);
CREATE INDEX ix_accounts_email ON accounts (email);
CREATE INDEX ix_accounts_is_active ON accounts (is_active);
CREATE INDEX ix_account_types_id ON account_types (id);
CREATE INDEX ix_account_types_code ON account_types (code);
```

---

## 6. 状态流转说明

- 活跃账号页（第一页）：`is_active = 1`
- 归档账号页（第二页）：`is_active = 0`
- 软删除（归档）：更新 `is_active = 0`
- 硬删除：`DELETE FROM accounts WHERE id = ?`

账号类型说明：
- 启动时会自动补齐默认类型：`team/member/plus/idle`（其中 `idle` 前端显示为 `Idle`）
- 类型编码存储于 `accounts.account_type`
- 类型颜色存储于 `account_types.color`，用于前端标签颜色展示

---

## 7. 计算字段说明（不落库）

前端显示 `days_since_refresh` 不是数据库字段。  
由后端在返回时动态计算：

```python
days_since_refresh = max((datetime.utcnow() - account.last_refresh_time).days, 0)
```

---

## 8. 初始化与验证

### 启动后端（自动建库建表）

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

### 检查 SQLite 文件是否生成

- 文件路径：`ms-mail-fetcher-server/ms_mail_fetcher.db`

### 查看表结构（sqlite3）

```sql
.schema accounts
.schema account_types
PRAGMA table_info(accounts);
PRAGMA table_info(account_types);
```

---

## 9. 备份建议

SQLite 备份可直接复制数据库文件：

- 备份源：`ms_mail_fetcher.db`
- 建议频率：每日/每次批量导入后
- 建议保留：最近 7~30 个版本
