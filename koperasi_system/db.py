import os
from contextlib import contextmanager

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.exc import SQLAlchemyError

from koperasi_system.settings import (
    DATABASE_URL,
    DB_ECHO,
    DB_MAX_OVERFLOW,
    DB_POOL_SIZE,
    DB_POOL_TIMEOUT,
    DB_USE_SSL,
)

_engine: Engine | None = None


def _build_connect_args() -> dict:
    if not DB_USE_SSL:
        return {}
    return {"sslmode": "require"}


def get_engine() -> Engine:
    global _engine
    if _engine is None:
        if not DATABASE_URL:
            raise RuntimeError("postgresql://neondb_owner:npg_sjeEU6b5cfHh@ep-late-paper-amohg8m3-pooler.c-5.us-east-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require.")
        _engine = create_engine(
            DATABASE_URL,
            pool_size=DB_POOL_SIZE,
            max_overflow=DB_MAX_OVERFLOW,
            pool_timeout=DB_POOL_TIMEOUT,
            echo=DB_ECHO,
            future=True,
            connect_args=_build_connect_args(),
        )
    return _engine


@contextmanager
def db_session():
    engine = get_engine()
    conn = engine.connect()
    trans = conn.begin()
    try:
        yield conn
        trans.commit()
    except Exception:
        trans.rollback()
        raise
    finally:
        conn.close()


def ping_database() -> bool:
    try:
        with db_session() as conn:
            conn.execute(text("select 1"))
        return True
    except SQLAlchemyError:
        return False


def init_db_schema(sql_path: str) -> None:
    if not os.path.exists(sql_path):
        raise FileNotFoundError(sql_path)
    with open(sql_path, "r", encoding="utf-8") as f:
        schema_sql = f.read().strip()
    if not schema_sql:
        raise ValueError("Schema SQL kosong.")
    with db_session() as conn:
        for statement in [s.strip() for s in schema_sql.split(";") if s.strip()]:
            conn.execute(text(statement))
