"""
database.py — SQLAlchemy 엔진 및 세션 팩토리

DB URL은 config.json에서 로드한다 (하드코딩 금지).
향후 PostgreSQL 전환을 고려하여 순수 SQL 사용을 지양한다.
"""

import logging
from pathlib import Path

from sqlalchemy import create_engine, event
from sqlalchemy.orm import DeclarativeBase, sessionmaker

from core.config_loader import get

logger = logging.getLogger(__name__)

_db_url: str = get("database.url")

# SQLite 파일 경로의 상위 디렉토리가 없으면 생성
if _db_url.startswith("sqlite:///"):
    db_path = Path(_db_url.replace("sqlite:///", ""))
    db_path.parent.mkdir(parents=True, exist_ok=True)

engine = create_engine(_db_url, echo=False)


# SQLite WAL 모드 및 외래키 활성화
@event.listens_for(engine, "connect")
def _set_sqlite_pragma(dbapi_conn, connection_record):
    if _db_url.startswith("sqlite"):
        cursor = dbapi_conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("PRAGMA foreign_keys=ON;")
        cursor.close()


SessionLocal = sessionmaker(bind=engine)


class Base(DeclarativeBase):
    """모든 ORM 모델의 베이스 클래스."""
    pass


def init_db() -> None:
    """등록된 모든 모델의 테이블을 생성한다."""
    import core.models  # noqa: F401 — 모델 등록을 위한 임포트
    Base.metadata.create_all(bind=engine)
    logger.info("데이터베이스 테이블 초기화 완료: %s", _db_url)
