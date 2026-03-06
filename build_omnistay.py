"""
build_omnistay.py — OmniStay Core 자가 증식 스크립트

이 스크립트 하나를 로컬 윈도우 PC에서 실행하면
omnistay-core/ 폴더와 모든 하위 모듈이 자동 생성된다.

사용법:
    python build_omnistay.py
"""

import os
from pathlib import Path

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 루트 디렉토리
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ROOT = Path(os.getcwd()) / "omnistay-core"

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 파일 정의: (상대경로, 내용)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FILES: dict[str, str] = {}

# ──────────────────────────────────────────────────────────────
# 1. CLAUDE.md
# ──────────────────────────────────────────────────────────────
FILES["CLAUDE.md"] = r"""# CLAUDE.md — OmniStay Core

## Project Overview

**Repository:** Ait426/omnistay

OmniStay is a B2B SaaS integration platform that extracts reservation and revenue data from multiple external PMS (Property Management Systems), normalizes it to a global standard format, and asynchronously writes it to Excel ledgers and other outputs.

**Core philosophy:** Modularity, zero single-points-of-failure (SPOF), scalability, data integrity.

**Critical warning:** The `hotel-automation` folder (legacy ~30k-line codebase) is used for live operations. **NEVER scan, read, or modify it.** All work happens exclusively inside `omnistay-core`.

## System Architecture

The system is strictly isolated into 3 layers with a **one-way data pipeline**:

```
Adapters  ──▶  Core DB  ──▶  Workers
(collect)      (normalize)    (output)
```

### Directory Structure

```
omnistay-core/
├── adapters/       # Data collection — PMS integrations & web scraping
│                   # Fetches data and passes it to Core
│                   # ⛔ No business logic allowed here
│
├── core/           # Central control — the system's heart
│   ├── models/     # SQLAlchemy ORM models (DB schema)
│   ├── normalize/  # Global data normalization logic
│   └── config_loader.py  # Reads from /config/config.json
│
├── workers/        # One-way output — consumes Core DB data
│                   # Writes to Excel, sends notifications, etc.
│                   # ⛔ Must run independently; no knowledge of other workers or adapters
│
└── config/
    └── config.json # All environment/runtime settings (NEVER hardcode values)
```

### Data Flow Rules

- **Adapters → Core DB → Workers** (one direction only)
- Adapters CANNOT call Workers directly
- Workers CANNOT call or depend on Adapters
- All cross-layer communication goes through the Core DB

## Tech Stack

| Component          | Technology                  | Notes                                        |
| ------------------ | --------------------------- | -------------------------------------------- |
| Language           | Python 3.14+                |                                              |
| Database           | SQLite via SQLAlchemy ORM   | No raw SQL — design for future PostgreSQL migration |
| Excel Automation   | xlwings (COM control)       | Preserves formulas; no openpyxl for writes   |

## Strict Rules for AI Assistants

### Rule 1: No Hardcoding

All configurable values MUST be loaded from `/config/config.json` via `config_loader.py`:

- Excel cell coordinates (e.g., `B4`, `K25`)
- Timezone identifiers (e.g., `Asia/Seoul`)
- Currency codes (e.g., `KRW`)
- Any value that could change between deployments

**Never** embed these directly in Python source files.

### Rule 2: Global Data Standards

**Time:** All timestamps stored in the DB (`check_in_time`, `check_out_time`, etc.) MUST be in **UTC**. Convert to local timezone only at the presentation/output layer (display, Excel writes).

**Currency:** All monetary amounts (`total_amount`, etc.) MUST use Python's `Decimal` type. Every amount column MUST be paired with an ISO 4217 currency code column (e.g., `currency="KRW"`).

**Reservation status:** Only these 5 enum values are allowed system-wide:

| Enum Value          | Meaning         |
| ------------------- | --------------- |
| `EXPECTED_CHECKIN`  | Expected arrival |
| `CHECKED_IN`        | Guest checked in |
| `EXPECTED_CHECKOUT` | Expected departure |
| `CHECKED_OUT`       | Guest checked out |
| `CANCELED`          | Reservation canceled |

External PMS status codes (e.g., Yanolja's `IH`, `SO`, `CO`) MUST be mapped to one of these 5 values inside the adapter layer before reaching Core.

### Rule 3: Fail-Fast & Explicit Logging

- **Never** swallow errors with bare `try-except: pass`
- If Excel is locked (access denied) or data is empty, log the error clearly and halt the process immediately
- Design for operator awareness — silent failures are forbidden

### Rule 4: One-Way Data Pipeline

Enforce the strict flow: `Adapters → Core DB → Workers`

- Adapters must not import from or call Workers
- Workers must not import from or call Adapters
- Core is the only shared dependency

## Development Workflow

### Branch Strategy

- Development branches follow the pattern `claude/<description>-<session-id>`
- Always push with `git push -u origin <branch-name>`
- Never force-push without explicit permission

### Commit Conventions

- Write clear, descriptive commit messages
- Use imperative mood in commit subjects (e.g., "Add feature" not "Added feature")
- Keep subject lines under 72 characters

## Build & Test Commands

_No build or test commands configured yet. Update this section when a build system is added._

## General AI Assistant Guidelines

1. **Read before writing** — Always read existing files before modifying them
2. **Minimal changes** — Only make changes that are directly requested or clearly necessary
3. **No over-engineering** — Keep solutions simple; avoid abstractions for hypothetical future needs
4. **Security first** — Never introduce injection, XSS, or other OWASP Top 10 vulnerabilities
5. **Never touch `hotel-automation`** — That folder is off-limits, period
6. **Update this file** — When adding infrastructure (build tools, test frameworks, CI/CD), update the relevant sections here
"""

# ──────────────────────────────────────────────────────────────
# 2. config/config.json
# ──────────────────────────────────────────────────────────────
FILES["config/config.json"] = r"""{
  "database": {
    "url": "sqlite:///data/omnistay.db"
  },
  "timezone": {
    "local": "Asia/Seoul"
  },
  "currency": {
    "default": "KRW"
  },
  "scheduler": {
    "interval_minutes": 10
  },
  "adapters": {
    "yanolja": {
      "base_url": "https://pms-mashup.yflux.biz/v1/properties/{property_id}/",
      "property_id": "YOUR_PROPERTY_ID",
      "access_token": "YOUR_ACCESS_TOKEN",
      "endpoints": {
        "reservations": "reservations"
      }
    }
  },
  "excel": {
    "ledger_path": "C:/Users/OmniStay/Desktop/장부.xlsx",
    "sheet_name": "예약현황",
    "start_cell": "B4",
    "columns": {
      "guest_name": "B",
      "room_number": "C",
      "room_type": "D",
      "check_in_time": "E",
      "check_out_time": "F",
      "total_amount": "G",
      "currency": "H",
      "status": "I",
      "source_pms": "J",
      "external_reservation_id": "K"
    }
  }
}
"""

# ──────────────────────────────────────────────────────────────
# 3. __init__.py (빈 파일 3개)
# ──────────────────────────────────────────────────────────────
FILES["core/__init__.py"] = ""
FILES["adapters/__init__.py"] = ""
FILES["workers/__init__.py"] = ""

# ──────────────────────────────────────────────────────────────
# 4. core/config_loader.py
# ──────────────────────────────────────────────────────────────
FILES["core/config_loader.py"] = r'''"""
config_loader.py — /config/config.json 로더

모든 설정값은 이 모듈을 통해서만 접근한다.
하드코딩 금지 (CLAUDE.md Rule 1).
"""

import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "config.json"
_config_cache: dict | None = None


def load_config() -> dict:
    """config.json을 읽어 딕셔너리로 반환한다. 실패 시 즉시 중단(Fail-Fast)."""
    global _config_cache
    if _config_cache is not None:
        return _config_cache

    if not _CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"설정 파일을 찾을 수 없습니다: {_CONFIG_PATH}"
        )

    try:
        text = _CONFIG_PATH.read_text(encoding="utf-8")
        _config_cache = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"config.json 파싱 실패: {exc}"
        ) from exc

    logger.info("config.json 로드 완료: %s", _CONFIG_PATH)
    return _config_cache


def get(key: str, *, default=None):
    """점(.) 구분 키로 중첩 값을 조회한다. 예: get('database.url')"""
    cfg = load_config()
    parts = key.split(".")
    node = cfg
    for part in parts:
        if not isinstance(node, dict) or part not in node:
            if default is not None:
                return default
            raise KeyError(f"config.json에 '{key}' 키가 없습니다.")
        node = node[part]
    return node
'''

# ──────────────────────────────────────────────────────────────
# 5. core/database.py
# ──────────────────────────────────────────────────────────────
FILES["core/database.py"] = r'''"""
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
'''

# ──────────────────────────────────────────────────────────────
# 6. core/models.py
# ──────────────────────────────────────────────────────────────
FILES["core/models.py"] = r'''"""
models.py — SQLAlchemy ORM 모델 정의

글로벌 데이터 표준 (CLAUDE.md Rule 2):
  - 시간: UTC로만 저장
  - 통화: Decimal 타입 + ISO 4217 통화 코드
  - 상태: ReservationStatus Enum (5가지만 허용)
"""

import enum
from datetime import datetime, timezone
from decimal import Decimal

from sqlalchemy import Boolean, DateTime, Enum, Integer, Numeric, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from core.database import Base


class ReservationStatus(enum.Enum):
    """예약 상태 — 시스템 전체에서 이 5가지만 허용된다."""

    EXPECTED_CHECKIN = "EXPECTED_CHECKIN"
    CHECKED_IN = "CHECKED_IN"
    EXPECTED_CHECKOUT = "EXPECTED_CHECKOUT"
    CHECKED_OUT = "CHECKED_OUT"
    CANCELED = "CANCELED"


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


class Reservation(Base):
    """예약 정보 — 어댑터가 수집한 데이터의 정규화된 형태."""

    __tablename__ = "reservations"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)

    # 외부 PMS 원본 식별자
    source_pms: Mapped[str] = mapped_column(String(50), nullable=False)
    external_reservation_id: Mapped[str] = mapped_column(String(100), nullable=False)

    # 투숙객 정보
    guest_name: Mapped[str] = mapped_column(String(200), nullable=False)
    room_number: Mapped[str | None] = mapped_column(String(50), nullable=True)
    room_type: Mapped[str | None] = mapped_column(String(100), nullable=True)

    # 시간 — 반드시 UTC (Rule 2)
    check_in_time: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False
    )
    check_out_time: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False
    )

    # 금액 — Decimal + ISO 4217 통화 코드 (Rule 2)
    total_amount: Mapped[Decimal] = mapped_column(
        Numeric(precision=15, scale=2), nullable=False
    )
    currency: Mapped[str] = mapped_column(String(3), nullable=False)

    # 상태 — 5가지 Enum만 허용 (Rule 2)
    status: Mapped[ReservationStatus] = mapped_column(
        Enum(ReservationStatus, native_enum=False, length=20),
        nullable=False,
    )

    # 엑셀 기입 상태 추적 — 중복 기입 방지
    is_exported_to_excel: Mapped[bool] = mapped_column(
        Boolean, default=False, nullable=False
    )

    # 메타
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=_utcnow, nullable=False
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=_utcnow, onupdate=_utcnow, nullable=False
    )

    def __repr__(self) -> str:
        return (
            f"<Reservation(id={self.id}, pms={self.source_pms!r}, "
            f"guest={self.guest_name!r}, status={self.status.value})>"
        )
'''

# ──────────────────────────────────────────────────────────────
# 7. adapters/yanolja_scraper.py
# ──────────────────────────────────────────────────────────────
FILES["adapters/yanolja_scraper.py"] = r'''"""
yanolja_scraper.py — 야놀자 플럭스 PMS 예약 데이터 수집 어댑터

단방향 파이프라인의 최초 입력 단계 (Yanolja API → Core DB).
⛔ 비즈니스 로직 금지 — 데이터 수집·정규화·DB 삽입만 수행한다.

통신 방식 (Critical Rule):
  야놀자 WAF가 requests/aiohttp를 차단하므로, 반드시 Playwright
  Headless 브라우저의 page.evaluate()를 통해 브라우저 네이티브
  fetch() API로만 데이터를 추출한다.
"""

import asyncio
import logging
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from zoneinfo import ZoneInfo

from playwright.async_api import async_playwright

from core.config_loader import get
from core.database import SessionLocal, init_db
from core.models import Reservation, ReservationStatus

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


# ── 설정값 로드 (Rule 1: 하드코딩 절대 금지) ──────────────────────
_BASE_URL: str = get("adapters.yanolja.base_url")
_PROPERTY_ID: str = get("adapters.yanolja.property_id")
_ACCESS_TOKEN: str = get("adapters.yanolja.access_token")
_ENDPOINTS: dict = get("adapters.yanolja.endpoints")
_KST = ZoneInfo(get("timezone.local"))

# ── 야놀자 상태 코드 → OmniStay Enum 매핑 (어댑터 내부 하드코딩) ──
_STATUS_MAP: dict[str, str] = {
    "IH": "CHECKED_IN",
    "SO": "CHECKED_OUT",
    "CI": "EXPECTED_CHECKIN",
    "CO": "EXPECTED_CHECKOUT",
}

# ── 야놀자 채널 코드 매핑 ─────────────────────────────────────────
_CHANNEL_MAP: dict[str, str] = {
    "워크인": "WLK",
    "야놀자모텔/호텔": "YGN",
    "여기어때": "YGI",
    "아고다": "AGD",
}
_CHANNEL_DEFAULT = "ETC"


def _build_api_url(endpoint_key: str) -> str:
    """config에서 로드한 base_url과 endpoint를 조합하여 완전한 URL을 생성한다."""
    base = _BASE_URL.replace("{property_id}", _PROPERTY_ID)
    if not base.endswith("/"):
        base += "/"
    return base + _ENDPOINTS[endpoint_key]


def _parse_kst_to_utc(kst_string: str) -> datetime:
    """야놀자 KST 시간 문자열을 파싱하여 UTC datetime으로 변환한다 (Rule 2)."""
    # 야놀자 일반 포맷: "2026-03-06 14:00:00" 또는 "2026-03-06T14:00:00"
    normalized = kst_string.replace("T", " ").strip()
    try:
        kst_dt = datetime.strptime(normalized, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        # 초 단위 없는 포맷 대응: "2026-03-06 14:00"
        kst_dt = datetime.strptime(normalized, "%Y-%m-%d %H:%M")

    kst_aware = kst_dt.replace(tzinfo=_KST)
    return kst_aware.astimezone(ZoneInfo("UTC"))


def _map_status(yanolja_code: str) -> ReservationStatus | None:
    """야놀자 상태 코드를 OmniStay ReservationStatus로 변환한다.

    매핑 불가 시 None을 반환하고 에러를 로깅한다.
    """
    mapped = _STATUS_MAP.get(yanolja_code)
    if mapped is None:
        logger.error("알 수 없는 야놀자 상태 코드 '%s' — 해당 행 무시", yanolja_code)
        return None

    try:
        return ReservationStatus(mapped)
    except ValueError:
        logger.error(
            "상태 코드 '%s' → '%s' 매핑 결과가 유효한 Enum이 아닙니다 — 해당 행 무시",
            yanolja_code, mapped,
        )
        return None


def _map_channel(channel_name: str) -> str:
    """야놀자 채널명을 단축 코드로 변환한다. 미매핑 시 'ETC'."""
    return _CHANNEL_MAP.get(channel_name, _CHANNEL_DEFAULT)


def _normalize_reservation(raw: dict) -> Reservation | None:
    """야놀자 원시 데이터 1건을 Reservation ORM 객체로 정규화한다.

    정규화 실패 시 None을 반환하고 에러를 로깅한다 (해당 행 무시).
    """
    reservation_id = raw.get("reservationNo", raw.get("id", "UNKNOWN"))

    # 상태 매핑
    status = _map_status(raw.get("stayStatus", ""))
    if status is None:
        return None

    # 금액 → Decimal 변환 (Rule 2: 통화)
    try:
        total_amount = Decimal(str(raw.get("salePrice", 0)))
    except (InvalidOperation, TypeError) as exc:
        logger.error(
            "예약 %s: 금액 변환 실패 (%s) — 해당 행 무시", reservation_id, exc,
        )
        return None

    # 시간 → UTC 변환 (Rule 2: 시간)
    try:
        check_in_time = _parse_kst_to_utc(raw["checkInDate"])
        check_out_time = _parse_kst_to_utc(raw["checkOutDate"])
    except (KeyError, ValueError) as exc:
        logger.error(
            "예약 %s: 시간 파싱 실패 (%s) — 해당 행 무시", reservation_id, exc,
        )
        return None

    channel = _map_channel(raw.get("channelName", ""))

    return Reservation(
        source_pms="YANOLJA",
        external_reservation_id=str(reservation_id),
        guest_name=raw.get("guestName", "N/A"),
        room_number=raw.get("roomNo"),
        room_type=raw.get("roomTypeName"),
        check_in_time=check_in_time,
        check_out_time=check_out_time,
        total_amount=total_amount,
        currency="KRW",
        status=status,
        notes=f"channel={channel}",
    )


async def _fetch_via_playwright(url: str) -> list[dict]:
    """Playwright Headless 브라우저에서 네이티브 fetch()로 데이터를 추출한다.

    야놀자 WAF가 requests/aiohttp를 차단하므로, page.evaluate() 내에서
    브라우저 컨텍스트의 fetch()를 호출하여 우회한다.
    """
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        try:
            page = await browser.new_page()

            # ──────────────────────────────────────────────────
            # WAF 우회 핵심: page.evaluate()로 브라우저 네이티브
            # fetch() API를 직접 호출한다.
            # JavaScript 컨텍스트에서 실행되므로 브라우저 핑거프린트,
            # 쿠키, TLS 스택이 모두 실제 브라우저와 동일하게 동작한다.
            # ──────────────────────────────────────────────────
            result = await page.evaluate(
                """
                async ([url, token]) => {
                    const response = await fetch(url, {
                        method: 'GET',
                        headers: {
                            'Authorization': `Bearer ${token}`,
                            'Content-Type': 'application/json',
                        },
                    });
                    if (!response.ok) {
                        return {
                            error: true,
                            status: response.status,
                            statusText: response.statusText,
                        };
                    }
                    return await response.json();
                }
                """,
                [url, _ACCESS_TOKEN],
            )

            if isinstance(result, dict) and result.get("error"):
                raise RuntimeError(
                    f"야놀자 API 응답 에러: "
                    f"HTTP {result['status']} {result['statusText']}"
                )

            # API 응답 구조: 최상위 리스트 또는 { "data": [...] }
            if isinstance(result, list):
                return result
            if isinstance(result, dict) and "data" in result:
                return result["data"]

            logger.warning("예상치 못한 응답 구조: %s", type(result).__name__)
            return result if isinstance(result, list) else []

        finally:
            await browser.close()


async def _scrape_and_store() -> None:
    """야놀자 예약 데이터를 수집하여 Core DB에 삽입한다."""
    url = _build_api_url("reservations")
    logger.info("야놀자 API 호출 시작: %s", url)

    raw_reservations = await _fetch_via_playwright(url)
    logger.info("야놀자 원시 데이터 %d건 수신", len(raw_reservations))

    session = SessionLocal()
    try:
        inserted = 0
        skipped = 0

        for raw in raw_reservations:
            reservation = _normalize_reservation(raw)
            if reservation is None:
                skipped += 1
                continue
            session.add(reservation)
            inserted += 1

        session.commit()
        logger.info(
            "DB 삽입 완료: %d건 성공, %d건 스킵 (정규화 실패)",
            inserted, skipped,
        )
    except Exception as exc:
        session.rollback()
        logger.error("DB 삽입 실패 — 롤백 완료: %s", exc)
        raise
    finally:
        session.close()


async def scrape_and_store() -> None:
    """외부(main.py 등)에서 호출 가능한 비동기 진입점."""
    init_db()
    await _scrape_and_store()


def run() -> None:
    """어댑터 진입점 — 동기 컨텍스트에서 호출 가능."""
    init_db()
    try:
        asyncio.run(_scrape_and_store())
    except Exception as exc:
        logger.error("야놀자 스크래퍼 치명적 오류: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    run()
'''

# ──────────────────────────────────────────────────────────────
# 8. workers/excel_writer.py
# ──────────────────────────────────────────────────────────────
FILES["workers/excel_writer.py"] = r'''"""
excel_writer.py — 엑셀 장부 기입 워커

단방향 파이프라인의 최종 출력 단계 (Core DB → Excel).
is_exported_to_excel == False 인 예약만 조회하여 xlwings로 기입한다.

치명적 제약 조건:
  - xlwings(COM) 전용: 기존 수식·서식 보존 (openpyxl 사용 금지)
  - 하드코딩 금지: 모든 경로·좌표는 config.json에서 로드 (Rule 1)
  - COM 에러 시 Fail-Fast: DB 플래그를 절대 변경하지 않고 즉시 종료 (Rule 3)
"""

import logging
import sys
from datetime import timezone
from zoneinfo import ZoneInfo

import xlwings as xw
from pywintypes import com_error
from sqlalchemy import select

from core.config_loader import get
from core.database import SessionLocal
from core.models import Reservation

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


# ── 설정값 로드 (Rule 1: 하드코딩 절대 금지) ──────────────────────
_LEDGER_PATH: str = get("excel.ledger_path")
_SHEET_NAME: str = get("excel.sheet_name")
_START_CELL: str = get("excel.start_cell")
_COLUMNS: dict[str, str] = get("excel.columns")
_LOCAL_TZ = ZoneInfo(get("timezone.local"))


def _parse_start_row(start_cell: str) -> int:
    """시작 셀 문자열(예: 'B4')에서 행 번호(4)를 추출한다."""
    digits = "".join(ch for ch in start_cell if ch.isdigit())
    if not digits:
        raise ValueError(f"시작 셀에서 행 번호를 추출할 수 없습니다: {start_cell!r}")
    return int(digits)


def _find_next_empty_row(sheet: xw.Sheet, start_row: int, column: str) -> int:
    """지정된 열에서 start_row부터 아래로 탐색하여 빈 행 번호를 반환한다."""
    row = start_row
    while sheet.range(f"{column}{row}").value is not None:
        row += 1
    return row


def _to_local(utc_dt, local_tz: ZoneInfo):
    """UTC datetime을 로컬 시간대로 변환한다 (출력 계층에서만 변환 — Rule 2)."""
    if utc_dt.tzinfo is None:
        utc_dt = utc_dt.replace(tzinfo=timezone.utc)
    return utc_dt.astimezone(local_tz)


def _write_reservations_to_excel(reservations: list[Reservation]) -> None:
    """xlwings COM을 통해 예약 데이터를 엑셀에 기입한다.

    COM 에러 발생 시 예외를 상위로 전파하여 DB 플래그가 변경되지 않도록 한다.
    """
    if not _LEDGER_PATH:
        raise FileNotFoundError(
            "config.json의 'excel.ledger_path'가 비어 있습니다. "
            "엑셀 파일 경로를 설정하세요."
        )

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(_LEDGER_PATH)
        sheet = wb.sheets[_SHEET_NAME]

        start_row = _parse_start_row(_START_CELL)
        first_col = next(iter(_COLUMNS.values()))
        current_row = _find_next_empty_row(sheet, start_row, first_col)

        logger.info(
            "엑셀 기입 시작: %s [%s] — %d건, 시작 행: %d",
            _LEDGER_PATH, _SHEET_NAME, len(reservations), current_row,
        )

        for res in reservations:
            row_data = {
                "guest_name": res.guest_name,
                "room_number": res.room_number or "",
                "room_type": res.room_type or "",
                "check_in_time": _to_local(res.check_in_time, _LOCAL_TZ),
                "check_out_time": _to_local(res.check_out_time, _LOCAL_TZ),
                "total_amount": float(res.total_amount),
                "currency": res.currency,
                "status": res.status.value,
                "source_pms": res.source_pms,
                "external_reservation_id": res.external_reservation_id,
            }

            for field_name, col_letter in _COLUMNS.items():
                cell_ref = f"{col_letter}{current_row}"
                sheet.range(cell_ref).value = row_data[field_name]

            current_row += 1

        wb.save()
        wb.close()
        logger.info("엑셀 저장 완료: %d건 기입", len(reservations))

    except com_error as exc:
        # ──────────────────────────────────────────────────────────
        # COM 에러 방어 (Rule 3: Fail-Fast)
        #
        # 엑셀 파일이 열려 있거나 편집 중일 때 발생한다.
        # 이 시점에서 DB의 is_exported_to_excel은 아직 False이므로
        # 다음 스케줄러 실행 시 동일 데이터를 다시 시도할 수 있다.
        # ──────────────────────────────────────────────────────────
        logger.error(
            "COM 에러 — 엑셀 파일 접근 실패 (파일이 열려 있을 수 있음): %s",
            exc,
        )
        logger.error(
            "DB 플래그(is_exported_to_excel)는 변경되지 않았습니다. "
            "다음 실행 시 재시도됩니다."
        )
        sys.exit(1)

    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


def run() -> None:
    """미기입 예약을 조회 → 엑셀 기입 → DB 플래그 업데이트.

    엑셀 기입이 성공한 경우에만 is_exported_to_excel = True로 변경한다.
    기입 중 어떤 예외라도 발생하면 DB는 일절 변경되지 않는다.
    """
    session = SessionLocal()
    try:
        stmt = (
            select(Reservation)
            .where(Reservation.is_exported_to_excel == False)  # noqa: E712
            .order_by(Reservation.check_in_time)
        )
        reservations = list(session.scalars(stmt))

        if not reservations:
            logger.info("기입할 새 예약이 없습니다.")
            return

        logger.info("미기입 예약 %d건 조회 완료", len(reservations))

        # 엑셀 기입 — 실패 시 예외 전파, DB 플래그 미변경
        _write_reservations_to_excel(reservations)

        # 엑셀 기입 성공 후에만 DB 플래그 업데이트
        for res in reservations:
            res.is_exported_to_excel = True
        session.commit()
        logger.info("DB 플래그 업데이트 완료: %d건 → is_exported_to_excel=True", len(reservations))

    except SystemExit:
        # sys.exit(1)은 그대로 전파
        raise
    except Exception as exc:
        session.rollback()
        logger.error("예상치 못한 오류 — DB 롤백 완료: %s", exc)
        sys.exit(1)
    finally:
        session.close()


if __name__ == "__main__":
    run()
'''

# ──────────────────────────────────────────────────────────────
# 9. main.py
# ──────────────────────────────────────────────────────────────
FILES["main.py"] = r'''"""
main.py — OmniStay 중앙 오케스트레이터

전체 파이프라인(Adapters → Core DB → Workers)을 관장하는
asyncio 기반 무한 루프 스케줄러.

실행 순서 (매 주기):
  1) yanolja_scraper  → Core DB에 예약 데이터 스크래핑·저장
  2) excel_writer     → 미기입 예약(is_exported_to_excel=False)을 엑셀에 기입

내결함성(Fault Tolerance):
  개별 단계의 에러가 메인 루프를 죽이지 않는다.
  에러를 로깅하고 다음 주기에서 재시도한다.
"""

import asyncio
import logging
import sys
from datetime import datetime, timezone

from core.config_loader import get
from core.database import init_db

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("omnistay.scheduler")


def _get_interval_seconds() -> int:
    """config.json에서 스케줄러 실행 주기를 읽어 초 단위로 반환한다."""
    minutes = get("scheduler.interval_minutes", default=10)
    return int(minutes) * 60


async def _run_adapter_phase() -> None:
    """Phase 1: 야놀자 스크래퍼로 데이터 수집 → Core DB 저장."""
    from adapters.yanolja_scraper import scrape_and_store

    logger.info("── Phase 1: 야놀자 스크래퍼 시작 ──")
    await scrape_and_store()
    logger.info("── Phase 1: 야놀자 스크래퍼 완료 ──")


def _run_worker_phase() -> None:
    """Phase 2: 엑셀 워커로 미기입 예약을 장부에 기입."""
    from workers.excel_writer import run as excel_run

    logger.info("── Phase 2: 엑셀 워커 시작 ──")
    excel_run()
    logger.info("── Phase 2: 엑셀 워커 완료 ──")


async def _execute_pipeline() -> None:
    """단일 파이프라인 주기를 실행한다.

    각 단계를 독립적으로 실행하여, 하나가 실패해도
    에러를 상위로 전파하되 메인 루프는 보호한다.
    """
    # Phase 1: Adapter (비동기)
    await _run_adapter_phase()

    # Phase 2: Worker (동기 — xlwings COM은 동기 전용)
    # asyncio 이벤트 루프를 블록하지 않도록 executor에서 실행
    loop = asyncio.get_running_loop()
    await loop.run_in_executor(None, _run_worker_phase)


async def scheduler() -> None:
    """asyncio 무한 루프 스케줄러.

    내결함성 설계:
      - 파이프라인 내 어떤 에러(네트워크, COM, DB 등)가 발생해도
        메인 루프는 절대 죽지 않는다.
      - 에러를 명시적으로 로깅(Rule 3)한 후, 설정된 대기 시간 뒤
        다음 주기를 정상 재시도한다.
    """
    logger.info("OmniStay 스케줄러 기동")
    init_db()

    cycle = 0
    while True:
        cycle += 1
        interval = _get_interval_seconds()
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
        logger.info(
            "━━━ 주기 #%d 시작 [%s] (간격: %d분) ━━━",
            cycle, now, interval // 60,
        )

        try:
            await _execute_pipeline()
            logger.info("━━━ 주기 #%d 정상 완료 ━━━", cycle)

        except SystemExit:
            # ──────────────────────────────────────────────────
            # excel_writer가 COM 에러 시 sys.exit(1)을 호출한다.
            # 스케줄러에서는 이를 치명적 종료가 아닌 "이번 주기 실패"로
            # 재해석하여 다음 주기에서 재시도한다.
            # ──────────────────────────────────────────────────
            logger.error(
                "주기 #%d: 워커가 SystemExit을 발생시켰습니다. "
                "다음 주기에서 재시도합니다.",
                cycle,
            )

        except Exception as exc:
            # ──────────────────────────────────────────────────
            # Fail-Safe 핵심: 모든 예외를 포착하되, 절대로
            # 무시(pass)하지 않는다. 명시적 로깅 후 루프 유지.
            # Rule 3 준수: 에러 내용을 상세히 기록한다.
            # ──────────────────────────────────────────────────
            logger.error(
                "주기 #%d 실패 — %s: %s",
                cycle, type(exc).__name__, exc,
                exc_info=True,
            )

        logger.info("다음 주기까지 %d분 대기...", interval // 60)
        await asyncio.sleep(interval)


def main() -> None:
    """동기 진입점."""
    try:
        asyncio.run(scheduler())
    except KeyboardInterrupt:
        logger.info("스케줄러 수동 종료 (Ctrl+C)")
        sys.exit(0)


if __name__ == "__main__":
    main()
'''


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 빌드 실행
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def build() -> None:
    print(f"[BUILD] omnistay-core 프로젝트 생성 시작: {ROOT}")

    for relative_path, content in FILES.items():
        file_path = ROOT / relative_path
        file_path.parent.mkdir(parents=True, exist_ok=True)
        file_path.write_text(content.lstrip("\n"), encoding="utf-8")
        print(f"  [OK] {relative_path}")

    print(f"\n[BUILD] 완료! 총 {len(FILES)}개 파일 생성됨.")
    print(f"  경로: {ROOT}")
    print(f"\n[NEXT] cd omnistay-core && python main.py")


if __name__ == "__main__":
    build()
