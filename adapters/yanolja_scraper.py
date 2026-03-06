"""
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
    "CO": "EXPECTED_CHECKIN",
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
