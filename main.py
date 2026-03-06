"""
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
