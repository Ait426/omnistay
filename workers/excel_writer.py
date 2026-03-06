"""
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
