"""
models.py — SQLAlchemy ORM 모델 정의

글로벌 데이터 표준 (CLAUDE.md Rule 2):
  - 시간: UTC로만 저장
  - 통화: Decimal 타입 + ISO 4217 통화 코드
  - 상태: ReservationStatus Enum (4가지만 허용)
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

    # 상태 — 4가지 Enum만 허용 (Rule 2)
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
