from __future__ import annotations

import datetime as dt
import os
from dataclasses import dataclass, field


def _time(value: str) -> dt.time:
    hours, minutes = value.strip().split(":")
    return dt.time(int(hours), int(minutes))


@dataclass(frozen=True)
class ReportSettings:
    """Значения по умолчанию; в самом Excel их можно поменять на листе «Сводка»."""

    work_start: dt.time = dt.time(9, 0)
    work_end: dt.time = dt.time(18, 0)
    late_grace_min: int = 0
    workdays: frozenset[int] = field(default_factory=lambda: frozenset({1, 2, 3, 4, 5}))  # ISO: 1=Пн

    @classmethod
    def from_env(cls) -> "ReportSettings":
        return cls(
            work_start=_time(os.getenv("WORK_START", "09:00")),
            work_end=_time(os.getenv("WORK_END", "18:00")),
            late_grace_min=int(os.getenv("LATE_GRACE_MIN", "0")),
            workdays=frozenset(int(d) for d in os.getenv("WORKDAYS", "1,2,3,4,5").split(",") if d.strip()),
        )


@dataclass(frozen=True)
class BotConfig:
    token: str
    allowed_users: frozenset[int]
    report: ReportSettings

    @classmethod
    def from_env(cls) -> "BotConfig":
        token = os.getenv("BOT_TOKEN", "").strip()
        if not token:
            raise RuntimeError("BOT_TOKEN не задан")
        allowed = frozenset(int(u) for u in os.getenv("ALLOWED_USERS", "").replace(" ", "").split(",") if u)
        return cls(token=token, allowed_users=allowed, report=ReportSettings.from_env())
