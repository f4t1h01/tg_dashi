"""Сведение проходов в строки «сотрудник × день».

Первый приход / последний уход зависят от того, какие точки учитываются.
Для каждого из 4 режимов (только КПП, +общежитие, +техпомещения, все точки)
значения считаются заранее: Excel выбирает нужный режим переключателем,
а здесь те же правила повторены для итогов в боте и кэша формул.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass

from .config import ReportSettings
from .parser import DayRecord, Event, Export

GATE, DORM, TECH = "gate", "dorm", "tech"
# Индекс режима = 1 + общежитие + 2 * техпомещения (так же считается в Excel)
MODES: list[frozenset[str]] = [
    frozenset({GATE}),
    frozenset({GATE, DORM}),
    frozenset({GATE, TECH}),
    frozenset({GATE, DORM, TECH}),
]
DEFAULT_MODE = 1

ON_TIME = "Вовремя"
LATE = "Опоздание"
ABSENT = "Отсутствие"
DAY_OFF = "Выходной"
WEEKEND_WORK = "Работа в выходной"
NO_EXIT_NOTE = "нет отметки ухода"

WEEKDAYS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]


def point_category(point: str) -> str:
    p = point.lower()
    if "dorm" in p or "общеж" in p:
        return DORM
    if "room" in p or "серверн" in p or "щитов" in p:
        return TECH
    return GATE


def minutes(t: dt.time) -> int:
    return t.hour * 60 + t.minute


@dataclass(frozen=True, order=True)
class Employee:
    department: str
    name: str


@dataclass(frozen=True)
class Passes:
    """Первый приход / последний уход для одного режима учёта точек."""

    first_in: dt.time | None = None
    first_in_point: str | None = None
    last_out: dt.time | None = None
    last_out_point: str | None = None
    count: int = 0


@dataclass(frozen=True)
class DayRow:
    date: dt.date
    employee: Employee
    personnel_number: str | None
    workday: bool
    modes: tuple[Passes, ...]  # по одному на каждый режим из MODES


@dataclass(frozen=True)
class DayResult:
    row: DayRow
    passes: Passes
    status: str
    duration_min: int | None
    late_min: int
    early_min: int
    note: str

    @property
    def present(self) -> bool:
        return self.status not in (ABSENT, DAY_OFF)


@dataclass
class Attendance:
    export: Export
    settings: ReportSettings
    dates: list[dt.date]
    employees: list[Employee]
    personnel_numbers: dict[Employee, str | None]
    rows: list[DayRow]  # по дате, затем отдел, затем ФИО

    def is_workday(self, date: dt.date) -> bool:
        return date.isoweekday() in self.settings.workdays

    @property
    def departments(self) -> list[str]:
        return sorted({e.department for e in self.employees})

    @property
    def points(self) -> dict[str, str]:
        return {
            ev.point: point_category(ev.point)
            for rec in self.export.records for ev in rec.events
        }


def _passes(events: list[Event], mode: frozenset[str]) -> Passes:
    included = [e for e in events if point_category(e.point) in mode]
    if not included:
        return Passes()
    # Приход — самый ранний проход в любую сторону: если первая отметка за день —
    # выход, человек уже был на территории.
    first = min(included, key=lambda e: e.time)
    leavings = [e for e in included if e.direction == "out"]
    last = max(leavings, key=lambda e: e.time) if leavings else None
    return Passes(
        first_in=first.time,
        first_in_point=first.point,
        last_out=last.time if last else None,
        last_out_point=last.point if last else None,
        count=len(included),
    )


def build_attendance(export: Export, settings: ReportSettings) -> Attendance:
    by_key: dict[tuple[dt.date, Employee], DayRecord] = {}
    personnel_numbers: dict[Employee, str | None] = {}
    for rec in export.records:
        emp = Employee(rec.department, rec.name)
        by_key[(rec.date, emp)] = rec
        if rec.personnel_number or emp not in personnel_numbers:
            personnel_numbers[emp] = rec.personnel_number or personnel_numbers.get(emp)

    employees = sorted(personnel_numbers)
    dates = [export.begin + dt.timedelta(days=i) for i in range((export.end - export.begin).days + 1)]
    rows: list[DayRow] = []
    for date in dates:
        workday = date.isoweekday() in settings.workdays
        for emp in employees:
            rec = by_key.get((date, emp))
            if rec is None and not workday:
                continue
            events = rec.events if rec else []
            rows.append(DayRow(
                date=date,
                employee=emp,
                personnel_number=personnel_numbers[emp],
                workday=workday,
                modes=tuple(_passes(events, mode) for mode in MODES),
            ))
    return Attendance(export, settings, dates, employees, personnel_numbers, rows)


def evaluate(row: DayRow, settings: ReportSettings, mode: int = DEFAULT_MODE) -> DayResult:
    """Та же логика, что в формулах листа «За день»."""
    p = row.modes[mode - 1]
    late = early = 0
    duration = None
    if p.first_in and row.workday:
        diff = minutes(p.first_in) - minutes(settings.work_start)
        late = diff if diff > settings.late_grace_min else 0
    if p.last_out and row.workday:
        early = max(0, minutes(settings.work_end) - minutes(p.last_out))
    if p.first_in and p.last_out and p.last_out > p.first_in:
        duration = minutes(p.last_out) - minutes(p.first_in)

    if not p.first_in:
        status = ABSENT if row.workday else DAY_OFF
    elif not row.workday:
        status = WEEKEND_WORK
    else:
        status = LATE if late > 0 else ON_TIME
    note = NO_EXIT_NOTE if p.first_in and not p.last_out else ""
    return DayResult(row, p, status, duration, late, early, note)
