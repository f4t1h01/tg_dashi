"""Парсер выгрузки СКУД «Passages through perimeter».

Формат: шапка с датами периода, затем таблица Date / Department / Name /
Personnel Number / Events (arrival: время + точка, leaving: время + точка).
Строка с заполненной датой открывает день сотрудника, следующие строки без
даты и ФИО — его остальные проходы за этот же день. «(no)» — события нет.
"""
from __future__ import annotations

import datetime as dt
import io
from dataclasses import dataclass, field

import openpyxl

NO_EVENT = "(no)"

HEADER_ALIASES = {
    "date": ("date", "дата"),
    "department": ("department", "подразделение", "отдел"),
    "name": ("name", "фио", "сотрудник"),
    "personnel_number": ("personnel number", "табельный номер", "таб. номер", "таб. №"),
    "events": ("events", "события"),
}
ARRIVAL_ALIASES = ("arrival", "приход", "вход")
LEAVING_ALIASES = ("leaving", "уход", "выход")
META_ALIASES = {
    "created": ("creation date", "дата создания"),
    "begin": ("begin date", "дата начала"),
    "end": ("end date", "дата окончания"),
}


class ParseError(ValueError):
    pass


@dataclass(frozen=True)
class Event:
    time: dt.time
    direction: str  # "in" | "out"
    point: str


@dataclass
class DayRecord:
    date: dt.date
    department: str
    name: str
    personnel_number: str | None
    events: list[Event] = field(default_factory=list)


@dataclass
class Export:
    title: str
    created: dt.datetime | None
    begin: dt.date
    end: dt.date
    records: list[DayRecord]


def _norm(value) -> str:
    return " ".join(str(value).split()).strip().lower().rstrip(":") if value is not None else ""


def _clean(value) -> str:
    return " ".join(str(value).split()) if value is not None else ""


def _parse_time(value) -> dt.time | None:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, dt.time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, (int, float)) and 0 <= value < 1:
        minutes = round(value * 1440)
        return dt.time(minutes // 60 % 24, minutes % 60)
    text = str(value).strip()
    if not text or text.lower() == NO_EVENT:
        return None
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return dt.datetime.strptime(text, fmt).time().replace(second=0)
        except ValueError:
            continue
    return None


def _parse_date(value) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        for fmt in ("%d.%m.%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M", "%m/%d/%Y %H:%M"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _match(cell, aliases) -> bool:
    return _norm(cell) in aliases


def _find_header(rows: list[tuple]) -> tuple[int, dict[str, int]]:
    for idx, row in enumerate(rows[:50]):
        cols: dict[str, int] = {}
        for col, cell in enumerate(row):
            for key, aliases in HEADER_ALIASES.items():
                if key not in cols and _match(cell, aliases):
                    cols[key] = col
        if {"date", "name"} <= cols.keys():
            return idx, cols
    raise ParseError("Не найдена строка заголовков (Date / Department / Name / Events).")


def _find_meta(rows: list[tuple], header_idx: int) -> dict[str, object]:
    meta: dict[str, object] = {}
    for row in rows[:header_idx]:
        for col, cell in enumerate(row):
            for key, aliases in META_ALIASES.items():
                if _match(cell, aliases):
                    meta[key] = next((v for v in row[col + 1:] if v is not None), None)
    return meta


def _event_columns(rows: list[tuple], header_idx: int, cols: dict[str, int]) -> tuple[int, int]:
    sub = rows[header_idx + 1] if header_idx + 1 < len(rows) else ()
    arrival = next((c for c, v in enumerate(sub) if _match(v, ARRIVAL_ALIASES)), None)
    leaving = next((c for c, v in enumerate(sub) if _match(v, LEAVING_ALIASES)), None)
    if arrival is None or leaving is None:
        if "events" not in cols:
            raise ParseError("Не найдены колонки событий (arrival / leaving).")
        arrival, leaving = cols["events"], cols["events"] + 2
    return arrival, leaving


def parse_workbook(source: bytes | str) -> Export:
    stream = io.BytesIO(source) if isinstance(source, bytes) else source
    try:
        wb = openpyxl.load_workbook(stream, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 — любой битый файл
        raise ParseError("Не удалось открыть файл как Excel (.xlsx).") from exc
    ws = wb.worksheets[0]
    rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    header_idx, cols = _find_header(rows)
    meta = _find_meta(rows, header_idx)
    arr_col, out_col = _event_columns(rows, header_idx, cols)

    def cell(row, col):
        return row[col] if col is not None and col < len(row) else None

    records: dict[tuple, DayRecord] = {}
    current: DayRecord | None = None
    for row in rows[header_idx + 2:]:
        date = _parse_date(cell(row, cols["date"]))
        name = _clean(cell(row, cols["name"]))
        if date and name:
            dept = _clean(cell(row, cols.get("department"))) or "Без отдела"
            pn = _clean(cell(row, cols.get("personnel_number"))) or None
            key = (date, dept, name)
            current = records.get(key)
            if current is None:
                current = records[key] = DayRecord(date, dept, name, pn)
            elif pn and not current.personnel_number:
                current.personnel_number = pn
        if current is None:
            continue
        for col, direction in ((arr_col, "in"), (out_col, "out")):
            time = _parse_time(cell(row, col))
            if time is not None:
                point = _clean(cell(row, col + 1)) or "—"
                current.events.append(Event(time, direction, point))

    if not records:
        raise ParseError("В файле не найдено ни одной записи о проходах.")

    dates = [r.date for r in records.values()]
    begin = _parse_date(meta.get("begin")) or min(dates)
    end = _parse_date(meta.get("end")) or max(dates)
    created = meta.get("created") if isinstance(meta.get("created"), dt.datetime) else None
    title = _clean(rows[0][0]) if rows and rows[0] and rows[0][0] else "Passages through perimeter"
    return Export(title, created, min(begin, min(dates)), max(end, max(dates)), list(records.values()))
