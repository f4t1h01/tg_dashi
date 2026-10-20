"""Сборка Excel-отчёта: «Сводка» (дашборд), «За месяц», «За день».

Всё, что зависит от фильтров на «Сводке», — формулы. Чтобы файл сразу
показывал цифры даже там, где формулы не пересчитываются (превью на телефоне),
каждой формуле передаётся заранее посчитанное значение для фильтров по умолчанию.
"""
from __future__ import annotations

import datetime as dt
import io
import math
from collections.abc import Iterable
from dataclasses import dataclass

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from .analytics import (
    ABSENT, DAY_OFF, DEFAULT_MODE, LATE, ON_TIME, WEEKDAYS_RU, WEEKEND_WORK,
    Attendance, DayResult, evaluate, minutes,
)

SUMMARY, MONTH, DAY, CALC = "Сводка", "За месяц", "За день", "_calc"
ALL = "Все"
YES, NO = "Учитывать", "Не учитывать"
METRICS = ["Присутствовали", "Опоздали", "Отсутствовали", "Ранние уходы", "Посещаемость, %", "Ср. часы на территории"]
RANKS = ["Опоздания (раз)", "Опоздания (минуты)", "Отсутствия", "Ранние уходы", "Часы на территории"]
NA = "#N/A"

C_DARK, C_HEAD, C_INPUT = "#1F3864", "#2F5597", "#FFF2CC"
STATUS_COLORS = {
    ON_TIME: ("#C6EFCE", "#006100"),
    LATE: ("#FFEB9C", "#9C5700"),
    ABSENT: ("#FFC7CE", "#9C0006"),
    DAY_OFF: ("#EDEDED", "#595959"),
    WEEKEND_WORK: ("#DDEBF7", "#1F4E78"),
}


def q(sheet: str) -> str:
    return f"'{sheet}'"


def frac(t: dt.time | None) -> float | str:
    return minutes(t) / 1440 if t else ""


def hhmm(t: dt.time | None) -> str:
    return t.strftime("%H:%M") if t else ""


def _avg(values: list[float]) -> float | None:
    return sum(values) / len(values) if values else None


def xround(value: float, digits: int = 0) -> float:
    """Округление как в Excel ROUND (половина — вверх), а не банковское round()."""
    k = 10 ** digits
    return math.floor(value * k + 0.5 + 1e-9) / k


def _safe_pct(num: int, den: int) -> float | None:
    return xround(100 * num / den, 1) if den else None


@dataclass
class Totals:
    on_time: int
    late: int
    absent: int
    weekend: int
    early: int
    late_min: int
    hours: float
    avg_in: float | None   # доля суток
    avg_out: float | None
    avg_dur: float | None

    @property
    def present(self) -> int:
        return self.on_time + self.late

    @property
    def attendance_pct(self) -> float | None:
        return _safe_pct(self.present, self.present + self.absent)

    @property
    def late_pct(self) -> float | None:
        return _safe_pct(self.late, self.present)


def totals(results: Iterable[DayResult]) -> Totals:
    res = list(results)
    work = [r for r in res if r.row.workday]
    return Totals(
        on_time=sum(r.status == ON_TIME for r in res),
        late=sum(r.status == LATE for r in res),
        absent=sum(r.status == ABSENT for r in res),
        weekend=sum(r.status == WEEKEND_WORK for r in res),
        early=sum(r.early_min > 0 for r in res),
        late_min=sum(r.late_min for r in res),
        hours=xround(sum(r.duration_min or 0 for r in res) / 60, 1),
        avg_in=_avg([minutes(r.passes.first_in) / 1440 for r in work if r.passes.first_in]),
        avg_out=_avg([minutes(r.passes.last_out) / 1440 for r in work if r.passes.last_out]),
        avg_dur=_avg([r.duration_min / 1440 for r in work if r.duration_min]),
    )


class Report:
    def __init__(self, att: Attendance):
        self.att = att
        self.s = att.settings
        self.results = [evaluate(row, self.s, DEFAULT_MODE) for row in att.rows]
        self.buf = io.BytesIO()
        self.wb = xlsxwriter.Workbook(self.buf, {"in_memory": True})
        self.wb.set_calc_mode("auto")
        self.wb.set_properties({"title": "Посещаемость", "author": "tg_dashi"})
        self._formats()
        self.ws_summary = self.wb.add_worksheet(SUMMARY)
        self.ws_month = self.wb.add_worksheet(MONTH)
        self.ws_day = self.wb.add_worksheet(DAY)
        self.ws_calc = self.wb.add_worksheet(CALC)
        self.ws_calc.hide()

        n = len(att.rows)
        self.day_first, self.day_last = 5, 4 + n  # строки Excel (1-based) данных на «За день»
        self.day_row_of = {(r.date, r.employee): self.day_first + i for i, r in enumerate(att.rows)}
        self.names = sorted({e.name for e in att.employees})

    # ---------- общее ----------
    def _formats(self) -> None:
        f = self.wb.add_format
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        self.f = {
            "title": f({**base, "bold": True, "font_size": 18, "font_color": C_DARK}),
            "subtitle": f({**base, "font_color": "#595959", "italic": True}),
            "section": f({**base, "bold": True, "font_size": 12, "font_color": C_DARK, "bottom": 2, "bottom_color": C_HEAD}),
            "label": f({**base, "bold": True, "font_color": "#404040", "align": "right"}),
            "input": f({**base, "bg_color": C_INPUT, "border": 1, "border_color": "#BF9000", "bold": True}),
            "input_date": f({**base, "bg_color": C_INPUT, "border": 1, "border_color": "#BF9000", "bold": True, "num_format": "dd.mm.yyyy", "align": "left"}),
            "input_time": f({**base, "bg_color": C_INPUT, "border": 1, "border_color": "#BF9000", "bold": True, "num_format": "hh:mm", "align": "left"}),
            "input_int": f({**base, "bg_color": C_INPUT, "border": 1, "border_color": "#BF9000", "bold": True, "align": "left"}),
            "note": f({**base, "font_color": "#7F7F7F", "italic": True}),
            "hint": f({**base, "font_color": "#7F7F7F", "italic": True, "text_wrap": True}),
            "kpi_label": f({**base, "font_size": 9, "font_color": "#FFFFFF", "bg_color": C_HEAD, "align": "center", "text_wrap": True}),
            "kpi": f({**base, "font_size": 18, "bold": True, "font_color": C_DARK, "bg_color": "#EEF3FA", "align": "center", "bottom": 2, "bottom_color": C_HEAD}),
            "kpi_pct": f({**base, "font_size": 18, "bold": True, "font_color": C_DARK, "bg_color": "#EEF3FA", "align": "center", "num_format": '0.0"%"', "bottom": 2, "bottom_color": C_HEAD}),
            "kpi_time": f({**base, "font_size": 18, "bold": True, "font_color": C_DARK, "bg_color": "#EEF3FA", "align": "center", "num_format": "hh:mm", "bottom": 2, "bottom_color": C_HEAD}),
            "kpi_dur": f({**base, "font_size": 18, "bold": True, "font_color": C_DARK, "bg_color": "#EEF3FA", "align": "center", "num_format": "[h]:mm", "bottom": 2, "bottom_color": C_HEAD}),
            "th": f({**base, "bold": True, "font_color": "#FFFFFF", "bg_color": C_HEAD, "align": "center", "text_wrap": True, "border": 1, "border_color": "#BFBFBF"}),
            "th_off": f({**base, "bold": True, "font_color": "#FFFFFF", "bg_color": "#8497B0", "align": "center", "text_wrap": True, "border": 1, "border_color": "#BFBFBF"}),
            "td": f({**base, "border": 1, "border_color": "#D9D9D9"}),
            "td_c": f({**base, "border": 1, "border_color": "#D9D9D9", "align": "center"}),
            "td_date": f({**base, "border": 1, "border_color": "#D9D9D9", "num_format": "dd.mm.yyyy", "align": "center"}),
            "td_short": f({**base, "border": 1, "border_color": "#D9D9D9", "num_format": "dd.mm", "align": "center"}),
            "td_time": f({**base, "border": 1, "border_color": "#D9D9D9", "num_format": "hh:mm", "align": "center"}),
            "td_dur": f({**base, "border": 1, "border_color": "#D9D9D9", "num_format": "[h]:mm", "align": "center"}),
            "td_num": f({**base, "border": 1, "border_color": "#D9D9D9", "num_format": "0.0", "align": "center"}),
            "td_int": f({**base, "border": 1, "border_color": "#D9D9D9", "num_format": "0", "align": "center"}),
            "td_muted": f({**base, "border": 1, "border_color": "#D9D9D9", "font_color": "#A6A6A6", "num_format": "hh:mm", "align": "center"}),
            "td_muted_int": f({**base, "font_color": "#A6A6A6", "align": "center"}),
            "cell": f({**base, "font_size": 9, "border": 1, "border_color": "#D9D9D9", "align": "center", "text_wrap": True}),
            "cell_off": f({**base, "font_size": 9, "border": 1, "border_color": "#D9D9D9", "bg_color": "#EDEDED"}),
            "foot": f({**base, "bold": True, "border": 1, "border_color": "#D9D9D9", "bg_color": "#F2F2F2", "align": "center"}),
            "foot_label": f({**base, "bold": True, "border": 1, "border_color": "#D9D9D9", "bg_color": "#F2F2F2", "align": "right"}),
            "calc_time": f({"num_format": "hh:mm"}),
            "calc_date": f({"num_format": "dd.mm.yyyy"}),
        }
        self.status_fmt = {
            st: f({"bg_color": bg, "font_color": fg}) for st, (bg, fg) in STATUS_COLORS.items()
        }

    def d(self, letter: str) -> str:
        """Диапазон колонки на листе «За день»."""
        return f"{q(DAY)}!${letter}${self.day_first}:${letter}${self.day_last}"

    def filt(self, dept: str = "pDept", name: str = "pName", date: str | None = None) -> str:
        """Критерии COUNTIFS: отдел, сотрудник, период (или конкретная дата)."""
        parts = [self.d("C"), dept, self.d("D"), name]
        if date:
            parts += [self.d("A"), date]
        else:
            parts += [self.d("A"), '">="&pFrom', self.d("A"), '"<="&pTo']
        return ",".join(parts)

    def cnt(self, crit: str, status: str) -> str:
        return f'COUNTIFS({crit},{self.d("F")},"{status}")'

    def build(self) -> bytes:
        self._calc_sheet()
        self._day_sheet()
        self._month_sheet()
        self._summary_sheet()
        self.ws_summary.activate()
        self.wb.close()
        return self.buf.getvalue()

    # ---------- _calc ----------
    def _calc_sheet(self) -> None:
        ws, att = self.ws_calc, self.att
        ws.write_column("A1", ["deptCrit", "nameCrit", "mode", "offset"])
        ws.write_formula("B1", f'=IF(fDept="{ALL}","*",fDept)', None, "*")
        ws.write_formula("B2", f'=IF(fName="{ALL}","*",fName)', None, "*")
        ws.write_formula("B3", f'=1+(fDorm="{YES}")+2*(fTech="{YES}")', None, DEFAULT_MODE)
        ws.write_formula("B4", "=(pMode-1)*4", None, (DEFAULT_MODE - 1) * 4)
        for name, ref in {"pDept": "$B$1", "pName": "$B$2", "pMode": "$B$3", "pOff": "$B$4"}.items():
            self.wb.define_name(name, f"={q(CALC)}!{ref}")

        ws.write_column("D1", ["Отдел", ALL, *att.departments])
        ws.write_column("E1", ["Сотрудник", ALL, *self.names])
        ws.write("F1", "Дата")
        ws.write("G1", "Рабочий")
        for i, date in enumerate(att.dates, start=1):
            ws.write_datetime(i, 5, date, self.f["calc_date"])
            ws.write_number(i, 6, int(att.is_workday(date)))
        ws.write_column("H1", ["Показатель", *METRICS])
        ws.write_column("I1", ["Рейтинг", *RANKS])
        ws.write_column("J1", ["Точки", NO, YES])

        self.lists = {
            "dept": f"={q(CALC)}!$D$2:$D${len(att.departments) + 2}",
            "name": f"={q(CALC)}!$E$2:$E${len(self.names) + 2}",
            "date": f"={q(CALC)}!$F$2:$F${len(att.dates) + 1}",
            "metric": f"={q(CALC)}!$H$2:$H${len(METRICS) + 1}",
            "rank": f"={q(CALC)}!$I$2:$I${len(RANKS) + 1}",
            "points": f"={q(CALC)}!$J$2:$J$3",
        }
        self.calc_dates = f"{q(CALC)}!$F$2:$F${len(att.dates) + 1}"
        self.calc_wd = f"{q(CALC)}!$G$2:$G${len(att.dates) + 1}"

        # Сводные показатели по выбранным фильтрам (их используют плитки KPI)
        t = totals(self.results)
        crit = self.filt()
        wd = f'{self.d("P")},1'
        avg = lambda col: (  # noqa: E731
            f'=IFERROR(SUMIFS({self.d(col)},{crit},{self.d(col)},">0",{wd})'
            f'/COUNTIFS({crit},{self.d(col)},">0",{wd}),"—")'
        )
        kpis = [
            ("onTime", f"={self.cnt(crit, ON_TIME)}", t.on_time),
            ("late", f"={self.cnt(crit, LATE)}", t.late),
            ("absent", f"={self.cnt(crit, ABSENT)}", t.absent),
            ("weekend", f"={self.cnt(crit, WEEKEND_WORK)}", t.weekend),
            ("early", f'=COUNTIFS({crit},{self.d("M")},">0")', t.early),
            ("avgIn", avg("G"), t.avg_in if t.avg_in is not None else "—"),
            ("avgOut", avg("I"), t.avg_out if t.avg_out is not None else "—"),
            ("avgDur", avg("K"), t.avg_dur if t.avg_dur is not None else "—"),
            ("workdays", f'=COUNTIFS({self.calc_dates},">="&pFrom,{self.calc_dates},"<="&pTo,{self.calc_wd},1)',
             sum(att.is_workday(x) for x in att.dates)),
            ("employees", f"=SUM({q(CALC)}!$N$2:$N${len(att.employees) + 1})", len(att.employees)),
        ]
        self.kpi_ref = {}
        for i, (key, formula, value) in enumerate(kpis, start=6):
            ws.write(f"A{i}", key)
            ws.write_formula(f"B{i}", formula, None, value)
            self.kpi_ref[key] = (f"{q(CALC)}!$B${i}", value)
        self.totals = t

        # Помощник по сотрудникам: показатели за период и ключ для рейтинга
        ws.write_row("L1", ["Отдел", "Сотрудник", "В фильтре", "Присутствий", "Опозданий", "Отсутствий",
                            "Ранних уходов", "Минут опозданий", "Часов", "Значение", "Ключ"])
        by_emp: dict = {}
        for r in self.results:
            by_emp.setdefault(r.row.employee, []).append(r)
        self.emp_values = []
        for i, emp in enumerate(att.employees, start=2):
            fe = self.filt(dept=f"$L{i}", name=f"$M{i}")
            et = totals(by_emp.get(emp, []))
            ws.write(f"L{i}", emp.department)
            ws.write(f"M{i}", emp.name)
            ws.write_formula(f"N{i}", f'=IF(AND(OR(fDept="{ALL}",$L{i}=fDept),OR(fName="{ALL}",$M{i}=fName)),1,0)', None, 1)
            ws.write_formula(f"O{i}", f"={self.cnt(fe, ON_TIME)}+{self.cnt(fe, LATE)}", None, et.present)
            ws.write_formula(f"P{i}", f"={self.cnt(fe, LATE)}", None, et.late)
            ws.write_formula(f"Q{i}", f"={self.cnt(fe, ABSENT)}", None, et.absent)
            ws.write_formula(f"R{i}", f'=COUNTIFS({fe},{self.d("M")},">0")', None, et.early)
            ws.write_formula(f"S{i}", f'=SUMIFS({self.d("L")},{fe})', None, et.late_min)
            ws.write_formula(f"T{i}", f'=ROUND(SUMIFS({self.d("K")},{fe})*24,1)', None, et.hours)
            ws.write_formula(f"U{i}", f"=CHOOSE(MATCH(fRank,{q(CALC)}!$I$2:$I${len(RANKS) + 1},0),P{i},S{i},Q{i},R{i},T{i})",
                             None, et.late)
            ws.write_formula(f"V{i}", f"=IF(N{i}=1,U{i}-ROW()/1000000,-1)", None, et.late - i / 1_000_000)
            self.emp_values.append((et.late - i / 1_000_000, i, emp, et.late))
        self.emp_last = len(att.employees) + 1

    # ---------- За день ----------
    def _day_sheet(self) -> None:
        ws, att, f = self.ws_day, self.att, self.f
        ws.write("A1", "Посещаемость по дням", f["title"])
        ws.write_formula("A2", self._mode_caption_formula(), f["subtitle"], self._mode_caption())
        ws.write("A3", "Фильтруйте по дате, отделу, сотруднику и статусу стрелками в заголовках.", f["note"])

        headers = ["Дата", "День", "Отдел", "Сотрудник", "Таб. №", "Статус", "Приход", "Точка прихода",
                   "Уход", "Точка ухода", "На территории", "Опоздание, мин", "Ранний уход, мин", "Проходов",
                   "Примечание"]
        widths = [11, 6, 20, 40, 10, 17, 8, 20, 8, 20, 12, 11, 11, 9, 18]
        for c, w in enumerate(widths):
            ws.set_column(c, c, w)
        # P — признак рабочего дня, Q:AF — приход/уход по 4 режимам, AG:AJ — число проходов по режимам
        ws.set_column("P:AJ", 10, None, {"hidden": True})
        hidden = ["Рабочий"] + [f"{k}{m}" for m in range(1, 5) for k in ("Приход", "ТочкаП", "Уход", "ТочкаУ")] \
            + [f"Проходов{m}" for m in range(1, 5)]
        ws.write_row(3, 15, hidden)

        n = len(att.rows)
        ws.add_table(3, 0, 3 + n, len(headers) - 1, {
            "columns": [{"header": h} for h in headers],
            "style": "Table Style Medium 2",
            "autofilter": True,
        })
        ws.set_row(3, 30)
        for i, res in enumerate(self.results):
            r0, x = 4 + i, 5 + i
            row = res.row
            ws.write_datetime(r0, 0, row.date, f["td_date"])
            ws.write_string(r0, 1, WEEKDAYS_RU[row.date.weekday()], f["td_c"])
            ws.write_string(r0, 2, row.employee.department, f["td"])
            ws.write_string(r0, 3, row.employee.name, f["td"])
            ws.write(r0, 4, row.personnel_number or "", f["td_c"])
            ws.write_number(r0, 15, int(row.workday))
            for m, p in enumerate(row.modes):
                base = 16 + m * 4
                if p.first_in:
                    ws.write_number(r0, base, minutes(p.first_in) / 1440, f["calc_time"])
                    ws.write_string(r0, base + 1, p.first_in_point or "")
                if p.last_out:
                    ws.write_number(r0, base + 2, minutes(p.last_out) / 1440, f["calc_time"])
                    ws.write_string(r0, base + 3, p.last_out_point or "")
                ws.write_number(r0, 32 + m, p.count)

            pick = lambda k: f"INDEX($Q{x}:$AF{x},pOff+{k})"  # noqa: E731
            p = res.passes
            ws.write_formula(r0, 6, f'=IF({pick(1)}=0,"",{pick(1)})', f["td_time"], frac(p.first_in))
            ws.write_formula(r0, 7, f'=IF({pick(2)}=0,"",{pick(2)})', f["td"], p.first_in_point or "")
            ws.write_formula(r0, 8, f'=IF({pick(3)}=0,"",{pick(3)})', f["td_time"], frac(p.last_out))
            ws.write_formula(r0, 9, f'=IF({pick(4)}=0,"",{pick(4)})', f["td"], p.last_out_point or "")
            ws.write_formula(r0, 10, f'=IF(AND(ISNUMBER($G{x}),ISNUMBER($I{x})),IF($I{x}>$G{x},$I{x}-$G{x},""),"")',
                             f["td_dur"], res.duration_min / 1440 if res.duration_min else "")
            late_expr = f"ROUND(($G{x}-pStart)*1440,0)"
            ws.write_formula(r0, 11, f"=IF(AND($P{x}=1,ISNUMBER($G{x})),IF({late_expr}>pGrace,{late_expr},0),0)",
                             f["td_int"], res.late_min)
            ws.write_formula(r0, 12, f"=IF(AND($P{x}=1,ISNUMBER($I{x})),MAX(0,ROUND((pEnd-$I{x})*1440,0)),0)",
                             f["td_int"], res.early_min)
            ws.write_formula(r0, 13, f"=INDEX($AG{x}:$AJ{x},pMode)", f["td_int"], p.count)
            ws.write_formula(r0, 5,
                             f'=IF(NOT(ISNUMBER($G{x})),IF($P{x}=1,"{ABSENT}","{DAY_OFF}"),'
                             f'IF($P{x}=0,"{WEEKEND_WORK}",IF($L{x}>0,"{LATE}","{ON_TIME}")))',
                             f["td_c"], res.status)
            ws.write_formula(r0, 14, f'=IF(AND(ISNUMBER($G{x}),NOT(ISNUMBER($I{x}))),"нет отметки ухода","")',
                             f["td"], res.note)

        last = 3 + n
        for st, fmt in self.status_fmt.items():
            ws.conditional_format(4, 5, last, 5, {"type": "cell", "criteria": "==", "value": f'"{st}"', "format": fmt})
        ws.conditional_format(4, 11, last, 12, {"type": "cell", "criteria": ">", "value": 0,
                                                 "format": self.status_fmt[LATE]})
        ws.freeze_panes(4, 4)

    # ---------- За месяц ----------
    def _month_sheet(self) -> None:
        ws, att, f = self.ws_month, self.att, self.f
        dates = att.dates
        n_emp = len(att.employees)
        first_date_col = 4
        tot_col = first_date_col + len(dates)
        tot_headers = ["Присутствий", "Опозданий", "Отсутствий", "Ранних уходов", "Работа в вых.",
                       "Часов", "Ср. приход", "Ср. уход"]
        grid_col = tot_col + len(tot_headers) + 1  # скрытая сетка статусов для подсветки
        hr = 4  # строка заголовка (0-based)

        period = f"{att.export.begin:%d.%m.%Y} – {att.export.end:%d.%m.%Y}"
        ws.write("A1", f"Посещаемость за период {period}", f["title"])
        ws.write_formula("A2", self._mode_caption_formula(), f["subtitle"], self._mode_caption())
        legend = [("Вовремя", ON_TIME), ("Опоздание", LATE), ("Н — отсутствие", ABSENT),
                  ("Работа в выходной", WEEKEND_WORK)]
        ws.write("A3", "В ячейке: приход / уход.", f["note"])
        for i, (label, st) in enumerate(legend):
            bg, fg = STATUS_COLORS[st]
            ws.write(2, 3 + i * 3, label, self.wb.add_format({"bg_color": bg, "font_color": fg, "font_size": 9,
                                                             "align": "center", "border": 1, "border_color": "#D9D9D9"}))

        ws.set_column(0, 0, 4)
        ws.set_column(1, 1, 18)
        ws.set_column(2, 2, 38)
        ws.set_column(3, 3, 8)
        ws.set_column(first_date_col, tot_col - 1, 6.3)
        ws.set_column(tot_col, tot_col + len(tot_headers) - 1, 10)
        ws.set_column(grid_col, grid_col + len(dates) - 1, 4, None, {"hidden": True})

        ws.set_row(hr, 32)
        for c, h in enumerate(["№", "Отдел", "Сотрудник", "Таб. №"]):
            ws.write(hr, c, h, f["th"])
        for j, date in enumerate(dates):
            fmt = f["th"] if att.is_workday(date) else f["th_off"]
            ws.write(hr, first_date_col + j, f"{date:%d.%m}\n{WEEKDAYS_RU[date.weekday()]}", fmt)
            ws.write(hr, grid_col + j, f"{date:%d.%m}")
        for k, h in enumerate(tot_headers):
            ws.write(hr, tot_col + k, h, f["th"])

        by_key = {(r.row.date, r.row.employee): r for r in self.results}
        by_emp: dict = {}
        for r in self.results:
            by_emp.setdefault(r.row.employee, []).append(r)

        def t(col: str, x: int) -> str:
            return f"RIGHT(\"0\"&HOUR({q(DAY)}!${col}${x}),2)&\":\"&RIGHT(\"0\"&MINUTE({q(DAY)}!${col}${x}),2)"

        for i, emp in enumerate(att.employees):
            r0 = hr + 1 + i
            ws.set_row(r0, 26)
            ws.write_number(r0, 0, i + 1, f["td_c"])
            ws.write_string(r0, 1, emp.department, f["td"])
            ws.write_string(r0, 2, emp.name, f["td"])
            ws.write(r0, 3, att.personnel_numbers[emp] or "", f["td_c"])
            for j, date in enumerate(dates):
                res = by_key.get((date, emp))
                if res is None:
                    ws.write_blank(r0, first_date_col + j, None, f["cell_off"])
                    continue
                x = self.day_row_of[(date, emp)]
                st = f"{q(DAY)}!$F${x}"
                if res.status == ABSENT:
                    cached = "Н"
                elif res.passes.first_in:
                    cached = f"{hhmm(res.passes.first_in)}\n{hhmm(res.passes.last_out) or '—'}"
                else:
                    cached = ""
                ws.write_formula(
                    r0, first_date_col + j,
                    f'=IF({st}="{ABSENT}","Н",IF(ISNUMBER({q(DAY)}!$G${x}),{t("G", x)}&CHAR(10)&'
                    f'IF(ISNUMBER({q(DAY)}!$I${x}),{t("I", x)},"—"),""))',
                    f["cell"], cached)
                ws.write_formula(r0, grid_col + j, f"={st}", None, res.status)

            fe = f'{self.d("C")},$B{r0 + 1},{self.d("D")},$C{r0 + 1}'
            et = totals(by_emp.get(emp, []))
            wd = f'{self.d("P")},1'
            avg = lambda col: (  # noqa: E731
                f'=IFERROR(SUMIFS({self.d(col)},{fe},{self.d(col)},">0",{wd})/COUNTIFS({fe},{self.d(col)},">0",{wd}),"")'
            )
            cells = [
                (f"={self.cnt(fe, ON_TIME)}+{self.cnt(fe, LATE)}", f["td_int"], et.present),
                (f"={self.cnt(fe, LATE)}", f["td_int"], et.late),
                (f"={self.cnt(fe, ABSENT)}", f["td_int"], et.absent),
                (f'=COUNTIFS({fe},{self.d("M")},">0")', f["td_int"], et.early),
                (f"={self.cnt(fe, WEEKEND_WORK)}", f["td_int"], et.weekend),
                (f'=ROUND(SUMIFS({self.d("K")},{fe})*24,1)', f["td_num"], et.hours),
                (avg("G"), f["td_time"], et.avg_in if et.avg_in is not None else ""),
                (avg("I"), f["td_time"], et.avg_out if et.avg_out is not None else ""),
            ]
            for k, (formula, fmt, value) in enumerate(cells):
                ws.write_formula(r0, tot_col + k, formula, fmt, value)

        first_r0, last_r0 = hr + 1, hr + n_emp
        grid_first = xl_col_to_name(grid_col)
        for st, fmt in self.status_fmt.items():
            ws.conditional_format(first_r0, first_date_col, last_r0, tot_col - 1, {
                "type": "formula", "criteria": f'={grid_first}{first_r0 + 1}="{st}"', "format": fmt})

        # Итоги по дням внизу
        foot = last_r0 + 2
        for k, (label, statuses) in enumerate([
            ("Присутствовали", (ON_TIME, LATE, WEEKEND_WORK)),
            ("Опоздали", (LATE,)),
            ("Отсутствовали", (ABSENT,)),
        ]):
            r0 = foot + k
            ws.merge_range(r0, 0, r0, 3, label, f["foot_label"])
            for j, date in enumerate(dates):
                g = xl_col_to_name(grid_col + j)
                rng = f"${g}${first_r0 + 1}:${g}${last_r0 + 1}"
                formula = "=" + "+".join(f'COUNTIF({rng},"{st}")' for st in statuses)
                value = sum(1 for r in self.results if r.row.date == date and r.status in statuses)
                ws.write_formula(r0, first_date_col + j, formula, f["foot"], value)

        ws.autofilter(hr, 0, last_r0, tot_col + len(tot_headers) - 1)
        ws.freeze_panes(hr + 1, first_date_col)

    # ---------- Сводка ----------
    def _summary_sheet(self) -> None:
        ws, att, f, t = self.ws_summary, self.att, self.f, self.totals
        ws.hide_gridlines(2)
        ws.set_column("A:A", 2)
        ws.set_column("B:K", 13)
        ws.set_column("L:L", 3)
        ws.set_column("M:M", 34)
        ws.set_column("N:T", 13)

        exp = att.export
        ws.write("B1", "Сводка посещаемости", f["title"])
        created = f" · выгрузка от {exp.created:%d.%m.%Y %H:%M}" if exp.created else ""
        ws.write("B2", f"Период {exp.begin:%d.%m.%Y} – {exp.end:%d.%m.%Y}{created} · "
                       f"сотрудников: {len(att.employees)} · отделов: {len(att.departments)}", f["subtitle"])

        # --- фильтры ---
        def field(row: int, col: int, label: str, value, fmt: str, name: str, validation: dict | None) -> None:
            ws.merge_range(row, col, row, col + 1, label, f["label"])
            ws.merge_range(row, col + 2, row, col + 4, "", f[fmt])
            if isinstance(value, dt.date):
                ws.write_datetime(row, col + 2, value, f[fmt])
            elif isinstance(value, dt.time):
                ws.write_number(row, col + 2, minutes(value) / 1440, f[fmt])
            else:
                ws.write(row, col + 2, value, f[fmt])
            if validation:
                ws.data_validation(row, col + 2, row, col + 2, validation)
            self.wb.define_name(name, f"={q(SUMMARY)}!{xl_rowcol_to_cell(row, col + 2, True, True)}")

        lst = lambda key: {"validate": "list", "source": self.lists[key]}  # noqa: E731
        ws.write("B4", "Фильтры", f["section"])
        ws.write("G4", "Настройки расчёта", f["section"])
        field(4, 1, "Отдел", ALL, "input", "fDept", lst("dept"))
        field(5, 1, "Сотрудник", ALL, "input", "fName", lst("name"))
        field(6, 1, "Дата с", att.dates[0], "input_date", "pFrom", lst("date"))
        field(7, 1, "Дата по", att.dates[-1], "input_date", "pTo", lst("date"))
        field(8, 1, "График по дням", METRICS[0], "input", "fMetric", lst("metric"))
        field(9, 1, "Рейтинг сотрудников", RANKS[0], "input", "fRank", lst("rank"))
        field(4, 6, "Общежитие (Dorm)", NO, "input", "fDorm", lst("points"))
        field(5, 6, "Тех. помещения", NO, "input", "fTech", lst("points"))
        time_v = {"validate": "time", "criteria": "between", "minimum": dt.time(0, 0), "maximum": dt.time(23, 59)}
        field(6, 6, "Начало дня", self.s.work_start, "input_time", "pStart", time_v)
        field(7, 6, "Конец дня", self.s.work_end, "input_time", "pEnd", time_v)
        field(8, 6, "Допуск опоздания, мин", self.s.late_grace_min, "input_int", "pGrace",
              {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 240})
        for r in range(4, 9):
            ws.set_row(r, 20)
        ws.merge_range(9, 6, 9, 10, "Жёлтые ячейки можно менять — пересчитываются все листы. "
                                    "КПП и турникеты учитываются всегда.", f["hint"])
        ws.set_row(9, 28)

        # --- KPI ---
        k = self.kpi_ref
        tiles = [
            ("Сотрудников", f"={k['employees'][0]}", "kpi", len(att.employees)),
            ("Рабочих дней", f"={k['workdays'][0]}", "kpi", k["workdays"][1]),
            ("Посещаемость", f'=IFERROR(100*({k["onTime"][0]}+{k["late"][0]})/({k["onTime"][0]}+{k["late"][0]}+{k["absent"][0]}),"—")',
             "kpi_pct", t.attendance_pct if t.attendance_pct is not None else "—"),
            ("Опозданий", f"={k['late'][0]}", "kpi", t.late),
            ("Доля опозданий", f'=IFERROR(100*{k["late"][0]}/({k["onTime"][0]}+{k["late"][0]}),"—")',
             "kpi_pct", t.late_pct if t.late_pct is not None else "—"),
            ("Отсутствий", f"={k['absent'][0]}", "kpi", t.absent),
            ("Ранних уходов", f"={k['early'][0]}", "kpi", t.early),
            ("Ср. время прихода", f"={k['avgIn'][0]}", "kpi_time", k["avgIn"][1]),
            ("Ср. время ухода", f"={k['avgOut'][0]}", "kpi_time", k["avgOut"][1]),
            ("Ср. на территории", f"={k['avgDur'][0]}", "kpi_dur", k["avgDur"][1]),
        ]
        ws.set_row(11, 26)
        ws.set_row(12, 36)
        for i, (label, formula, fmt, value) in enumerate(tiles):
            ws.write(11, 1 + i, label, f["kpi_label"])
            ws.write_formula(12, 1 + i, formula, f[fmt], value)

        # --- данные для графиков (справа) ---
        row = self._daily_block(3)
        row = self._dept_block(row + 2)
        row = self._buckets_block(row + 2)
        self._top_block(row + 2)
        self._charts()

    def _daily_block(self, r: int) -> int:
        ws, f, att = self.ws_summary, self.f, self.att
        ws.write(r, 12, "Данные по дням", f["section"])
        heads = ["Дата", "Присутствовали", "Опоздали", "Отсутствовали", "Ранние уходы", "Посещаемость, %",
                 "Ср. часы на терр."]
        for c, h in enumerate(heads):
            ws.write(r + 1, 12 + c, h, f["th"])
        ws.write_formula(r + 1, 19, "=fMetric", f["th"], METRICS[0])
        self.metric_title = (r + 1, 19)
        self.daily_rows = (r + 2, r + 1 + len(att.dates))
        ws.set_row(r + 1, 30)
        for j, date in enumerate(att.dates):
            r0 = r + 2 + j
            x = r0 + 1
            crit = self.filt(date=f"$M{x}")
            res = [s for s in self.results if s.row.date == date]
            tt = totals(res)
            inr = f"AND($M{x}>=pFrom,$M{x}<=pTo)"
            on, late, ab = (self.cnt(crit, s) for s in (ON_TIME, LATE, ABSENT))
            present = tt.present + tt.weekend
            pct = tt.attendance_pct if tt.attendance_pct is not None else NA
            hours = xround(tt.avg_dur * 24, 2) if tt.avg_dur else NA
            ws.write_datetime(r0, 12, date, f["td_date"])
            cells = [
                (f"=IF({inr},{on}+{late}+{self.cnt(crit, WEEKEND_WORK)},NA())", f["td_int"], present),
                (f"=IF({inr},{late},NA())", f["td_int"], tt.late),
                (f"=IF({inr},{ab},NA())", f["td_int"], tt.absent),
                (f'=IF({inr},COUNTIFS({crit},{self.d("M")},">0"),NA())', f["td_int"], tt.early),
                (f"=IF({inr},IFERROR(ROUND(100*({on}+{late})/({on}+{late}+{ab}),1),NA()),NA())", f["td_num"], pct),
                (f'=IF({inr},IFERROR(ROUND(SUMIFS({self.d("K")},{crit},{self.d("K")},">0",{self.d("P")},1)'
                 f'/COUNTIFS({crit},{self.d("K")},">0",{self.d("P")},1)*24,2),NA()),NA())', f["td_num"], hours),
            ]
            for c, (formula, fmt, value) in enumerate(cells, start=13):
                ws.write_formula(r0, c, formula, fmt, value)
            ws.write_formula(r0, 19, f"=CHOOSE(MATCH(fMetric,{self.lists['metric'][1:]},0),N{x},O{x},P{x},Q{x},R{x},S{x})",
                             f["td_c"], present)
        return r + 1 + len(att.dates)

    def _dept_block(self, r: int) -> int:
        ws, f, att = self.ws_summary, self.f, self.att
        ws.write(r, 12, "По отделам (без фильтра по отделу)", f["section"])
        heads = ["Отдел", "Сотрудников", "Присутствий", "Опозданий", "Отсутствий", "Посещаемость, %", "Доля опозданий, %"]
        for c, h in enumerate(heads):
            ws.write(r + 1, 12 + c, h, f["th"])
        ws.set_row(r + 1, 30)
        self.dept_rows = (r + 2, r + 1 + len(att.departments))
        for j, dept in enumerate(att.departments):
            r0 = r + 2 + j
            x = r0 + 1
            crit = self.filt(dept=f"$M{x}")
            tt = totals(s for s in self.results if s.row.employee.department == dept)
            emp_n = sum(e.department == dept for e in att.employees)
            ws.write_string(r0, 12, dept, f["td"])
            cells = [
                (f"=COUNTIFS({q(CALC)}!$L$2:$L${self.emp_last},$M{x},{q(CALC)}!$M$2:$M${self.emp_last},pName)",
                 f["td_int"], emp_n),
                (f"={self.cnt(crit, ON_TIME)}+{self.cnt(crit, LATE)}", f["td_int"], tt.present),
                (f"={self.cnt(crit, LATE)}", f["td_int"], tt.late),
                (f"={self.cnt(crit, ABSENT)}", f["td_int"], tt.absent),
                (f"=IFERROR(ROUND(100*O{x}/(O{x}+Q{x}),1),0)", f["td_num"], tt.attendance_pct or 0),
                (f"=IFERROR(ROUND(100*P{x}/O{x},1),0)", f["td_num"], tt.late_pct or 0),
            ]
            for c, (formula, fmt, value) in enumerate(cells, start=13):
                ws.write_formula(r0, c, formula, fmt, value)
        return r + 1 + len(att.departments)

    def _buckets_block(self, r: int) -> int:
        ws, f = self.ws_summary, self.f
        ws.write(r, 12, "Распределение времени прихода (рабочие дни)", f["section"])
        for c, h in enumerate(["Интервал", "Приходов", "с", "до"]):
            ws.write(r + 1, 12 + c, h, f["th"])
        start = minutes(self.s.work_start)
        edges = [0, *[max(0, min(1439, start + d)) for d in (-60, -30, 0, 15, 30, 60, 120)], 1440]
        edges = sorted(set(edges))
        fmt_m = lambda m: f"{m // 60:02d}:{m % 60:02d}"  # noqa: E731
        work_in = [minutes(s.passes.first_in) for s in self.results if s.row.workday and s.passes.first_in]
        self.bucket_rows = (r + 2, r + len(edges))
        for j, (lo, hi) in enumerate(zip(edges, edges[1:])):
            r0 = r + 2 + j
            x = r0 + 1
            label = f"до {fmt_m(hi)}" if lo == 0 else (f"после {fmt_m(lo)}" if hi == 1440 else f"{fmt_m(lo)}–{fmt_m(hi)}")
            ws.write_string(r0, 12, label, f["td"])
            ws.write_number(r0, 14, lo / 1440, f["td_muted"])
            ws.write_number(r0, 15, hi / 1440, f["td_muted"])
            formula = (f'=COUNTIFS({self.filt()},{self.d("G")},">="&$O{x},{self.d("G")},"<"&$P{x},'
                       f'{self.d("P")},1)')
            ws.write_formula(r0, 13, formula, f["td_int"], sum(lo <= m < hi for m in work_in))
        return r + len(edges)

    def _top_block(self, r: int) -> None:
        ws, f = self.ws_summary, self.f
        ws.write_formula(r, 12, '="Топ-10: "&fRank', f["section"], f"Топ-10: {RANKS[0]}")
        self.top_title = (r, 12)
        for c, h in enumerate(["Сотрудник", "Отдел", "Значение"]):
            ws.write(r + 1, 12 + c, h, f["th"])
        cache = sorted(self.emp_values, key=lambda v: (-v[0]))
        keys = f"{q(CALC)}!$V$2:$V${self.emp_last}"
        col = lambda c: f"{q(CALC)}!${c}$2:${c}${self.emp_last}"  # noqa: E731
        self.top_rows = (r + 2, r + 11)
        for k in range(10):
            r0 = r + 2 + k
            x = r0 + 1
            name = dept = ""
            value: float | str = NA
            idx = 0
            if k < len(cache):
                _, row_i, emp, val = cache[k]
                idx = row_i - 1
                if val > 0:
                    name, dept, value = emp.name, emp.department, val
            ws.write_formula(r0, 16, f"=IFERROR(MATCH(LARGE({keys},{k + 1}),{keys},0),0)", f["td_muted_int"], idx)
            ws.write_formula(r0, 12, f'=IF($Q{x}=0,"",IF(INDEX({col("U")},$Q{x})>0,INDEX({col("M")},$Q{x}),""))',
                             f["td"], name)
            ws.write_formula(r0, 13, f'=IF($M{x}="","",INDEX({col("L")},$Q{x}))', f["td"], dept)
            ws.write_formula(r0, 14, f'=IF($M{x}="",NA(),INDEX({col("U")},$Q{x}))', f["td_num"], value)
        ws.set_column("Q:Q", 5, None, {"hidden": True})

    def _charts(self) -> None:
        ws, wb = self.ws_summary, self.wb
        S = q(SUMMARY)
        rng = lambda col, rows: f"={S}!${col}${rows[0] + 1}:${col}${rows[1] + 1}"  # noqa: E731
        font = {"name": "Calibri", "size": 9, "color": "#404040"}
        grid = {"visible": True, "line": {"color": "#E7E6E6"}}

        def style(chart, title, y_fmt=None):
            chart.set_title({"name": title, "name_font": {"name": "Calibri", "size": 12, "bold": True, "color": "#1F3864"}})
            chart.set_legend({"none": True})
            chart.set_chartarea({"border": {"color": "#D9D9D9"}})
            chart.set_y_axis({"num_font": font, "major_gridlines": grid, **({"num_format": y_fmt} if y_fmt else {})})

        daily = wb.add_chart({"type": "column"})
        daily.add_series({
            "categories": rng("M", self.daily_rows), "values": rng("T", self.daily_rows),
            "fill": {"color": "#2F5597"}, "gap": 60,
            "data_labels": {"value": True, "num_format": "General", "font": {"size": 8, "color": "#404040"}},
        })
        style(daily, f"={S}!{xl_rowcol_to_cell(*self.metric_title, True, True)}")
        daily.set_x_axis({"num_font": font, "num_format": "dd.mm", "date_axis": False})
        ws.insert_chart("B15", daily, {"x_scale": 1.95, "y_scale": 1.15})

        buckets = wb.add_chart({"type": "column"})
        buckets.add_series({
            "categories": rng("M", self.bucket_rows), "values": rng("N", self.bucket_rows),
            "fill": {"color": "#548235"}, "gap": 40,
            "data_labels": {"value": True, "font": {"size": 8}},
        })
        style(buckets, "Во сколько приходят")
        buckets.set_x_axis({"num_font": font})
        ws.insert_chart("B33", buckets, {"x_scale": 0.97, "y_scale": 1.1})

        dept = wb.add_chart({"type": "bar"})
        for col, color, name in (("R", "#2F5597", "Посещаемость, %"), ("S", "#C55A11", "Доля опозданий, %")):
            dept.add_series({
                "name": name, "categories": rng("M", self.dept_rows), "values": rng(col, self.dept_rows),
                "fill": {"color": color}, "data_labels": {"value": True, "font": {"size": 8}},
            })
        style(dept, "По отделам, %")
        dept.set_legend({"position": "bottom", "font": font})
        dept.set_x_axis({"num_font": font, "major_gridlines": grid, "min": 0, "max": 100})
        dept.set_y_axis({"num_font": font, "reverse": True})
        ws.insert_chart("G33", dept, {"x_scale": 0.97, "y_scale": 1.1})

        top = wb.add_chart({"type": "bar"})
        top.add_series({
            "categories": rng("M", self.top_rows), "values": rng("O", self.top_rows),
            "fill": {"color": "#C55A11"}, "gap": 50, "data_labels": {"value": True, "font": {"size": 8}},
        })
        style(top, f"={S}!{xl_rowcol_to_cell(*self.top_title, True, True)}")
        top.set_x_axis({"num_font": font, "major_gridlines": grid})
        top.set_y_axis({"num_font": font, "reverse": True})
        ws.insert_chart("B51", top, {"x_scale": 1.95, "y_scale": 1.3})

    # ---------- подписи ----------
    def _mode_caption(self) -> str:
        return "Приход/уход считаются по точкам: КПП и турникеты (переключается на листе «Сводка»)"

    def _mode_caption_formula(self) -> str:
        return ('="Приход/уход считаются по точкам: КПП и турникеты"'
                f'&IF(fDorm="{YES}"," + общежитие","")&IF(fTech="{YES}"," + тех. помещения","")'
                '&" (переключается на листе «Сводка»)"')


def build_report(att: Attendance) -> bytes:
    return Report(att).build()


SUMMARY_LABELS = {
    "ru": {"title": "Посещаемость", "emp": "Сотрудников", "wd": "рабочих дней", "att": "Посещаемость",
           "late": "Опозданий (после {start})", "absent": "Отсутствий", "early": "ранних уходов",
           "avg_in": "Средний приход", "top": "Чаще всего опаздывают",
           "note": "Учтены только КПП и турникеты; общежитие и техпомещения включаются на листе «Сводка»."},
    "uz": {"title": "Davomat", "emp": "Xodimlar", "wd": "ish kunlari", "att": "Davomat",
           "late": "Kechikishlar ({start} dan keyin)", "absent": "Kelmaganlar", "early": "erta ketishlar",
           "avg_in": "O‘rtacha kelish", "top": "Eng ko‘p kechikkanlar",
           "note": "Faqat KPP va turniketlar hisobga olingan; yotoqxona va texnik xonalar «Сводка» varag‘ida yoqiladi."},
    "en": {"title": "Attendance", "emp": "Employees", "wd": "work days", "att": "Attendance",
           "late": "Late arrivals (after {start})", "absent": "Absences", "early": "early leaves",
           "avg_in": "Average arrival", "top": "Most often late",
           "note": "Only checkpoints and turnstiles are counted; dorms and tech rooms can be enabled on the «Сводка» sheet."},
}


def summary_text(att: Attendance, lang: str = "ru", limit: int = 5) -> str:
    """Короткая сводка для подписи к файлу в Telegram (HTML)."""
    L = SUMMARY_LABELS.get(lang, SUMMARY_LABELS["ru"])
    results = [evaluate(r, att.settings) for r in att.rows]
    t = totals(results)
    exp = att.export
    lines = [
        f"<b>{L['title']} {exp.begin:%d.%m} – {exp.end:%d.%m.%Y}</b>",
        f"{L['emp']}: {len(att.employees)} · {L['wd']}: {sum(att.is_workday(d) for d in att.dates)}",
        f"{L['att']}: {t.attendance_pct if t.attendance_pct is not None else '—'}%",
        f"{L['late'].format(start=f'{att.settings.work_start:%H:%M}')}: {t.late} ({t.late_pct or 0}%)",
        f"{L['absent']}: {t.absent} · {L['early']}: {t.early}",
    ]
    if t.avg_in is not None:
        avg_in = round(t.avg_in * 1440)
        lines.append(f"{L['avg_in']}: {avg_in // 60:02d}:{avg_in % 60:02d}")
    late_by: dict[str, int] = {}
    for r in results:
        if r.status == LATE:
            late_by[r.row.employee.name] = late_by.get(r.row.employee.name, 0) + 1
    top = sorted(late_by.items(), key=lambda kv: -kv[1])[:limit]
    if top:
        lines += ["", f"<b>{L['top']}:</b>"]
        lines += [f"{i}. {_esc(name)} — {n}" for i, (name, n) in enumerate(top, 1)]
    lines += ["", f"<i>{L['note']}</i>"]
    return "\n".join(lines)


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def report_filename(att: Attendance) -> str:
    exp = att.export
    return f"Посещаемость_{exp.begin:%d.%m}-{exp.end:%d.%m.%Y}.xlsx"

