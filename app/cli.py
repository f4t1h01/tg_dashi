"""Сборка отчёта без бота: python -m app.cli выгрузка.xlsx [отчёт.xlsx]"""
from __future__ import annotations

import sys
from pathlib import Path

from .analytics import build_attendance
from .config import ReportSettings
from .parser import parse_workbook
from .report import build_report, report_filename, summary_text


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 2
    src = Path(argv[0])
    att = build_attendance(parse_workbook(src.read_bytes()), ReportSettings.from_env())
    out = Path(argv[1]) if len(argv) > 1 else src.with_name(report_filename(att))
    out.write_bytes(build_report(att))
    print(summary_text(att))
    print(f"\n→ {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
