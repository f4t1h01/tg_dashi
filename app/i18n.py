"""Тексты бота на русском, узбекском (латиница) и английском."""
from __future__ import annotations

LANGS = ("ru", "uz", "en")
DEFAULT_LANG = "ru"

START = (
    "👋 <b>Здравствуйте!</b> Отправьте мне файл выгрузки СКУД <b>«Passages through perimeter»</b> (.xlsx) — "
    "я соберу отчёт: сводку с дашбордами, посещаемость за месяц и по дням.\n\n"
    "👋 <b>Assalomu alaykum!</b> Menga <b>«Passages through perimeter»</b> (.xlsx) faylini yuboring — "
    "hisobot tayyorlab beraman: dashboardlar bilan umumiy ko‘rinish, oylik va kunlik davomat.\n\n"
    "👋 <b>Hello!</b> Send me the <b>«Passages through perimeter»</b> access-control export (.xlsx) — "
    "I will build a report: a dashboard summary, monthly and daily attendance.\n\n"
    "📎 ⬇️"
)

TEXTS: dict[str, dict[str, str]] = {
    "ru": {
        "help": "Пришлите выгрузку .xlsx файлом. Отчёт: «Сводка» (фильтры и графики), «За месяц», «За день».\n"
                "/settings — правила расчёта",
        "settings": "Рабочий день: {start}–{end}\nДопуск опоздания: {grace} мин\nРабочие дни: {days}\n"
                    "Приход/уход по умолчанию — только КПП и турникеты.\n\n"
                    "Всё это можно поменять прямо в Excel на листе «Сводка».",
        "need_xlsx": "Нужен файл Excel (.xlsx) — выгрузка «Passages through perimeter».",
        "too_big": "Файл больше 20 МБ — Telegram не даст боту его скачать.",
        "parse_error": "❌ Не получилось разобрать файл: {error}",
        "failed": "❌ Ошибка при построении отчёта. Проверьте, что это выгрузка «Passages through perimeter».",
        "other": "Пришлите .xlsx-выгрузку файлом 📎",
        "progress_title": "Обработка файла",
        "done": "✅ Готово",
        "stages": "Скачиваю файл|Читаю проходы|Собираю дни сотрудников|Считаю опоздания и уходы|"
                  "Строю дашборды|Оформляю листы|Почти готово",
        "sec": "с",
    },
    "uz": {
        "help": "Faylni .xlsx ko‘rinishida yuboring. Hisobot: «Сводка» (filtrlar va grafiklar), «За месяц», «За день».\n"
                "/settings — hisoblash qoidalari",
        "settings": "Ish kuni: {start}–{end}\nKechikishga ruxsat: {grace} daq\nIsh kunlari: {days}\n"
                    "Kelish/ketish sukut bo‘yicha faqat KPP va turniketlar bo‘yicha.\n\n"
                    "Bularning barchasini Excel’dagi «Сводка» varag‘ida o‘zgartirish mumkin.",
        "need_xlsx": "Excel fayli (.xlsx) kerak — «Passages through perimeter» eksporti.",
        "too_big": "Fayl 20 MB dan katta — Telegram botga uni yuklab olishga ruxsat bermaydi.",
        "parse_error": "❌ Faylni o‘qib bo‘lmadi: {error}",
        "failed": "❌ Hisobot tuzishda xatolik. Bu «Passages through perimeter» eksporti ekanini tekshiring.",
        "other": "Iltimos, .xlsx faylni yuboring 📎",
        "progress_title": "Fayl qayta ishlanmoqda",
        "done": "✅ Tayyor",
        "stages": "Fayl yuklanmoqda|O‘tishlar o‘qilmoqda|Xodimlar kunlari yig‘ilmoqda|"
                  "Kechikishlar hisoblanmoqda|Dashboardlar tuzilmoqda|Varaqlar bezatilmoqda|Deyarli tayyor",
        "sec": "s",
    },
    "en": {
        "help": "Send the .xlsx export as a file. Report sheets: «Сводка» (filters and charts), «За месяц», «За день».\n"
                "/settings — calculation rules",
        "settings": "Work day: {start}–{end}\nLate grace: {grace} min\nWork days: {days}\n"
                    "By default arrival/leaving use only checkpoints and turnstiles.\n\n"
                    "All of this can be changed in Excel on the «Сводка» sheet.",
        "need_xlsx": "Please send an Excel file (.xlsx) — the «Passages through perimeter» export.",
        "too_big": "The file is larger than 20 MB — Telegram won't let the bot download it.",
        "parse_error": "❌ Couldn't read the file: {error}",
        "failed": "❌ Failed to build the report. Make sure it's a «Passages through perimeter» export.",
        "other": "Please send the .xlsx export as a file 📎",
        "progress_title": "Processing file",
        "done": "✅ Done",
        "stages": "Downloading|Reading passes|Grouping employee days|Counting late arrivals|"
                  "Building dashboards|Formatting sheets|Almost done",
        "sec": "s",
    },
}

WEEKDAYS = {
    "ru": ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"],
    "uz": ["Du", "Se", "Ch", "Pa", "Ju", "Sh", "Ya"],
    "en": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
}


def lang_of(code: str | None) -> str:
    code = (code or "").split("-")[0].lower()
    return code if code in LANGS else DEFAULT_LANG


def t(lang: str, key: str, **kwargs) -> str:
    text = TEXTS[lang].get(key) or TEXTS[DEFAULT_LANG][key]
    return text.format(**kwargs) if kwargs else text
