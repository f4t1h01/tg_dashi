"""Анимированный прогресс в одном сообщении, пока отчёт строится в фоне.

Минимум MIN_DURATION секунд. До 50 % полоска идёт равномерно; если работа
к этому моменту не закончилась — замедляется и подползает к 95 %, пока
задача не завершится. Когда задача готова, оставшийся путь до 100 %
растягивается так, чтобы общее время было не меньше MIN_DURATION.
"""
from __future__ import annotations

import asyncio
import logging
import time
from collections.abc import Awaitable
from typing import TypeVar

from aiogram.exceptions import TelegramBadRequest, TelegramRetryAfter
from aiogram.types import Message

from .i18n import t

log = logging.getLogger(__name__)
T = TypeVar("T")

MIN_DURATION = 5.0
TICK = 0.7          # Telegram спокойно переносит ~1 редактирование в секунду
SLOW_CEILING = 95.0
SLOW_FACTOR = 0.07
BAR_WIDTH = 14
SPINNER = "◐◓◑◒"


def next_percent(p: float, done: bool, elapsed: float, tick: float = TICK, min_duration: float = MIN_DURATION) -> float:
    if done:
        remaining = min_duration - elapsed
        if remaining > tick:
            return min(100.0, p + (100.0 - p) * tick / remaining)
        # минимум уже прошёл: добегаем за 1–2 шага, без резкого скачка с середины
        return 100.0 if 100.0 - p <= 25 else p + (100.0 - p) / 2
    if p < 50:
        return min(50.0, p + 100.0 * tick / min_duration)
    return p + (SLOW_CEILING - p) * SLOW_FACTOR


def render(lang: str, percent: float, elapsed: float, frame: int) -> str:
    filled = round(BAR_WIDTH * percent / 100)
    bar = "█" * filled + "░" * (BAR_WIDTH - filled)
    stages = t(lang, "stages").split("|")
    stage = stages[min(len(stages) - 1, int(percent / 100 * len(stages)))]
    spin = SPINNER[frame % len(SPINNER)]
    return (f"{spin} <b>{t(lang, 'progress_title')}</b>\n"
            f"<code>{bar}</code> {int(percent)}%\n"
            f"{stage}… · ⏱ {elapsed:.0f} {t(lang, 'sec')}")


async def _edit(message: Message, text: str) -> None:
    try:
        await message.edit_text(text)
    except TelegramRetryAfter as exc:
        await asyncio.sleep(exc.retry_after)
    except TelegramBadRequest as exc:
        if "not modified" not in str(exc):
            log.warning("progress edit failed: %s", exc)


async def run_with_progress(message: Message, lang: str, work: Awaitable[T]) -> T:
    """Показывает прогресс в `message`, пока выполняется `work`. Ошибки work пробрасываются сразу."""
    task = asyncio.ensure_future(work)
    start = time.monotonic()
    percent, frame = 0.0, 0
    while True:
        elapsed = time.monotonic() - start
        if task.done() and task.exception() is not None:
            raise task.exception()
        percent = next_percent(percent, task.done(), elapsed)
        if percent >= 100:
            break
        await _edit(message, render(lang, percent, elapsed, frame))
        frame += 1
        await asyncio.wait({task}, timeout=TICK)
        # если задача закончилась раньше тика — доспим остаток, чтобы шаг анимации был ровным
        rest = TICK - (time.monotonic() - start - elapsed)
        if rest > 0:
            await asyncio.sleep(rest)
    elapsed = time.monotonic() - start
    await _edit(message, f"{t(lang, 'done')} · <code>{'█' * BAR_WIDTH}</code> 100% · ⏱ {elapsed:.0f} {t(lang, 'sec')}")
    return task.result()
