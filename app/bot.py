from __future__ import annotations

import asyncio
import io
import logging

from aiogram import BaseMiddleware, Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.types import BotCommand, BufferedInputFile, Document, Message

from .analytics import build_attendance
from .config import BotConfig, ReportSettings
from .i18n import START, WEEKDAYS, lang_of, t
from .parser import ParseError, parse_workbook
from .progress import run_with_progress
from .report import build_report, report_filename, summary_text

log = logging.getLogger("tg_dashi")
router = Router()
MAX_FILE_SIZE = 20 * 1024 * 1024  # лимит скачивания Bot API

COMMANDS = {
    "ru": [("start", "Начать"), ("help", "Как пользоваться"), ("settings", "Правила расчёта")],
    "uz": [("start", "Boshlash"), ("help", "Qanday foydalanish"), ("settings", "Hisoblash qoidalari")],
    "en": [("start", "Start"), ("help", "How to use"), ("settings", "Calculation rules")],
}


class AccessMiddleware(BaseMiddleware):
    def __init__(self, allowed: frozenset[int]):
        self.allowed = allowed

    async def __call__(self, handler, event: Message, data):
        user = event.from_user
        data["lang"] = lang_of(user.language_code if user else None)
        if self.allowed and (user is None or user.id not in self.allowed):
            await event.answer(f"⛔ ID: <code>{user.id if user else '?'}</code>")
            return None
        return await handler(event, data)


def process(content: bytes, settings: ReportSettings, lang: str) -> tuple[bytes, str, str]:
    att = build_attendance(parse_workbook(content), settings)
    return build_report(att), report_filename(att), summary_text(att, lang)


async def build(bot: Bot, doc: Document, settings: ReportSettings, lang: str) -> tuple[bytes, str, str]:
    buf = io.BytesIO()
    await bot.download(doc, destination=buf)
    return await asyncio.to_thread(process, buf.getvalue(), settings, lang)


@router.message(CommandStart())
async def on_start(message: Message) -> None:
    await message.answer(START)


@router.message(Command("help"))
async def on_help(message: Message, lang: str) -> None:
    await message.answer(t(lang, "help"))


@router.message(Command("settings"))
async def on_settings(message: Message, config: BotConfig, lang: str) -> None:
    s = config.report
    await message.answer(t(
        lang, "settings",
        start=f"{s.work_start:%H:%M}", end=f"{s.work_end:%H:%M}", grace=s.late_grace_min,
        days=", ".join(WEEKDAYS[lang][d - 1] for d in sorted(s.workdays)),
    ))


@router.message(F.document)
async def on_document(message: Message, bot: Bot, config: BotConfig, lang: str) -> None:
    doc = message.document
    if not (doc.file_name or "").lower().endswith((".xlsx", ".xlsm")):
        await message.reply(t(lang, "need_xlsx"))
        return
    if doc.file_size and doc.file_size > MAX_FILE_SIZE:
        await message.reply(t(lang, "too_big"))
        return

    progress = await message.reply("⏳")
    try:
        content, filename, caption = await run_with_progress(progress, lang, build(bot, doc, config.report, lang))
    except ParseError as exc:
        await progress.edit_text(t(lang, "parse_error", error=exc))
        return
    except Exception:
        log.exception("report failed for %s", doc.file_name)
        await progress.edit_text(t(lang, "failed"))
        return

    log.info("report %s for user %s", filename, message.from_user.id if message.from_user else "?")
    await message.reply_document(BufferedInputFile(content, filename=filename), caption=caption[:1024])


@router.message()
async def on_other(message: Message, lang: str) -> None:
    await message.answer(t(lang, "other"))


async def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    config = BotConfig.from_env()
    bot = Bot(config.token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(config=config)
    router.message.outer_middleware(AccessMiddleware(config.allowed_users))
    dp.include_router(router)
    await bot.set_my_commands([BotCommand(command=c, description=d) for c, d in COMMANDS["ru"]])
    for lang in ("uz", "en"):
        await bot.set_my_commands([BotCommand(command=c, description=d) for c, d in COMMANDS[lang]],
                                  language_code=lang)
    log.info("bot started, access: %s", "restricted" if config.allowed_users else "open")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
