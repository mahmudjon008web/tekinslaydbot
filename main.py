from __future__ import annotations

import asyncio
import json
import logging
import os
import random
import re
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.enums import ChatMemberStatus
from aiogram.exceptions import (
    TelegramBadRequest,
    TelegramNetworkError,
    TelegramRetryAfter,
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from aiogram.utils.keyboard import InlineKeyboardBuilder
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from dotenv import load_dotenv
from google import genai
from pptx import Presentation
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches as PptxInches
from pptx.util import Pt as PptxPt

from database import (
    add_channel,
    add_user,
    check_and_update_limit,
    get_all_users,
    get_channels,
    get_remaining_cooldown,
    get_subscription_mode,
    get_user_count,
    init_db,
    release_limit,
    remove_channel,
    set_subscription_mode,
)

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
ADMIN_ID = int(os.getenv("ADMIN_ID", "0") or 0)

SERVICE_NAMES = {
    "slayd": "Taqdimot (Slayd)",
    "mustaqil_ish": "Mustaqil ish",
    "referat": "Referat",
    "tezis": "Tezis",
    "maqola": "Maqola",
}

BUTTON_TO_SERVICE = {
    "🆕 Taqdimot (Slayd) Yaratish": "slayd",
    "📁 Mustaqil Ish Yaratish": "mustaqil_ish",
    "📚 Referat Yaratish": "referat",
    "🎓 Tezis yaratish": "tezis",
    "✅ Maqola yaratish": "maqola",
}

router = Router()


def parse_env_channels() -> list[tuple[str, str]]:
    """Legacy .env-based channel list, used only once to migrate into the DB."""
    channels = []
    for item in os.getenv("REQUIRED_CHANNELS", "").split(","):
        if not item.strip():
            continue
        name, _, url = item.partition("|")
        name = name.strip()
        url = url.strip() or f"https://t.me/{name.lstrip('@')}"
        channels.append((name, url))
    return channels


def migrate_env_channels_if_needed() -> None:
    """One-time migration: if the DB has no channels yet but .env has
    REQUIRED_CHANNELS set, copy them into the DB so nothing is lost."""
    if get_channels():
        return
    for name, url in parse_env_channels():
        channel_ref = name if name.startswith("@") else f"@{name.lstrip('@')}"
        add_channel(channel_ref, name, url)


class WorkState(StatesGroup):
    waiting_for_topic = State()
    waiting_for_ad = State()
    waiting_for_channel_add = State()


def main_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🆕 Taqdimot (Slayd) Yaratish")],
            [
                KeyboardButton(text="📁 Mustaqil Ish Yaratish"),
                KeyboardButton(text="📚 Referat Yaratish"),
            ],
            [
                KeyboardButton(text="🎓 Tezis yaratish"),
                KeyboardButton(text="✅ Maqola yaratish"),
            ],
        ],
        resize_keyboard=True,
        input_field_placeholder="Xizmatni tanlang",
    )


def subscription_keyboard(channels: list[tuple[str, str, str]]) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    for channel_ref, title, url in channels:
        builder.add(InlineKeyboardButton(text=f"Obuna bo‘lish: {title}", url=url))
    builder.add(InlineKeyboardButton(text="Obunani tekshirish", callback_data="check_subscription"))
    builder.adjust(1)
    return builder.as_markup()


async def get_missing_channels(bot: Bot, user_id: int) -> list[tuple[str, str, str]]:
    """Foydalanuvchi hali obuna bo'lmagan majburiy kanallar ro'yxatini
    qaytaradi (masalan, 3 tadan 2 tasiga obuna bo'lgan bo'lsa, faqat
    qolgan 1 tasini qaytaradi — barcha 3 tasini emas)."""
    channels = get_channels()
    if not channels:
        return []

    # "soft" rejimda haqiqiy tekshiruv qilinmaydi — shuning uchun bot
    # kanallarda admin bo'lishi shart emas.
    if get_subscription_mode() == "soft":
        return []

    missing: list[tuple[str, str, str]] = []
    for channel_ref, title, url in channels:
        try:
            member = await bot.get_chat_member(channel_ref, user_id)
            if member.status in {ChatMemberStatus.LEFT, ChatMemberStatus.KICKED}:
                missing.append((channel_ref, title, url))
        except (TelegramBadRequest, TelegramNetworkError) as exc:
            logger.warning("Subscription check failed for %s: %s", channel_ref, exc)
            missing.append((channel_ref, title, url))
    return missing


async def is_subscribed(bot: Bot, user_id: int) -> bool:
    return not await get_missing_channels(bot, user_id)


async def require_subscription(message: Message, bot: Bot) -> bool:
    user = message.from_user
    if not user:
        return False
    missing = await get_missing_channels(bot, user.id)
    if not missing:
        return True
    await message.answer(
        "Botdan foydalanish uchun quyidagi kanal(lar)ga obuna bo‘ling:",
        reply_markup=subscription_keyboard(missing),
    )
    return False


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_-]+", "_", value)
    return cleaned[:48].strip("_") or "material"


# ---------------------------------------------------------------------------
# Markdown cleanup — Gemini ba'zan strukturaviy JSON o'rniga oddiy matn
# qaytarganda (fallback holatida) ishlatiladi, chunki python-docx/pptx
# markdown belgilarini tushunmaydi.
# ---------------------------------------------------------------------------
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")
_UNDERSCORE_BOLD_RE = re.compile(r"__(.+?)__")
_UNDERSCORE_ITALIC_RE = re.compile(r"(?<!_)_(?!_)(.+?)(?<!_)_(?!_)")
_INLINE_CODE_RE = re.compile(r"`(.+?)`")
_HEADING_RE = re.compile(r"^#{1,6}\s*(.+)$")
_BULLET_RE = re.compile(r"^[-*•]\s+(.+)$")


def clean_markdown(raw: str) -> tuple[str, bool]:
    """Strip common markdown syntax. Returns (clean_text, is_heading_like)."""
    text = raw.strip()
    is_heading = False

    heading_match = _HEADING_RE.match(text)
    if heading_match:
        text = heading_match.group(1).strip()
        is_heading = True

    bullet_match = _BULLET_RE.match(text)
    if bullet_match:
        text = "• " + bullet_match.group(1).strip()

    text = _BOLD_RE.sub(r"\1", text)
    text = _UNDERSCORE_BOLD_RE.sub(r"\1", text)
    text = _ITALIC_RE.sub(r"\1", text)
    text = _UNDERSCORE_ITALIC_RE.sub(r"\1", text)
    text = _INLINE_CODE_RE.sub(r"\1", text)
    text = text.strip()

    if not is_heading and not text.startswith("•"):
        if text.endswith(":") or len(text) < 60:
            is_heading = True

    return text, is_heading


def format_remaining(hours: float) -> str:
    """23.92781 kabi qiymatni "23 soat 56 daqiqa" ko'rinishiga o'tkazadi."""
    total_minutes = round(hours * 60)
    h, m = divmod(total_minutes, 60)
    if h and m:
        return f"{h} soat {m} daqiqa"
    if h:
        return f"{h} soat"
    return f"{max(m, 1)} daqiqa"


# ---------------------------------------------------------------------------
# AI orqali kontent generatsiya qilish — strukturaviy JSON so'raladi.
# ---------------------------------------------------------------------------

_JSON_FENCE_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.MULTILINE)

MIN_SLIDES = 15
MAX_SLIDES = 22

SLIDES_SCHEMA_HINT = f"""Faqat quyidagi JSON formatida javob bering, boshqa hech qanday matn, izoh yoki markdown qo'shmang:
{{
  "title": "Taqdimotning qisqa va aniq sarlavhasi",
  "slides": [
    {{"title": "Slayd sarlavhasi", "bullets": ["Qisqa va aniq punkt 1", "Qisqa va aniq punkt 2", "Qisqa va aniq punkt 3"]}}
  ]
}}
Talablar:
- KAMIDA {MIN_SLIDES} ta, ko'pi bilan {MAX_SLIDES} ta slayd yarating (kirish, asosiy mavzular bir nechta bo'limga bo'linib, xulosa mantig'ida). {MIN_SLIDES} tadan kam slayd yaratish TAQIQLANADI — mavzuni kerak bo'lsa mayda kichik pastki mavzularga bo'lib, slaydlar sonini oshiring.
- Har bir slaydda 3 tadan 5 tagacha bullet bo'lsin.
- Har bir bullet 1-2 gapdan oshmasin, telegraf uslubida emas, tushunarli va ma'noli yozing.
- Markdown belgilaridan (**, ##, -, `) foydalanmang, bullet matnini toza yozing."""

DOCUMENT_SCHEMA_HINT = """Faqat quyidagi JSON formatida javob bering, boshqa hech qanday matn, izoh yoki markdown qo'shmang:
{
  "title": "Ish sarlavhasi",
  "sections": [
    {"heading": "Kirish", "paragraphs": ["Chuqur va atroflicha yozilgan paragraf matni."]},
    {"heading": "1-bob: ...", "paragraphs": ["..."], "bullets": ["Ixtiyoriy: agar shu bo'limda ro'yxat kerak bo'lsa"]},
    {"heading": "Xulosa", "paragraphs": ["..."]}
  ]
}
Talablar:
- Kamida 5 ta, ko'pi bilan 8 ta bo'lim (Kirish va Xulosa shart).
- Har bir bo'limda kamida 2 ta chuqur, mazmunli paragraf bo'lsin (har biri 3-5 gap).
- "bullets" maydoni faqat ro'yxat mantiqan kerak bo'lgan joyda ishlatilsin, aks holda uni umuman qo'shmang.
- Markdown belgilaridan (**, ##, -, `) foydalanmang."""


def _extract_json(raw_text: str) -> dict:
    text = raw_text.strip()
    text = _JSON_FENCE_RE.sub("", text).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start : end + 1]
    return json.loads(text)


def _ensure_min_slides(slides: list[dict], minimum: int = MIN_SLIDES) -> list[dict]:
    """Agar AI kamroq slayd qaytarsa, mavjud slaydlarni bulletlar bo'yicha
    ikkiga bo'lib, minimal songa yetkazishga harakat qilamiz."""
    if len(slides) >= minimum:
        return slides

    expanded: list[dict] = []
    for slide in slides:
        bullets = [b for b in (slide.get("bullets") or []) if str(b).strip()]
        title = slide.get("title") or ""
        if len(bullets) <= 2:
            expanded.append({"title": title, "bullets": bullets})
            continue
        mid = len(bullets) // 2
        expanded.append({"title": title, "bullets": bullets[:mid]})
        expanded.append({"title": title, "bullets": bullets[mid:]})

    return expanded


def _call_gemini(prompt: str) -> str:
    client = genai.Client(api_key=GEMINI_API_KEY)

    candidate_models = [
        "gemini-flash-latest",
        "gemini-3.6-flash",
        "gemini-pro-latest",
        "gemini-3.1-pro-preview",
    ]

    last_exception: Exception | None = None
    for model_name in candidate_models:
        for attempt in range(2):
            try:
                logger.info("Model sinab ko'rilmoqda: %s (urinish %d)", model_name, attempt + 1)
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                )
                if response and response.text:
                    return response.text.strip()
                last_exception = RuntimeError(f"{model_name} bo'sh javob qaytardi.")
                break
            except Exception as e:
                logger.warning("Model %s xatosi (urinish %d): %s", model_name, attempt + 1, e)
                last_exception = e
                is_overloaded = "503" in str(e) or "UNAVAILABLE" in str(e)
                if is_overloaded and attempt == 0:
                    time.sleep(3)
                    continue
                break

    raise last_exception or RuntimeError("Barcha Gemini modellarida xatolik yuz berdi.")


def generate_slides_data(topic: str) -> dict:
    prompt = f"""Siz tajribali o'zbek tilidagi taqdimot muallifisiz.
Mavzu: {topic}

{SLIDES_SCHEMA_HINT}"""
    raw = _call_gemini(prompt)
    try:
        data = _extract_json(raw)
        slides = data.get("slides") or []
        if not slides:
            raise ValueError("slides bo'sh")
        data["slides"] = _ensure_min_slides(slides)
        return data
    except (json.JSONDecodeError, ValueError, AttributeError) as exc:
        logger.warning("Slayd JSON parse xatosi, fallback ishlatiladi: %s", exc)
        return _fallback_slides_from_text(topic, raw)


def generate_document_data(service: str, topic: str) -> dict:
    prompt = f"""Siz tajribali o'zbek tilidagi akademik muallifsiz.
{SERVICE_NAMES.get(service, 'Material')} uchun chuqur va atroflicha matn yarating.
Mavzu: {topic}

{DOCUMENT_SCHEMA_HINT}"""
    raw = _call_gemini(prompt)
    try:
        data = _extract_json(raw)
        sections = data.get("sections") or []
        if not sections:
            raise ValueError("sections bo'sh")
        return data
    except (json.JSONDecodeError, ValueError, AttributeError) as exc:
        logger.warning("Hujjat JSON parse xatosi, fallback ishlatiladi: %s", exc)
        return _fallback_document_from_text(topic, raw)


def _fallback_slides_from_text(topic: str, raw_text: str) -> dict:
    lines = []
    for raw_line in raw_text.split("\n"):
        raw_line = raw_line.strip()
        if not raw_line:
            continue
        text, _ = clean_markdown(raw_line)
        if text:
            lines.append(text)

    chunks = [lines[i : i + 3] for i in range(0, len(lines), 3)]
    slides = [
        {"title": f"{idx + 1}-qism", "bullets": chunk}
        for idx, chunk in enumerate(chunks[:MAX_SLIDES])
    ]
    return {"title": topic, "slides": slides or [{"title": topic, "bullets": [topic]}]}


def _fallback_document_from_text(topic: str, raw_text: str) -> dict:
    sections: list[dict] = []
    current_heading = "Kirish"
    current_paragraphs: list[str] = []

    def _flush() -> None:
        if current_paragraphs:
            sections.append({"heading": current_heading, "paragraphs": list(current_paragraphs)})

    for raw_line in raw_text.split("\n"):
        raw_line = raw_line.strip()
        if not raw_line:
            continue
        heading_match = _HEADING_RE.match(raw_line)
        if heading_match:
            _flush()
            current_heading = heading_match.group(1).strip()
            current_paragraphs = []
            continue

        text, _ = clean_markdown(raw_line)
        if text:
            current_paragraphs.append(text)
    _flush()

    return {"title": topic, "sections": sections or [{"heading": topic, "paragraphs": [raw_text.strip() or topic]}]}


# ---------------------------------------------------------------------------
# Word hujjat yaratish — strukturaviy JSON asosida, to'g'ri Heading
# darajalari, "List Bullet" uslubi va sarlavhalar ostida aksent chiziq bilan.
# ---------------------------------------------------------------------------

DOC_THEMES = ["1D4ED8", "0F766E", "9A3412", "7C3AED", "B91C6C", "0369A1"]


def _set_paragraph_bottom_border(paragraph, color_hex: str, size: int = 10) -> None:
    """Paragraf ostiga rangli chiziq (bo'luvchi) qo'shadi — sarlavhalarni
    yanada dizaynliroq ko'rsatish uchun."""
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color_hex)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def build_document(service: str, topic: str, data: dict) -> Path:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(0.8)

    accent_hex = random.choice(DOC_THEMES)
    accent = RGBColor.from_string(accent_hex)
    title_text = data.get("title") or topic

    # --- Muqova sahifasi ---
    document.add_paragraph("\n\n")
    kicker = document.add_paragraph(SERVICE_NAMES.get(service, "Hujjat").upper())
    kicker.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if kicker.runs:
        kicker.runs[0].font.size = Pt(12)
        kicker.runs[0].font.color.rgb = accent
        kicker.runs[0].bold = True

    title = document.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.color.rgb = accent
    _set_paragraph_bottom_border(title, accent_hex, size=16)

    subtitle = document.add_paragraph(f"Mavzu: {topic}")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if subtitle.runs:
        subtitle.runs[0].italic = True
        subtitle.runs[0].font.size = Pt(13)

    document.add_page_break()

    # --- Asosiy qism ---
    for sec in data.get("sections", []):
        heading_text = (sec.get("heading") or "").strip()
        if heading_text:
            h = document.add_heading(heading_text, level=1)
            for r in h.runs:
                r.font.color.rgb = accent
            _set_paragraph_bottom_border(h, accent_hex, size=8)

        for para_text in sec.get("paragraphs", []) or []:
            para_text = (para_text or "").strip()
            if not para_text:
                continue
            p = document.add_paragraph(para_text)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if p.runs:
                p.runs[0].font.size = Pt(12)

        for bullet_text in sec.get("bullets", []) or []:
            bullet_text = (bullet_text or "").strip()
            if not bullet_text:
                continue
            document.add_paragraph(bullet_text, style="List Bullet")

    path = Path(tempfile.gettempdir()) / f"{safe_filename(topic)}_{service}.docx"
    document.save(path)
    return path


# ---------------------------------------------------------------------------
# PowerPoint taqdimot yaratish — strukturaviy JSON, professional dizayn.
#
# - Har bir TAQDIMOT bitta random temadan foydalanadi (fon + aksent rang).
# - Har bir kontent slaydi 10 xil LAYOUT (uslub)dan birida chiziladi, shu
#   bilan bir xil "burchakdagi doira" ko'rinishi takrorlanmaydi.
# - Pillow mavjud bo'lsa, har bir taqdimot uchun bir nechta abstrakt
#   dekorativ RASM generatsiya qilinadi va slaydlarga joylashtiriladi
#   (haqiqiy rasm — sof shakllardan ko'ra chiroyliroq ko'rinadi).
# ---------------------------------------------------------------------------

PPT_THEMES = [
    {"name": "dark_teal", "background": "18212F", "surface": "243247", "text": "F4F7FB", "accent": "35C2B5"},
    {"name": "ocean_blue", "background": "F7F9FC", "surface": "E1EAF7", "text": "18212F", "accent": "2F6BFF"},
    {"name": "royal_purple", "background": "2C1854", "surface": "432571", "text": "FFF8F0", "accent": "FFB86B"},
    {"name": "warm_academic", "background": "F1EEE7", "surface": "E4DCCB", "text": "24313A", "accent": "A34F35"},
    {"name": "forest", "background": "0F2318", "surface": "1C3B27", "text": "EAF4EC", "accent": "6FCF97"},
    {"name": "sunset_rose", "background": "2B1B2E", "surface": "432A46", "text": "FBEFEF", "accent": "F2678B"},
    {"name": "sand_coral", "background": "FBF3EA", "surface": "F0DFCB", "text": "3B2A22", "accent": "E0623F"},
    {"name": "midnight_gold", "background": "12141F", "surface": "1F2333", "text": "F5F1E6", "accent": "E8B93A"},
]

SLIDE_W = 13.333
SLIDE_H = 7.5


def _ppt_rgb(hex_value: str) -> PptxRGBColor:
    return PptxRGBColor.from_string(hex_value)


def _set_slide_background(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _ppt_rgb(color)


def _add_text(slide, text: str, left: float, top: float, width: float, height: float, size: int, color: str, bold: bool = False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(PptxInches(left), PptxInches(top), PptxInches(width), PptxInches(height))
    frame = box.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.size = PptxPt(size)
    run.font.bold = bold
    run.font.color.rgb = _ppt_rgb(color)
    return box


def _add_shape(slide, shape_type, left: float, top: float, width: float, height: float, color: str):
    shape = slide.shapes.add_shape(shape_type, PptxInches(left), PptxInches(top), PptxInches(width), PptxInches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _ppt_rgb(color)
    shape.line.fill.background()
    return shape


def _add_accent_bar(slide, accent_hex: str, top: float = 1.5, left: float = 0.8) -> None:
    _add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, 1.6, 0.06, accent_hex)


def _add_footer(slide, accent_hex: str, topic: str, page_no: int, total: int, left_x: float = 0.8) -> None:
    _add_text(slide, topic[:55], left_x, 7.05, 8.0, 0.35, size=10, color=accent_hex)
    _add_text(slide, f"{page_no}/{total}", 12.2, 7.05, 0.8, 0.35, size=10, color=accent_hex, align=PP_ALIGN.RIGHT)


def _add_corner_visual(slide, theme: dict, accent_hex: str, corner: str = "br") -> None:
    """Slayd burchagiga nozik dekorativ shakl (doira) qo'yadi — rasm o'rnini
    bosuvchi, chuqurlik hissi beruvchi element. Har doim matn shakllaridan
    OLDIN chaqirilishi kerak (z-order: birinchi qo'shilgan shakl orqada
    qoladi)."""
    size = random.uniform(4.0, 6.0)
    offset = size * 0.6
    positions = {
        "br": (SLIDE_W - offset, SLIDE_H - offset),
        "tl": (-offset, -offset),
        "tr": (SLIDE_W - offset, -offset),
        "bl": (-offset, SLIDE_H - offset),
    }
    x, y = positions.get(corner, positions["br"])
    _add_shape(slide, MSO_SHAPE.OVAL, x, y, size, size, theme["surface"])


def _add_bullet_row(slide, theme: dict, accent_hex: str, index: int, text: str, top: float, left: float = 0.8) -> None:
    circle = _add_shape(slide, MSO_SHAPE.OVAL, left, top, 0.42, 0.42, accent_hex)
    circle_frame = circle.text_frame
    circle_frame.word_wrap = False
    cp = circle_frame.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    crun = cp.add_run()
    crun.text = str(index)
    crun.font.size = PptxPt(14)
    crun.font.bold = True
    crun.font.color.rgb = _ppt_rgb(theme["background"])

    _add_text(slide, text, left + 0.7, top - 0.05, 11.0 - left, 0.9, size=16, color=theme["text"])


# --- Layout 1: raqamli ro'yxat ----------------------------------------------
def _layout_numbered_list(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="br")
    _add_text(slide, title_text, 0.8, 0.55, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.35)
    row_height = min(1.05, 5.2 / max(len(bullets), 1))
    for b_idx, bullet_text in enumerate(bullets):
        _add_bullet_row(slide, theme, accent, b_idx + 1, bullet_text, top=1.75 + b_idx * row_height)
    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 2: kartalar to'ri ------------------------------------------------
def _layout_card_grid(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="tl")
    _add_text(slide, title_text, 0.8, 0.5, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.3)

    cards = bullets[:6]
    cols = 2
    card_w, card_h = 5.65, 2.15
    gap_x, gap_y = 0.5, 0.35
    start_x, start_y = 0.8, 1.7
    for i, bullet_text in enumerate(cards):
        col, row = i % cols, i // cols
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h, theme["surface"])
        _add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, card_w, 0.08, accent)
        _add_text(slide, bullet_text, x + 0.3, y + 0.3, card_w - 0.6, card_h - 0.5, size=14, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 3: yon panel ------------------------------------------------------
def _layout_split_panel(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    panel_w = 4.3
    _add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, panel_w, SLIDE_H, accent)

    _add_shape(slide, MSO_SHAPE.OVAL, panel_w - 1.6, SLIDE_H - 1.9, 2.6, 2.6, theme["background"])

    _add_text(slide, title_text, 0.5, SLIDE_H / 2 - 1.1, panel_w - 0.8, 2.2, size=24, color=theme["background"], bold=True)

    right_x = panel_w + 0.6
    row_h = min(1.1, 5.4 / max(len(bullets), 1))
    for i, bullet_text in enumerate(bullets):
        _add_shape(slide, MSO_SHAPE.RECTANGLE, right_x, 1.15 + i * row_h + 0.12, 0.32, 0.06, accent)
        _add_text(slide, bullet_text, right_x + 0.55, 0.85 + i * row_h, 12.0 - right_x, row_h, size=15, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total, left_x=right_x)
    return slide


# --- Layout 4: katta-urg'u (1-2 kalit fikr uchun) ----------------------------
def _layout_big_focus(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="tr")
    _add_text(slide, title_text, 0.8, 1.0, 11.8, 1.0, size=30, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=2.0)
    for i, bullet_text in enumerate(bullets[:2]):
        _add_text(slide, bullet_text, 0.8, 2.6 + i * 1.5, 11.8, 1.3, size=22, color=theme["text"])
    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 5: gorizontal vaqt chizig'i (ketma-ket qadamlar uchun) ----------
def _layout_timeline_row(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="tr")
    _add_text(slide, title_text, 0.8, 0.55, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.35)

    items = bullets[:5]
    n = max(len(items), 1)
    col_w = 11.0 / n
    line_y = 3.1
    _add_shape(slide, MSO_SHAPE.RECTANGLE, 0.95, line_y, 11.0, 0.05, accent)
    for i, text in enumerate(items):
        cx = 0.95 + col_w * i + col_w / 2
        circle = _add_shape(slide, MSO_SHAPE.OVAL, cx - 0.28, line_y - 0.24, 0.56, 0.56, accent)
        cf = circle.text_frame
        cf.word_wrap = False
        cp = cf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        r = cp.add_run()
        r.text = str(i + 1)
        r.font.bold = True
        r.font.size = PptxPt(16)
        r.font.color.rgb = _ppt_rgb(theme["background"])
        _add_text(slide, text, cx - col_w / 2 + 0.1, line_y + 0.55, col_w - 0.2, 2.2, size=13, color=theme["text"], align=PP_ALIGN.CENTER)

    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 6: checklist (kvadrat belgilar) ---------------------------------
def _layout_checklist(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="bl")
    _add_text(slide, title_text, 0.8, 0.55, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.35)

    row_height = min(1.0, 5.0 / max(len(bullets), 1))
    for i, text in enumerate(bullets):
        top = 1.75 + i * row_height
        _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.8, top, 0.4, 0.4, accent)
        _add_text(slide, text, 1.4, top - 0.05, 10.5, 0.9, size=16, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 7: ikki ustunli matn ---------------------------------------------
def _layout_two_column(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="tl")
    _add_text(slide, title_text, 0.8, 0.55, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.35)

    mid = (len(bullets) + 1) // 2
    left_items, right_items = bullets[:mid], bullets[mid:]
    row_h = min(1.0, 5.0 / max(mid, 1))
    for i, text in enumerate(left_items):
        top = 1.8 + i * row_h
        _add_shape(slide, MSO_SHAPE.RECTANGLE, 0.8, top + 0.1, 0.25, 0.25, accent)
        _add_text(slide, text, 1.25, top, 5.2, row_h, size=14, color=theme["text"])
    for i, text in enumerate(right_items):
        top = 1.8 + i * row_h
        _add_shape(slide, MSO_SHAPE.RECTANGLE, 6.9, top + 0.1, 0.25, 0.25, accent)
        _add_text(slide, text, 7.35, top, 5.2, row_h, size=14, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 8: diagonal aksent -----------------------------------------------
def _layout_diagonal(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_shape(slide, MSO_SHAPE.PARALLELOGRAM, -1.0, 0, 4.6, SLIDE_H, accent)
    _add_shape(slide, MSO_SHAPE.OVAL, SLIDE_W - 4.3, SLIDE_H - 4.3, 3.6, 3.6, theme["surface"])

    left_x = 4.0
    _add_text(slide, title_text, left_x, 0.6, 8.5, 0.9, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.4, left=left_x)

    row_h = min(1.0, 5.3 / max(len(bullets), 1))
    for i, text in enumerate(bullets):
        top = 1.8 + i * row_h
        _add_text(slide, text, left_x, top, 8.3, row_h, size=15, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total, left_x=left_x)
    return slide


# --- Layout 9: harfli ikonka to'ri -------------------------------------------
def _layout_icon_grid(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="br")
    _add_text(slide, title_text, 0.8, 0.5, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.3)

    items = bullets[:6]
    cols = 3
    card_w, card_h = 3.65, 2.1
    gap_x, gap_y = 0.35, 0.3
    start_x, start_y = 0.8, 1.7
    letters = "ABCDEF"
    for i, text in enumerate(items):
        col, row = i % cols, i // cols
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        badge = _add_shape(slide, MSO_SHAPE.OVAL, x, y, 0.5, 0.5, accent)
        bf = badge.text_frame
        bf.word_wrap = False
        bp = bf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = letters[i % len(letters)]
        br.font.bold = True
        br.font.size = PptxPt(16)
        br.font.color.rgb = _ppt_rgb(theme["background"])
        _add_text(slide, text, x, y + 0.65, card_w, card_h - 0.6, size=13, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total)
    return slide


# --- Layout 10: markazlashgan urg'u kartasi ----------------------------------
def _layout_quote_highlight(presentation, blank_layout, theme, accent, title_text, bullets, topic, page_no, total):
    slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(slide, theme["background"])
    _add_corner_visual(slide, theme, accent, corner="tl")
    _add_text(slide, title_text, 0.8, 0.55, 11.8, 0.8, size=26, color=accent, bold=True)
    _add_accent_bar(slide, accent, top=1.35)

    _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.4, 1.9, 10.5, 4.6, theme["surface"])
    _add_shape(slide, MSO_SHAPE.RECTANGLE, 1.4, 1.9, 0.1, 4.6, accent)
    row_h = min(0.95, 4.1 / max(len(bullets), 1))
    for i, text in enumerate(bullets):
        _add_text(slide, text, 1.9, 2.2 + i * row_h, 9.6, row_h, size=16, color=theme["text"])

    _add_footer(slide, accent, topic, page_no, total)
    return slide


_ROTATING_LAYOUTS = [
    _layout_numbered_list,
    _layout_split_panel,
    _layout_card_grid,
    _layout_timeline_row,
    _layout_checklist,
    _layout_two_column,
    _layout_diagonal,
    _layout_icon_grid,
    _layout_quote_highlight,
]


def build_presentation(topic: str, data: dict) -> Path:
    theme = random.choice(PPT_THEMES)
    accent = theme["accent"]

    presentation = Presentation()
    presentation.slide_width = PptxInches(SLIDE_W)
    presentation.slide_height = PptxInches(SLIDE_H)

    blank_layout = presentation.slide_layouts[6]

    title_text = data.get("title") or topic
    slides_data = data.get("slides") or [{"title": topic, "bullets": [topic]}]
    slides_data = _ensure_min_slides(slides_data)
    slides_data = slides_data[:MAX_SLIDES]
    total_slides = len(slides_data) + 2  # + muqova + yopilish

    # --- Muqova slaydi ---
    title_slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(title_slide, theme["background"])
    _add_shape(title_slide, MSO_SHAPE.OVAL, SLIDE_W - 3.5, -1.5, 5.0, 5.0, theme["surface"])
    _add_accent_bar(title_slide, accent, top=3.3)
    _add_text(title_slide, title_text, 0.8, 2.6, 7.4, 1.6, size=34, color=theme["text"], bold=True)
    _add_text(title_slide, f"Mavzu: {topic}", 0.8, 3.65, 7.4, 0.6, size=16, color=accent)
    _add_text(title_slide, "Sun'iy intellekt yordamida tayyorlandi", 0.8, 6.6, 9.3, 0.5, size=12, color=accent)

    # --- Kontent slaydlari (bulletlar soniga qarab turli layout) ---
    for idx, slide_info in enumerate(slides_data):
        slide_title, _ = clean_markdown(str(slide_info.get("title") or f"{idx + 1}-qism"))
        bullets = [clean_markdown(str(b))[0] for b in (slide_info.get("bullets") or []) if str(b).strip()]
        bullets = bullets[:5] or [slide_title]

        if len(bullets) <= 2:
            layout_fn = _layout_big_focus
        else:
            layout_fn = _ROTATING_LAYOUTS[idx % len(_ROTATING_LAYOUTS)]

        layout_fn(presentation, blank_layout, theme, accent, slide_title, bullets, topic, idx + 2, total_slides)

    # --- Yopilish slaydi ---
    closing_slide = presentation.slides.add_slide(blank_layout)
    _set_slide_background(closing_slide, theme["background"])
    _add_corner_visual(closing_slide, theme, accent, corner="br")
    _add_accent_bar(closing_slide, accent, top=3.3)
    _add_text(closing_slide, "Diqqatingiz uchun rahmat!", 0.8, 2.6, 11.8, 1.2, size=32, color=theme["text"], bold=True, align=PP_ALIGN.LEFT)
    _add_footer(closing_slide, accent, topic, total_slides, total_slides)

    path = Path(tempfile.gettempdir()) / f"{safe_filename(topic)}_slayd.pptx"
    presentation.save(path)
    return path


def admin_panel_keyboard() -> InlineKeyboardMarkup:
    mode = get_subscription_mode()
    mode_label = "🔒 Qattiq (admin talab qiladi)" if mode == "strict" else "🔓 Yumshoq (admin talab qilmaydi)"
    builder = InlineKeyboardBuilder()
    builder.add(InlineKeyboardButton(text="📢 Reklama yuborish", callback_data="adm:ads"))
    builder.add(InlineKeyboardButton(text="➕ Kanal qo'shish", callback_data="adm:addch"))
    builder.add(InlineKeyboardButton(text="➖ Kanal o'chirish", callback_data="adm:rmch"))
    builder.add(InlineKeyboardButton(text="📋 Kanallar ro'yxati", callback_data="adm:listch"))
    builder.add(InlineKeyboardButton(text=f"Obuna rejimi: {mode_label}", callback_data="adm:togglemode"))
    builder.adjust(1)
    return builder.as_markup()


@router.message(Command("admin"))
async def admin_handler(message: Message) -> None:
    if not message.from_user or message.from_user.id != ADMIN_ID:
        return

    await message.answer(
        f"📊 Admin panel.\nFoydalanuvchilar: {get_user_count()}\n\nKerakli bo'limni tanlang:",
        reply_markup=admin_panel_keyboard(),
    )


@router.callback_query(F.data.startswith("adm:"))
async def admin_menu_callback(callback: CallbackQuery, state: FSMContext) -> None:
    if not callback.from_user or callback.from_user.id != ADMIN_ID:
        await callback.answer("Bu sizga tegishli emas.", show_alert=True)
        return

    action = callback.data.split(":", 1)[1]

    if action == "ads":
        await callback.answer()
        await state.set_state(WorkState.waiting_for_ad)
        if callback.message:
            await callback.message.answer("Reklama matni, rasm yoki video yuboring.")

    elif action == "addch":
        await callback.answer()
        await state.set_state(WorkState.waiting_for_channel_add)
        mode = get_subscription_mode()
        if callback.message:
            if mode == "strict":
                await callback.message.answer(
                    "Kanalni qo'shish uchun:\n"
                    "• kanaldan istalgan postni shu yerga forward qiling (eng qulay usul), yoki\n"
                    "• kanal username'ini yuboring, masalan: @mening_kanalim\n\n"
                    "Diqqat: hozir rejim \"Qattiq\" — obuna chinakam tekshiriladi, shuning uchun "
                    "botni oldin o'sha kanalga admin qilib qo'shing.\n\n"
                    "Agar botni admin qilishni istamasangiz, avval rejimni \"Yumshoq\"ga o'zgartiring "
                    "(admin paneldagi \"Obuna rejimi\" tugmasi)."
                )
            else:
                await callback.message.answer(
                    "Kanalni qo'shish uchun:\n"
                    "• kanaldan istalgan postni shu yerga forward qiling, yoki\n"
                    "• kanal username'ini yuboring, masalan: @mening_kanalim\n\n"
                    "Hozir rejim \"Yumshoq\" — obuna chinakam tekshirilmaydi, shuning uchun "
                    "botni bu kanalda admin qilish shart emas."
                )

    elif action == "rmch":
        await callback.answer()
        channels = get_channels()
        if not channels:
            if callback.message:
                await callback.message.answer("Hozircha hech qanday majburiy kanal yo'q.")
            return
        builder = InlineKeyboardBuilder()
        for channel_ref, title, _url in channels:
            builder.add(InlineKeyboardButton(text=f"🗑 {title}", callback_data=f"rmch:{channel_ref}"))
        builder.adjust(1)
        if callback.message:
            await callback.message.answer("O'chirish uchun kanalni tanlang:", reply_markup=builder.as_markup())

    elif action == "listch":
        await callback.answer()
        channels = get_channels()
        if not channels:
            text = "Hozircha hech qanday majburiy kanal yo'q."
        else:
            lines = [f"• {title} ({channel_ref})" for channel_ref, title, _url in channels]
            text = "📋 Majburiy kanallar:\n" + "\n".join(lines)
        if callback.message:
            await callback.message.answer(text)

    elif action == "togglemode":
        current = get_subscription_mode()
        new_mode = "soft" if current == "strict" else "strict"
        set_subscription_mode(new_mode)
        if new_mode == "soft":
            await callback.answer("Yumshoq rejim yoqildi — admin talab qilinmaydi", show_alert=True)
        else:
            await callback.answer("Qattiq rejim yoqildi — obuna chinakam tekshiriladi", show_alert=True)
        if callback.message:
            await callback.message.edit_reply_markup(reply_markup=admin_panel_keyboard())


@router.callback_query(F.data.startswith("rmch:"))
async def remove_channel_callback(callback: CallbackQuery) -> None:
    if not callback.from_user or callback.from_user.id != ADMIN_ID:
        await callback.answer("Bu sizga tegishli emas.", show_alert=True)
        return

    channel_ref = callback.data.split(":", 1)[1]
    removed = remove_channel(channel_ref)
    await callback.answer("✅ O'chirildi" if removed else "Topilmadi", show_alert=True)
    if callback.message:
        await callback.message.edit_reply_markup(reply_markup=None)


@router.message(WorkState.waiting_for_channel_add)
async def channel_add_handler(message: Message, state: FSMContext, bot: Bot) -> None:
    if not message.from_user or message.from_user.id != ADMIN_ID:
        await state.clear()
        return

    await state.clear()

    channel_ref: str | None = None
    title: str | None = None
    url: str | None = None

    forward_chat = message.forward_from_chat
    if forward_chat and forward_chat.type == "channel":
        title = forward_chat.title or "Kanal"
        if forward_chat.username:
            channel_ref = f"@{forward_chat.username}"
            url = f"https://t.me/{forward_chat.username}"
        else:
            channel_ref = str(forward_chat.id)
            try:
                invite = await bot.create_chat_invite_link(forward_chat.id)
                url = invite.invite_link
            except Exception:
                await message.answer(
                    "Bu kanalda username yo'q va men taklif havolasi yarata olmadim. "
                    "Botni kanalda \"Foydalanuvchilarni taklif qilish\" huquqi bilan admin qiling va qaytadan urinib ko'ring."
                )
                return
    elif message.text and message.text.strip():
        username = message.text.strip().lstrip("@")
        if not username:
            await message.answer("Noto'g'ri format. @username yuboring yoki kanal postini forward qiling.")
            return
        channel_ref = f"@{username}"
        url = f"https://t.me/{username}"
        title = f"@{username}"
    else:
        await message.answer("Kanal postini forward qiling yoki @username yuboring.")
        return

    mode = get_subscription_mode()

    if mode == "strict":
        try:
            me = await bot.get_me()
            bot_member = await bot.get_chat_member(channel_ref, me.id)
            if bot_member.status not in {ChatMemberStatus.ADMINISTRATOR, ChatMemberStatus.CREATOR}:
                await message.answer(
                    f"⚠️ Men {title} kanalida topildim, lekin admin emasman. "
                    "Iltimos, meni o'sha kanalda admin qiling va qaytadan urinib ko'ring.\n\n"
                    "Yoki admin talab qilinmasligi uchun \"Obuna rejimi\"ni \"Yumshoq\"ga o'zgartiring."
                )
                return
        except Exception as e:
            logger.warning("Channel admin check failed for %s: %s", channel_ref, e)
            await message.answer(
                f"⚠️ {title} kanalini topa olmadim. Kanal username'i to'g'riligini va "
                "meni o'sha kanalga admin qilib qo'shganingizni tekshiring, keyin qaytadan urinib ko'ring.\n\n"
                "Yoki admin talab qilinmasligi uchun \"Obuna rejimi\"ni \"Yumshoq\"ga o'zgartiring."
            )
            return

    add_channel(channel_ref, title, url)
    if mode == "strict":
        await message.answer(
            f"✅ Kanal qo'shildi: {title}\n"
            "Endi yangi foydalanuvchilar botdan foydalanishdan oldin shu kanalga obuna bo'lishlari talab qilinadi."
        )
    else:
        await message.answer(
            f"✅ Kanal qo'shildi: {title}\n"
            "Yumshoq rejimda bo'lgani uchun foydalanuvchilarga faqat obuna bo'lish tugmasi ko'rsatiladi, "
            "lekin bosish majburlanmaydi."
        )


@router.message(Command("start"))
async def start_handler(message: Message, bot: Bot) -> None:
    if not message.from_user:
        return
    add_user(message.from_user.id)
    if not await require_subscription(message, bot):
        return
    await message.answer("Assalomu alaykum. Kerakli xizmatni tanlang:", reply_markup=main_keyboard())


@router.callback_query(F.data == "check_subscription")
async def subscription_callback(callback: CallbackQuery, bot: Bot) -> None:
    missing = await get_missing_channels(bot, callback.from_user.id)
    if not missing:
        await callback.answer("Obuna tasdiqlandi!")
        if callback.message:
            await callback.message.edit_text("Obuna tasdiqlandi. Endi botdan foydalanishingiz mumkin.")
            await callback.message.answer("Xizmatni tanlang:", reply_markup=main_keyboard())
    else:
        await callback.answer("Hali barcha kanallarga obuna bo‘lmagansiz.", show_alert=True)
        # Endi faqat obuna bo'lib ulgurmagan kanal(lar)ni ko'rsatamiz —
        # allaqachon obuna bo'lgan kanallar tugmalarini olib tashlaymiz.
        if callback.message:
            try:
                await callback.message.edit_reply_markup(reply_markup=subscription_keyboard(missing))
            except TelegramBadRequest:
                # Xabar matni/tugmalari o'zgarmagan bo'lsa Telegram xato qaytaradi — e'tiborsiz qoldiramiz.
                pass


@router.message(F.text.in_(BUTTON_TO_SERVICE.keys()))
async def service_handler(message: Message, state: FSMContext, bot: Bot) -> None:
    if not message.from_user or not message.text:
        return
    if not await require_subscription(message, bot):
        return

    service = BUTTON_TO_SERVICE[message.text]
    user_id = message.from_user.id

    # Limitni MAVZU SO'RALISHIDAN OLDIN tekshiramiz — foydalanuvchi 24
    # soat ichida shu xizmatdan qayta foydalanolmasa, mavzuni yozishga
    # umuman vaqt sarflamasin va darhol qancha kutish kerakligini bilsin.
    if user_id != ADMIN_ID:
        remaining = get_remaining_cooldown(user_id, service)
        if remaining is not None:
            await message.answer(
                f"Bu xizmatdan qayta foydalanish uchun {format_remaining(remaining)}dan keyin urinib ko'ring."
            )
            return

    await state.set_state(WorkState.waiting_for_topic)
    await state.update_data(service=service)

    await message.answer(f"{SERVICE_NAMES[service]} uchun mavzuni yuboring:")


@router.message(WorkState.waiting_for_topic, F.text)
async def topic_handler(message: Message, state: FSMContext) -> None:
    if not message.text or not message.from_user:
        return

    data = await state.get_data()
    service = data.get("service", "slayd")
    topic = message.text.strip()
    user_id = message.from_user.id

    if len(topic) < 3:
        await message.answer("Mavzu kamida 3 ta belgidan iborat bo‘lsin.")
        return

    if user_id == ADMIN_ID:
        allowed, remaining = True, 0.0
    else:
        allowed, remaining = check_and_update_limit(user_id, service)
    if not allowed:
        await message.answer(f"Bu xizmatdan qayta foydalanish uchun {format_remaining(remaining)}dan keyin urinib ko'ring.")
        await state.clear()
        return

    await state.clear()
    status_msg = await message.answer("⏳ Material tayyorlanmoqda. Iltimos, kuting...")

    path: Path | None = None
    try:
        if service == "slayd":
            slides_data = await asyncio.to_thread(generate_slides_data, topic)
            path = await asyncio.to_thread(build_presentation, topic, slides_data)
        else:
            doc_data = await asyncio.to_thread(generate_document_data, service, topic)
            path = await asyncio.to_thread(build_document, service, topic, doc_data)

        await status_msg.delete()
        await message.answer_document(
            FSInputFile(path),
            caption=f"✨ {SERVICE_NAMES.get(service)} tayyor bo'ldi!",
        )

    except Exception as e:
        logger.exception("Generation failed: %s", e)
        release_limit(user_id, service)
        await status_msg.delete()
        await message.answer(
            "Kechirasiz, materialni yaratishda xatolik yuz berdi. "
            "Urinishingiz hisoblanmadi — qaytadan urinib ko'rishingiz mumkin."
        )
    finally:
        if path is not None:
            path.unlink(missing_ok=True)


@router.message(WorkState.waiting_for_topic)
async def topic_invalid_handler(message: Message) -> None:
    await message.answer("Iltimos, mavzuni oddiy matn ko'rinishida yuboring.")


@router.message(WorkState.waiting_for_ad)
async def ad_handler(message: Message, state: FSMContext, bot: Bot) -> None:
    if not message.from_user or message.from_user.id != ADMIN_ID:
        await state.clear()
        return

    await state.clear()
    users = get_all_users()
    sent = 0
    failed = 0
    await message.answer("🚀 Reklama tarqatilmoqda...")

    for uid in users:
        try:
            await message.copy_to(chat_id=uid)
            sent += 1
        except TelegramRetryAfter as exc:
            await asyncio.sleep(exc.retry_after)
            try:
                await message.copy_to(chat_id=uid)
                sent += 1
            except Exception:
                failed += 1
        except Exception:
            failed += 1
        await asyncio.sleep(0.05)

    await message.answer(
        f"✅ Reklama {sent} ta foydalanuvchiga muvaffaqiyatli yuborildi."
        + (f"\n⚠️ {failed} ta foydalanuvchiga yetkazilmadi (bot bloklangan bo'lishi mumkin)." if failed else "")
    )


async def main() -> None:
    if not BOT_TOKEN or not GEMINI_API_KEY or not ADMIN_ID:
        raise RuntimeError("BOT_TOKEN, GEMINI_API_KEY va ADMIN_ID .env faylida sozlanishi kerak")

    init_db()
    migrate_env_channels_if_needed()

    # Slayd yaratish (Gemini so'rovi + rasm generatsiyasi + pptx qurish) va
    # hujjat yaratish (Gemini so'rovi + docx qurish) ikkalasi ham
    # asyncio.to_thread orqali umumiy thread pool'da ishlaydi. Standart
    # pool juda kichik bo'lishi mumkin (ayniqsa kam yadroli serverda),
    # shuning uchun uni kengaytiramiz — aks holda bir nechta foydalanuvchi
    # bir vaqtda slayd so'rasa, boshqalarning hujjat so'rovlari navbatda
    # uzoq kutib qoladi (yoki "jonatilmayapti" kabi ko'rinadi).
    executor = ThreadPoolExecutor(max_workers=16)
    asyncio.get_running_loop().set_default_executor(executor)

    session = AiohttpSession(timeout=60)
    bot = Bot(BOT_TOKEN, session=session)

    dispatcher = Dispatcher(storage=MemoryStorage())
    dispatcher.include_router(router)

    try:
        await dispatcher.start_polling(bot, allowed_updates=dispatcher.resolve_used_update_types())
    finally:
        await bot.session.close()


if __name__ == "__main__":
    asyncio.run(main())