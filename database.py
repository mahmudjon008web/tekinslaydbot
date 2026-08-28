from __future__ import annotations

import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

DB_PATH = Path(__file__).with_name("bot_limits.db")


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _connect() -> sqlite3.Connection:
    connection = sqlite3.connect(
        DB_PATH,
        timeout=30,
    )
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA foreign_keys=ON")
    return connection


def init_db() -> None:
    connection = _connect()
    try:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                joined_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS user_limits (
                user_id INTEGER NOT NULL,
                service_type TEXT NOT NULL,
                last_used TEXT NOT NULL,
                PRIMARY KEY (user_id, service_type),
                FOREIGN KEY (user_id)
                    REFERENCES users(user_id)
                    ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS channels (
                channel_ref TEXT PRIMARY KEY,
                title TEXT NOT NULL,
                url TEXT NOT NULL,
                added_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )
        connection.commit()
    finally:
        connection.close()


def add_user(user_id: int) -> None:
    connection = _connect()
    try:
        connection.execute(
            """
            INSERT OR IGNORE INTO users (
                user_id,
                joined_at
            )
            VALUES (?, ?)
            """,
            (
                user_id,
                _now().isoformat(),
            ),
        )
        connection.commit()
    finally:
        connection.close()


def get_all_users() -> list[int]:
    connection = _connect()
    try:
        rows = connection.execute(
            """
            SELECT user_id
            FROM users
            ORDER BY user_id
            """
        ).fetchall()
    finally:
        connection.close()

    return [int(row["user_id"]) for row in rows]


def get_user_count() -> int:
    connection = _connect()
    try:
        row = connection.execute(
            "SELECT COUNT(*) AS count FROM users"
        ).fetchone()
    finally:
        connection.close()

    return int(row["count"])


def get_remaining_cooldown(user_id: int, service_type: str) -> float | None:
    """Non-mutating check: returns None if the user is allowed to use this
    service right now, or the remaining cooldown in hours if not.

    Unlike check_and_update_limit(), this does NOT record a new usage —
    it's meant to be called as an early, cheap check (e.g. right when the
    user taps a service button) BEFORE we bother asking them for a topic.
    """
    now = _now()
    cooldown = timedelta(hours=24)

    connection = _connect()
    try:
        row = connection.execute(
            """
            SELECT last_used
            FROM user_limits
            WHERE user_id = ?
              AND service_type = ?
            """,
            (
                user_id,
                service_type,
            ),
        ).fetchone()
    finally:
        connection.close()

    if not row:
        return None

    last_used = datetime.fromisoformat(row["last_used"])
    elapsed = now - last_used

    if elapsed < cooldown:
        remaining = (cooldown - elapsed).total_seconds() / 3600
        return max(0.1, remaining)

    return None


def check_and_update_limit(
    user_id: int,
    service_type: str,
) -> tuple[bool, float]:
    """Check the cooldown and, if allowed, immediately record usage.

    NOTE: if the caller later fails to deliver the result (e.g. the AI
    generation or file creation raises), it should call release_limit()
    with the same arguments so the user isn't charged for a failed attempt.
    """
    now = _now()
    cooldown = timedelta(hours=24)

    add_user(user_id)

    connection = _connect()
    try:
        connection.execute("BEGIN IMMEDIATE")

        row = connection.execute(
            """
            SELECT last_used
            FROM user_limits
            WHERE user_id = ?
              AND service_type = ?
            """,
            (
                user_id,
                service_type,
            ),
        ).fetchone()

        if row:
            last_used = datetime.fromisoformat(
                row["last_used"]
            )
            elapsed = now - last_used

            if elapsed < cooldown:
                remaining = (
                    cooldown - elapsed
                ).total_seconds() / 3600

                connection.rollback()

                return False, max(0.1, remaining)

        connection.execute(
            """
            INSERT INTO user_limits (
                user_id,
                service_type,
                last_used
            )
            VALUES (?, ?, ?)
            ON CONFLICT(user_id, service_type)
            DO UPDATE SET last_used = excluded.last_used
            """,
            (
                user_id,
                service_type,
                now.isoformat(),
            ),
        )

        connection.commit()
    finally:
        connection.close()

    return True, 0.0


def release_limit(user_id: int, service_type: str) -> None:
    """Undo a limit charge made by check_and_update_limit.

    Call this when generation/delivery failed after the limit was already
    recorded, so the user isn't punished with a 24h cooldown for nothing.
    """
    connection = _connect()
    try:
        connection.execute(
            """
            DELETE FROM user_limits
            WHERE user_id = ?
              AND service_type = ?
            """,
            (
                user_id,
                service_type,
            ),
        )
        connection.commit()
    finally:
        connection.close()


def add_channel(channel_ref: str, title: str, url: str) -> None:
    """Add or update a required subscription channel."""
    connection = _connect()
    try:
        connection.execute(
            """
            INSERT INTO channels (channel_ref, title, url, added_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(channel_ref)
            DO UPDATE SET title = excluded.title, url = excluded.url
            """,
            (
                channel_ref,
                title,
                url,
                _now().isoformat(),
            ),
        )
        connection.commit()
    finally:
        connection.close()


def remove_channel(channel_ref: str) -> bool:
    """Remove a required subscription channel. Returns True if it existed."""
    connection = _connect()
    try:
        cursor = connection.execute(
            "DELETE FROM channels WHERE channel_ref = ?",
            (channel_ref,),
        )
        connection.commit()
        return cursor.rowcount > 0
    finally:
        connection.close()


def get_channels() -> list[tuple[str, str, str]]:
    """Returns list of (channel_ref, title, url) tuples."""
    connection = _connect()
    try:
        rows = connection.execute(
            "SELECT channel_ref, title, url FROM channels ORDER BY added_at"
        ).fetchall()
    finally:
        connection.close()

    return [(row["channel_ref"], row["title"], row["url"]) for row in rows]


# ---------------------------------------------------------------------------
# Settings — hozircha faqat obuna tekshirish rejimi uchun ishlatiladi.
#
#   "strict" — bot get_chat_member orqali obunani chinakam tekshiradi.
#              Buning uchun bot HAR BIR kanalda ADMIN bo'lishi shart
#              (bu Telegram API'ning o'z cheklovi, botlar kanalga faqat
#              admin sifatida qo'shiladi — buni kod bilan aylanib o'tib
#              bo'lmaydi).
#   "soft"   — foydalanuvchiga "obuna bo'ling" tugmalari ko'rsatiladi,
#              lekin haqiqiy tekshiruv qilinmaydi (ishonchga asoslangan).
#              Bu rejimda botni HECH QAYERDA admin qilish shart emas.
# ---------------------------------------------------------------------------

_DEFAULT_SUBSCRIPTION_MODE = "strict"


def get_subscription_mode() -> str:
    connection = _connect()
    try:
        row = connection.execute(
            "SELECT value FROM settings WHERE key = 'subscription_mode'"
        ).fetchone()
    finally:
        connection.close()

    if row and row["value"] in {"strict", "soft"}:
        return row["value"]
    return _DEFAULT_SUBSCRIPTION_MODE


def set_subscription_mode(mode: str) -> None:
    if mode not in {"strict", "soft"}:
        raise ValueError("mode faqat 'strict' yoki 'soft' bo'lishi mumkin")

    connection = _connect()
    try:
        connection.execute(
            """
            INSERT INTO settings (key, value)
            VALUES ('subscription_mode', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (mode,),
        )
        connection.commit()
    finally:
        connection.close()