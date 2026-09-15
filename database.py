from __future__ import annotations

import os
from datetime import datetime, timedelta, timezone
import asyncpg
from dotenv import load_dotenv

load_dotenv()  # .env faylini shu yerning o'zida ham yuklaymiz — import tartibiga bog'liq bo'lmasin


def _now() -> datetime:
    return datetime.now(timezone.utc)


async def _get_connection() -> asyncpg.Connection:
    database_url = os.getenv("DATABASE_URL")  # HAR SAFAR yangidan o'qiladi, module yuklanganda emas
    if not database_url:
        raise ValueError("DATABASE_URL muhit o'zgaruvchisi topilmadi!")

    url = database_url
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)

    return await asyncpg.connect(url, ssl="require")


async def init_db() -> None:
    conn = await _get_connection()
    try:
        await conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id BIGINT PRIMARY KEY,
                joined_at TIMESTAMPTZ NOT NULL
            );

            CREATE TABLE IF NOT EXISTS user_limits (
                user_id BIGINT NOT NULL,
                service_type TEXT NOT NULL,
                last_used TIMESTAMPTZ NOT NULL,
                PRIMARY KEY (user_id, service_type),
                FOREIGN KEY (user_id)
                    REFERENCES users(user_id)
                    ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS channels (
                channel_ref TEXT PRIMARY KEY,
                title TEXT NOT NULL,
                url TEXT NOT NULL,
                added_at TIMESTAMPTZ NOT NULL
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )
    finally:
        await conn.close()


async def add_user(user_id: int) -> None:
    conn = await _get_connection()
    try:
        await conn.execute(
            """
            INSERT INTO users (user_id, joined_at)
            VALUES ($1, $2)
            ON CONFLICT (user_id) DO NOTHING
            """,
            user_id,
            _now(),
        )
    finally:
        await conn.close()


async def get_all_users() -> list[int]:
    conn = await _get_connection()
    try:
        rows = await conn.fetch(
            """
            SELECT user_id
            FROM users
            ORDER BY user_id
            """
        )
    finally:
        await conn.close()

    return [int(row["user_id"]) for row in rows]


async def get_user_count() -> int:
    conn = await _get_connection()
    try:
        count = await conn.fetchval("SELECT COUNT(*) FROM users")
    finally:
        await conn.close()

    return int(count) if count is not None else 0


async def get_remaining_cooldown(user_id: int, service_type: str) -> float | None:
    now = _now()
    cooldown = timedelta(hours=24)

    conn = await _get_connection()
    try:
        last_used = await conn.fetchval(
            """
            SELECT last_used
            FROM user_limits
            WHERE user_id = $1
              AND service_type = $2
            """,
            user_id,
            service_type,
        )
    finally:
        await conn.close()

    if not last_used:
        return None

    elapsed = now - last_used

    if elapsed < cooldown:
        remaining = (cooldown - elapsed).total_seconds() / 3600
        return max(0.1, remaining)

    return None


async def check_and_update_limit(
    user_id: int,
    service_type: str,
) -> tuple[bool, float]:
    now = _now()
    cooldown = timedelta(hours=24)

    await add_user(user_id)

    conn = await _get_connection()
    can_proceed = True
    remaining_time = 0.0

    try:
        async with conn.transaction():
            last_used = await conn.fetchval(
                """
                SELECT last_used
                FROM user_limits
                WHERE user_id = $1
                  AND service_type = $2
                FOR UPDATE
                """,
                user_id,
                service_type,
            )

            if last_used:
                elapsed = now - last_used
                if elapsed < cooldown:
                    remaining = (cooldown - elapsed).total_seconds() / 3600
                    can_proceed = False
                    remaining_time = max(0.1, remaining)

            if can_proceed:
                await conn.execute(
                    """
                    INSERT INTO user_limits (user_id, service_type, last_used)
                    VALUES ($1, $2, $3)
                    ON CONFLICT (user_id, service_type)
                    DO UPDATE SET last_used = EXCLUDED.last_used
                    """,
                    user_id,
                    service_type,
                    now,
                )
    finally:
        await conn.close()

    return can_proceed, remaining_time


async def release_limit(user_id: int, service_type: str) -> None:
    conn = await _get_connection()
    try:
        await conn.execute(
            """
            DELETE FROM user_limits
            WHERE user_id = $1
              AND service_type = $2
            """,
            user_id,
            service_type,
        )
    finally:
        await conn.close()


async def add_channel(channel_ref: str, title: str, url: str) -> None:
    conn = await _get_connection()
    try:
        await conn.execute(
            """
            INSERT INTO channels (channel_ref, title, url, added_at)
            VALUES ($1, $2, $3, $4)
            ON CONFLICT (channel_ref)
            DO UPDATE SET title = EXCLUDED.title, url = EXCLUDED.url
            """,
            channel_ref,
            title,
            url,
            _now(),
        )
    finally:
        await conn.close()


async def remove_channel(channel_ref: str) -> bool:
    conn = await _get_connection()
    try:
        result = await conn.execute(
            "DELETE FROM channels WHERE channel_ref = $1",
            channel_ref,
        )
        rows_affected = int(result.split(" ")[1])
        return rows_affected > 0
    finally:
        await conn.close()


async def get_channels() -> list[tuple[str, str, str]]:
    conn = await _get_connection()
    try:
        rows = await conn.fetch(
            "SELECT channel_ref, title, url FROM channels ORDER BY added_at"
        )
    finally:
        await conn.close()

    return [(row["channel_ref"], row["title"], row["url"]) for row in rows]


_DEFAULT_SUBSCRIPTION_MODE = "strict"


async def get_subscription_mode() -> str:
    conn = await _get_connection()
    try:
        val = await conn.fetchval(
            "SELECT value FROM settings WHERE key = 'subscription_mode'"
        )
    finally:
        await conn.close()

    if val and val in {"strict", "soft"}:
        return val
    return _DEFAULT_SUBSCRIPTION_MODE


async def set_subscription_mode(mode: str) -> None:
    if mode not in {"strict", "soft"}:
        raise ValueError("mode faqat 'strict' yoki 'soft' bo'lishi mumkin")

    conn = await _get_connection()
    try:
        await conn.execute(
            """
            INSERT INTO settings (key, value)
            VALUES ('subscription_mode', $1)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
            """,
            mode,
        )
    finally:
        await conn.close()