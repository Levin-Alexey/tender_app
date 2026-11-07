from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import asyncpg


def normalize_database_url(database_url: str) -> str:
    database_url = database_url.strip()
    if database_url.startswith("postgresql+asyncpg://"):
        database_url = "postgresql://" + database_url.removeprefix(
            "postgresql+asyncpg://"
        )
    if not database_url:
        raise RuntimeError("DATABASE_URL is not set")
    if not database_url.startswith(("postgresql://", "postgres://")):
        raise RuntimeError(
            "DATABASE_URL must use the postgresql:// or postgres:// scheme"
        )
    return database_url


async def create_database_pool() -> asyncpg.Pool:
    import asyncpg

    database_url = normalize_database_url(os.getenv("DATABASE_URL", ""))
    pool = await asyncpg.create_pool(database_url, min_size=1, max_size=10)
    try:
        schema = Path(__file__).with_name("schema.sql").read_text(encoding="utf-8")
        async with pool.acquire() as connection:
            async with connection.transaction():
                await connection.execute("SELECT pg_advisory_xact_lock(728194321)")
                await connection.execute(schema)
    except Exception:
        await pool.close()
        raise
    return pool
