from __future__ import annotations

import json
import os
import sqlite3


# ---------------------------------------------------------------------------
# Connection helpers — PostgreSQL when DATABASE_URL is set, SQLite otherwise
# ---------------------------------------------------------------------------

_SQLITE_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "coverage.db")


def _get_pg_conn():
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        return None
    import psycopg2
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    return psycopg2.connect(url)


def _use_postgres() -> bool:
    return bool(os.environ.get("DATABASE_URL", ""))


def _sqlite_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(_SQLITE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def init_db() -> None:
    if _use_postgres():
        conn = _get_pg_conn()
        with conn, conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS issuer_coverage (
                    ticker TEXT PRIMARY KEY,
                    primary_names TEXT NOT NULL DEFAULT '[]',
                    secondary_names TEXT NOT NULL DEFAULT '[]',
                    updated_at TIMESTAMP DEFAULT NOW()
                )
            """)
        conn.close()
    else:
        with _sqlite_conn() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS issuer_coverage (
                    ticker TEXT PRIMARY KEY,
                    primary_names TEXT NOT NULL DEFAULT '[]',
                    secondary_names TEXT NOT NULL DEFAULT '[]',
                    updated_at TEXT DEFAULT (datetime('now'))
                )
            """)


def get_all_coverage() -> dict:
    if _use_postgres():
        conn = _get_pg_conn()
        try:
            with conn, conn.cursor() as cur:
                cur.execute("SELECT ticker, primary_names, secondary_names FROM issuer_coverage")
                rows = cur.fetchall()
            return {
                row[0]: {
                    "primary": json.loads(row[1]),
                    "secondary": json.loads(row[2]),
                }
                for row in rows
            }
        finally:
            conn.close()
    else:
        with _sqlite_conn() as conn:
            rows = conn.execute(
                "SELECT ticker, primary_names, secondary_names FROM issuer_coverage"
            ).fetchall()
        return {
            row["ticker"]: {
                "primary": json.loads(row["primary_names"]),
                "secondary": json.loads(row["secondary_names"]),
            }
            for row in rows
        }


def save_coverage(ticker: str, primary: list, secondary: list) -> None:
    if _use_postgres():
        conn = _get_pg_conn()
        try:
            with conn, conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO issuer_coverage (ticker, primary_names, secondary_names, updated_at)
                    VALUES (%s, %s, %s, NOW())
                    ON CONFLICT (ticker) DO UPDATE SET
                        primary_names = EXCLUDED.primary_names,
                        secondary_names = EXCLUDED.secondary_names,
                        updated_at = NOW()
                    """,
                    (ticker, json.dumps(primary), json.dumps(secondary)),
                )
        finally:
            conn.close()
    else:
        with _sqlite_conn() as conn:
            conn.execute(
                """
                INSERT INTO issuer_coverage (ticker, primary_names, secondary_names, updated_at)
                VALUES (?, ?, ?, datetime('now'))
                ON CONFLICT (ticker) DO UPDATE SET
                    primary_names = excluded.primary_names,
                    secondary_names = excluded.secondary_names,
                    updated_at = datetime('now')
                """,
                (ticker, json.dumps(primary), json.dumps(secondary)),
            )
