from __future__ import annotations

import os

import psycopg
from psycopg import sql
from dotenv import load_dotenv

from rentability.repository import PostgresRepository


TARGET_DB_NAME = os.getenv("RENTABILITY_DB_NAME", "rentability_db")


def build_target_dsn() -> str:
    return os.getenv(
        "DATABASE_URL",
        f"postgresql://postgres:postgres@localhost:5432/{TARGET_DB_NAME}",
    )


def build_admin_dsn() -> str:
    return os.getenv(
        "POSTGRES_ADMIN_URL",
        "postgresql://postgres:postgres@localhost:5432/postgres",
    )


def ensure_database() -> None:
    admin_dsn = build_admin_dsn()
    with psycopg.connect(admin_dsn, autocommit=True) as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM pg_database WHERE datname = %s", (TARGET_DB_NAME,))
            exists = cur.fetchone() is not None
            if not exists:
                cur.execute(sql.SQL("CREATE DATABASE {}").format(sql.Identifier(TARGET_DB_NAME)))


def main() -> None:
    load_dotenv()
    ensure_database()
    repository = PostgresRepository(build_target_dsn())
    repository.initialize()
    print(f"Database '{TARGET_DB_NAME}' is ready.")
    print("Default enterprises are present:")
    for enterprise in repository.list_enterprises():
        print(f"- {enterprise.name}")


if __name__ == "__main__":
    main()
