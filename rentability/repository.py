from __future__ import annotations

import os
from contextlib import contextmanager
from datetime import date, datetime

import psycopg

from .default_data import DEFAULT_ENTERPRISES
from .models import Enterprise, FinancialRecord


class PostgresRepository:
    def __init__(self, dsn: str | None = None):
        self.dsn = dsn or os.getenv(
            "DATABASE_URL",
            "dbname=rentability_db user=postgres password=postgres host=localhost port=5432",
        )

    @contextmanager
    def connection(self):
        with psycopg.connect(self.dsn) as conn:
            yield conn

    def initialize(self) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS enterprises (
                        id SERIAL PRIMARY KEY,
                        name TEXT NOT NULL UNIQUE
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS enterprise_metrics (
                        id SERIAL PRIMARY KEY,
                        enterprise_id INTEGER NOT NULL REFERENCES enterprises(id) ON DELETE CASCADE,
                        period_date DATE NOT NULL,
                        revenue NUMERIC(14, 2) NOT NULL,
                        cost NUMERIC(14, 2) NOT NULL,
                        fixed_expenses NUMERIC(14, 2) NOT NULL,
                        variable_expenses NUMERIC(14, 2) NOT NULL,
                        tax NUMERIC(14, 2) NOT NULL,
                        UNIQUE (enterprise_id, period_date)
                    );
                    """
                )
            conn.commit()
        self.ensure_default_data()

    def ensure_default_data(self) -> None:
        for enterprise_name, rows in DEFAULT_ENTERPRISES.items():
            enterprise = self.get_enterprise_by_name(enterprise_name)
            if enterprise is None:
                enterprise = self.add_enterprise(enterprise_name)
            for row in rows:
                period_date = datetime.strptime(row["date"], "%Y-%m-%d").date()
                if self.get_record_by_date(enterprise.id, period_date) is None:
                    self.add_record(
                        FinancialRecord(
                            id=None,
                            enterprise_id=enterprise.id,
                            period_date=period_date,
                            revenue=float(row["revenue"]),
                            cost=float(row["cost"]),
                            fixed_expenses=float(row["fixed_expenses"]),
                            variable_expenses=float(row["variable_expenses"]),
                            tax=float(row["tax"]),
                        )
                    )

    def list_enterprises(self) -> list[Enterprise]:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT id, name FROM enterprises ORDER BY name")
                return [Enterprise(id=row[0], name=row[1]) for row in cur.fetchall()]

    def add_enterprise(self, name: str) -> Enterprise:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO enterprises (name)
                    VALUES (%s)
                    RETURNING id, name
                    """,
                    (name,),
                )
                row = cur.fetchone()
            conn.commit()
        return Enterprise(id=row[0], name=row[1])

    def update_enterprise_name(self, enterprise_id: int, name: str) -> Enterprise:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE enterprises
                    SET name = %s
                    WHERE id = %s
                    RETURNING id, name
                    """,
                    (name, enterprise_id),
                )
                row = cur.fetchone()
            conn.commit()
        return Enterprise(id=row[0], name=row[1])

    def get_enterprise_by_name(self, name: str) -> Enterprise | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT id, name FROM enterprises WHERE name = %s", (name,))
                row = cur.fetchone()
        return Enterprise(id=row[0], name=row[1]) if row else None

    def get_records(
        self,
        enterprise_id: int,
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[FinancialRecord]:
        params: list[object] = [enterprise_id]
        conditions = ["enterprise_id = %s"]
        if start_date:
            params.append(start_date)
            conditions.append("period_date >= %s")
        if end_date:
            params.append(end_date)
            conditions.append("period_date <= %s")

        query = f"""
            SELECT id, enterprise_id, period_date, revenue, cost, fixed_expenses, variable_expenses, tax
            FROM enterprise_metrics
            WHERE {' AND '.join(conditions)}
            ORDER BY period_date
        """

        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(query, params)
                rows = cur.fetchall()

        return [
            FinancialRecord(
                id=row[0],
                enterprise_id=row[1],
                period_date=row[2],
                revenue=float(row[3]),
                cost=float(row[4]),
                fixed_expenses=float(row[5]),
                variable_expenses=float(row[6]),
                tax=float(row[7]),
            )
            for row in rows
        ]

    def get_record_by_date(self, enterprise_id: int, period_date: date) -> FinancialRecord | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT id, enterprise_id, period_date, revenue, cost, fixed_expenses, variable_expenses, tax
                    FROM enterprise_metrics
                    WHERE enterprise_id = %s AND period_date = %s
                    """,
                    (enterprise_id, period_date),
                )
                row = cur.fetchone()

        if not row:
            return None

        return FinancialRecord(
            id=row[0],
            enterprise_id=row[1],
            period_date=row[2],
            revenue=float(row[3]),
            cost=float(row[4]),
            fixed_expenses=float(row[5]),
            variable_expenses=float(row[6]),
            tax=float(row[7]),
        )

    def get_latest_record_date(self, enterprise_id: int) -> date | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT MAX(period_date) FROM enterprise_metrics WHERE enterprise_id = %s",
                    (enterprise_id,),
                )
                row = cur.fetchone()
        return row[0] if row and row[0] else None

    def get_record_date_bounds(self, enterprise_id: int) -> tuple[date | None, date | None]:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT MIN(period_date), MAX(period_date)
                    FROM enterprise_metrics
                    WHERE enterprise_id = %s
                    """,
                    (enterprise_id,),
                )
                row = cur.fetchone()
        if not row:
            return None, None
        return row[0], row[1]

    def add_record(self, record: FinancialRecord) -> FinancialRecord:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO enterprise_metrics (
                        enterprise_id, period_date, revenue, cost, fixed_expenses, variable_expenses, tax
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    RETURNING id
                    """,
                    (
                        record.enterprise_id,
                        record.period_date,
                        record.revenue,
                        record.cost,
                        record.fixed_expenses,
                        record.variable_expenses,
                        record.tax,
                    ),
                )
                record.id = cur.fetchone()[0]
            conn.commit()
        return record

    def update_record(self, record: FinancialRecord) -> None:
        if record.id is None:
            raise ValueError("Для обновления записи требуется id")

        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE enterprise_metrics
                    SET period_date = %s,
                        revenue = %s,
                        cost = %s,
                        fixed_expenses = %s,
                        variable_expenses = %s,
                        tax = %s
                    WHERE id = %s
                    """,
                    (
                        record.period_date,
                        record.revenue,
                        record.cost,
                        record.fixed_expenses,
                        record.variable_expenses,
                        record.tax,
                        record.id,
                    ),
                )
            conn.commit()

    def upsert_record(self, record: FinancialRecord) -> FinancialRecord:
        existing = self.get_record_by_date(record.enterprise_id, record.period_date)
        if existing is None:
            return self.add_record(record)

        existing.revenue = record.revenue
        existing.cost = record.cost
        existing.fixed_expenses = record.fixed_expenses
        existing.variable_expenses = record.variable_expenses
        existing.tax = record.tax
        self.update_record(existing)
        return existing
