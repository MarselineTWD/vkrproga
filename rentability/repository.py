from __future__ import annotations

import os
from contextlib import contextmanager
from datetime import date, datetime

import psycopg

from .default_data import DEFAULT_ENTERPRISES
from .models import Enterprise, FinancialRecord, FinancialReport


ENTERPRISE_METRIC_DEFINITIONS = (
    ("revenue", "Выручка", "₽"),
    ("cost", "Себестоимость", "₽"),
    ("fixed_expenses", "Постоянные издержки", "₽"),
    ("variable_expenses", "Переменные издержки", "₽"),
    ("tax", "Налог", "₽"),
)

REPORT_METRIC_DEFINITIONS = (
    ("target_ros", "Целевой уровень ROS", "%"),
    ("alpha", "Уровень значимости alpha", "ед."),
    ("avg_ros", "Средний ROS", "%"),
    ("std_ros", "Стандартное отклонение ROS", "%"),
    ("min_ros", "Минимальный ROS", "%"),
    ("max_ros", "Максимальный ROS", "%"),
    ("avg_profit", "Средняя чистая прибыль", "₽"),
    ("t_stat", "t-статистика", "ед."),
    ("p_value", "p-уровень", "ед."),
)

UNIT_DEFINITIONS = (
    ("Рубль", "₽"),
    ("Процент", "%"),
    ("Коэффициент", "ед."),
)


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
                    CREATE TABLE IF NOT EXISTS unit (
                        unit_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        unit_name VARCHAR(255) NOT NULL UNIQUE,
                        unit_small_name VARCHAR(50) NOT NULL UNIQUE
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS enterprise_metric (
                        metric_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        metric_name VARCHAR(255) NOT NULL UNIQUE,
                        metric_small_name VARCHAR(50) NOT NULL UNIQUE
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS enterprise (
                        enterprise_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        enterprise_name VARCHAR(255) NOT NULL UNIQUE,
                        enterprise_small_name VARCHAR(50)
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS report_metric (
                        metric_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        metric_name VARCHAR(255) NOT NULL UNIQUE,
                        metric_small_name VARCHAR(50) NOT NULL UNIQUE
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS financial_report (
                        report_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        enterprise_id INTEGER NOT NULL REFERENCES enterprise(enterprise_id) ON DELETE CASCADE,
                        report_name VARCHAR(255) NOT NULL,
                        date_created DATE NOT NULL,
                        period_start DATE NOT NULL,
                        period_end DATE NOT NULL,
                        CONSTRAINT chk_financial_report_period CHECK (period_start <= period_end),
                        CONSTRAINT chk_financial_report_date_created CHECK (date_created <= CURRENT_DATE),
                        CONSTRAINT chk_financial_report_period_end CHECK (period_end <= CURRENT_DATE)
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS enterprise_metric_value (
                        value_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        enterprise_id INTEGER NOT NULL REFERENCES enterprise(enterprise_id) ON DELETE CASCADE,
                        metric_id INTEGER NOT NULL REFERENCES enterprise_metric(metric_id) ON DELETE RESTRICT,
                        unit_id INTEGER NOT NULL REFERENCES unit(unit_id) ON DELETE RESTRICT,
                        value_date DATE NOT NULL,
                        value NUMERIC(18, 6) NOT NULL,
                        CONSTRAINT chk_enterprise_metric_value_date CHECK (value_date <= CURRENT_DATE),
                        UNIQUE (enterprise_id, metric_id, value_date)
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS report_metric_value (
                        value_id INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                        report_id INTEGER NOT NULL REFERENCES financial_report(report_id) ON DELETE CASCADE,
                        metric_id INTEGER NOT NULL REFERENCES report_metric(metric_id) ON DELETE RESTRICT,
                        unit_id INTEGER NOT NULL REFERENCES unit(unit_id) ON DELETE RESTRICT,
                        value NUMERIC(18, 6) NOT NULL,
                        UNIQUE (report_id, metric_id)
                    );
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_enterprise_metric_value_lookup
                    ON enterprise_metric_value (enterprise_id, value_date, metric_id);
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_financial_report_enterprise
                    ON financial_report (enterprise_id, date_created DESC);
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_financial_report_enterprise_period
                    ON financial_report (enterprise_id, period_start, period_end);
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_financial_report_name
                    ON financial_report (report_name);
                    """
                )
            conn.commit()

        self.ensure_database_programmability()
        self.ensure_reference_data()
        self.migrate_legacy_schema()
        self.ensure_default_data()

    def ensure_database_programmability(self) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE OR REPLACE FUNCTION trg_validate_financial_report_dates()
                    RETURNS TRIGGER
                    LANGUAGE plpgsql
                    AS $$
                    BEGIN
                        IF NEW.date_created > CURRENT_DATE THEN
                            RAISE EXCEPTION 'Дата формирования отчёта не может быть в будущем';
                        END IF;
                        IF NEW.period_start > NEW.period_end THEN
                            RAISE EXCEPTION 'Начало периода не может быть позже конца периода';
                        END IF;
                        IF NEW.period_end > CURRENT_DATE THEN
                            RAISE EXCEPTION 'Конец периода отчёта не может быть в будущем';
                        END IF;
                        RETURN NEW;
                    END;
                    $$;
                    """
                )
                cur.execute("DROP TRIGGER IF EXISTS trg_financial_report_validate_dates ON financial_report;")
                cur.execute(
                    """
                    CREATE TRIGGER trg_financial_report_validate_dates
                    BEFORE INSERT OR UPDATE ON financial_report
                    FOR EACH ROW
                    EXECUTE FUNCTION trg_validate_financial_report_dates();
                    """
                )

                cur.execute(
                    """
                    CREATE OR REPLACE FUNCTION trg_validate_enterprise_metric_value_date()
                    RETURNS TRIGGER
                    LANGUAGE plpgsql
                    AS $$
                    BEGIN
                        IF NEW.value_date > CURRENT_DATE THEN
                            RAISE EXCEPTION 'Дата финансовой записи не может быть в будущем';
                        END IF;
                        RETURN NEW;
                    END;
                    $$;
                    """
                )
                cur.execute("DROP TRIGGER IF EXISTS trg_enterprise_metric_value_validate_date ON enterprise_metric_value;")
                cur.execute(
                    """
                    CREATE TRIGGER trg_enterprise_metric_value_validate_date
                    BEFORE INSERT OR UPDATE ON enterprise_metric_value
                    FOR EACH ROW
                    EXECUTE FUNCTION trg_validate_enterprise_metric_value_date();
                    """
                )

                cur.execute(
                    """
                    CREATE OR REPLACE VIEW vw_financial_report_overview AS
                    SELECT
                        fr.report_id,
                        fr.report_name,
                        fr.date_created,
                        fr.period_start,
                        fr.period_end,
                        e.enterprise_id,
                        e.enterprise_name
                    FROM financial_report fr
                    JOIN enterprise e ON e.enterprise_id = fr.enterprise_id;
                    """
                )
            conn.commit()

    def ensure_reference_data(self) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.executemany(
                    """
                    INSERT INTO unit (unit_name, unit_small_name)
                    VALUES (%s, %s)
                    ON CONFLICT (unit_small_name) DO NOTHING
                    """,
                    UNIT_DEFINITIONS,
                )
                cur.executemany(
                    """
                    INSERT INTO enterprise_metric (metric_small_name, metric_name)
                    VALUES (%s, %s)
                    ON CONFLICT (metric_small_name) DO NOTHING
                    """,
                    [(small_name, name) for small_name, name, _unit in ENTERPRISE_METRIC_DEFINITIONS],
                )
                cur.executemany(
                    """
                    INSERT INTO report_metric (metric_small_name, metric_name)
                    VALUES (%s, %s)
                    ON CONFLICT (metric_small_name) DO NOTHING
                    """,
                    [(small_name, name) for small_name, name, _unit in REPORT_METRIC_DEFINITIONS],
                )
            conn.commit()

    def migrate_legacy_schema(self) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT EXISTS (
                        SELECT 1 FROM information_schema.tables
                        WHERE table_schema = 'public' AND table_name = 'enterprises'
                    )
                    """
                )
                has_legacy_enterprises = bool(cur.fetchone()[0])
                cur.execute(
                    """
                    SELECT EXISTS (
                        SELECT 1 FROM information_schema.tables
                        WHERE table_schema = 'public' AND table_name = 'enterprise_metrics'
                    )
                    """
                )
                has_legacy_metrics = bool(cur.fetchone()[0])
                if not (has_legacy_enterprises and has_legacy_metrics):
                    return

                cur.execute("SELECT COUNT(*) FROM enterprise")
                if int(cur.fetchone()[0]) > 0:
                    return

                cur.execute(
                    """
                    INSERT INTO enterprise (enterprise_name)
                    SELECT name
                    FROM enterprises
                    ON CONFLICT (enterprise_name) DO NOTHING
                    """
                )

                metric_ids = self._get_metric_ids(cur, "enterprise_metric")
                unit_ids = self._get_unit_ids(cur)
                cur.execute(
                    """
                    SELECT e.name, em.period_date, em.revenue, em.cost, em.fixed_expenses, em.variable_expenses, em.tax
                    FROM enterprise_metrics em
                    JOIN enterprises e ON e.id = em.enterprise_id
                    ORDER BY e.name, em.period_date
                    """
                )
                for row in cur.fetchall():
                    enterprise = self._get_enterprise_by_name_with_cursor(cur, row[0])
                    if enterprise is None:
                        continue
                    values = {
                        "revenue": float(row[2]),
                        "cost": float(row[3]),
                        "fixed_expenses": float(row[4]),
                        "variable_expenses": float(row[5]),
                        "tax": float(row[6]),
                    }
                    self._upsert_enterprise_metric_values(
                        cur,
                        enterprise_id=enterprise.id,
                        period_date=row[1],
                        values=values,
                        metric_ids=metric_ids,
                        unit_ids=unit_ids,
                    )
            conn.commit()

    def ensure_default_data(self) -> None:
        for enterprise_name, rows in DEFAULT_ENTERPRISES.items():
            enterprise = self.get_enterprise_by_name(enterprise_name)
            if enterprise is None:
                enterprise = self.add_enterprise(enterprise_name)
            records = [
                FinancialRecord(
                    id=None,
                    enterprise_id=enterprise.id,
                    period_date=datetime.strptime(row["date"], "%Y-%m-%d").date(),
                    revenue=float(row["revenue"]),
                    cost=float(row["cost"]),
                    fixed_expenses=float(row["fixed_expenses"]),
                    variable_expenses=float(row["variable_expenses"]),
                    tax=float(row["tax"]),
                )
                for row in rows
            ]
            self.upsert_records(records)

    def list_enterprises(self) -> list[Enterprise]:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT enterprise_id, enterprise_name, enterprise_small_name
                    FROM enterprise
                    ORDER BY enterprise_name
                    """
                )
                return [Enterprise(id=row[0], name=row[1], small_name=row[2]) for row in cur.fetchall()]

    def add_enterprise(self, name: str, small_name: str | None = None) -> Enterprise:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO enterprise (enterprise_name, enterprise_small_name)
                    VALUES (%s, %s)
                    RETURNING enterprise_id, enterprise_name, enterprise_small_name
                    """,
                    (name, small_name),
                )
                row = cur.fetchone()
            conn.commit()
        return Enterprise(id=row[0], name=row[1], small_name=row[2])

    def delete_enterprise(self, enterprise_id: int) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM enterprise WHERE enterprise_id = %s", (enterprise_id,))
            conn.commit()

    def update_enterprise_name(self, enterprise_id: int, name: str) -> Enterprise:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE enterprise
                    SET enterprise_name = %s
                    WHERE enterprise_id = %s
                    RETURNING enterprise_id, enterprise_name, enterprise_small_name
                    """,
                    (name, enterprise_id),
                )
                row = cur.fetchone()
            conn.commit()
        return Enterprise(id=row[0], name=row[1], small_name=row[2])

    def get_enterprise_by_name(self, name: str) -> Enterprise | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                return self._get_enterprise_by_name_with_cursor(cur, name)

    def get_records(
        self,
        enterprise_id: int,
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[FinancialRecord]:
        conditions = ["emv.enterprise_id = %s"]
        params: list[object] = [enterprise_id]
        if start_date:
            conditions.append("emv.value_date >= %s")
            params.append(start_date)
        if end_date:
            conditions.append("emv.value_date <= %s")
            params.append(end_date)

        query = f"""
            SELECT
                MAX(CASE WHEN em.metric_small_name = 'revenue' THEN emv.value_id END) AS record_id,
                emv.enterprise_id,
                emv.value_date,
                MAX(CASE WHEN em.metric_small_name = 'revenue' THEN emv.value END) AS revenue,
                MAX(CASE WHEN em.metric_small_name = 'cost' THEN emv.value END) AS cost,
                MAX(CASE WHEN em.metric_small_name = 'fixed_expenses' THEN emv.value END) AS fixed_expenses,
                MAX(CASE WHEN em.metric_small_name = 'variable_expenses' THEN emv.value END) AS variable_expenses,
                MAX(CASE WHEN em.metric_small_name = 'tax' THEN emv.value END) AS tax
            FROM enterprise_metric_value emv
            JOIN enterprise_metric em ON em.metric_id = emv.metric_id
            WHERE {' AND '.join(conditions)}
            GROUP BY emv.enterprise_id, emv.value_date
            ORDER BY emv.value_date
        """

        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(query, params)
                rows = cur.fetchall()

        return [
            FinancialRecord(
                id=int(row[0]) if row[0] is not None else None,
                enterprise_id=row[1],
                period_date=row[2],
                revenue=float(row[3]),
                cost=float(row[4]),
                fixed_expenses=float(row[5]),
                variable_expenses=float(row[6]),
                tax=float(row[7]),
            )
            for row in rows
            if all(value is not None for value in row[3:8])
        ]

    def get_record_by_date(self, enterprise_id: int, period_date: date) -> FinancialRecord | None:
        records = self.get_records(enterprise_id, period_date, period_date)
        return records[0] if records else None

    def get_latest_record_date(self, enterprise_id: int) -> date | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT MAX(value_date)
                    FROM enterprise_metric_value
                    WHERE enterprise_id = %s
                    """,
                    (enterprise_id,),
                )
                row = cur.fetchone()
        return row[0] if row and row[0] else None

    def get_record_date_bounds(self, enterprise_id: int) -> tuple[date | None, date | None]:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT MIN(value_date), MAX(value_date)
                    FROM enterprise_metric_value
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
                metric_ids = self._get_metric_ids(cur, "enterprise_metric")
                unit_ids = self._get_unit_ids(cur)
                inserted_id = self._insert_enterprise_metric_values(
                    cur,
                    record.enterprise_id,
                    record.period_date,
                    {
                        "revenue": record.revenue,
                        "cost": record.cost,
                        "fixed_expenses": record.fixed_expenses,
                        "variable_expenses": record.variable_expenses,
                        "tax": record.tax,
                    },
                    metric_ids,
                    unit_ids,
                )
                record.id = inserted_id
            conn.commit()
        return record

    def update_record(self, record: FinancialRecord) -> None:
        if record.id is None:
            raise ValueError("Для обновления записи требуется id")

        with self.connection() as conn:
            with conn.cursor() as cur:
                period_context = self._get_record_identity_from_value_id(cur, record.id)
                if period_context is None:
                    raise ValueError("Запись для обновления не найдена")

                current_enterprise_id, old_period_date = period_context
                if current_enterprise_id != record.enterprise_id:
                    raise ValueError("Изменение предприятия для существующей записи не поддерживается")

                metric_ids = self._get_metric_ids(cur, "enterprise_metric")
                unit_ids = self._get_unit_ids(cur)
                cur.execute(
                    """
                    DELETE FROM enterprise_metric_value
                    WHERE enterprise_id = %s AND value_date = %s
                    """,
                    (record.enterprise_id, old_period_date),
                )
                self._insert_enterprise_metric_values(
                    cur,
                    record.enterprise_id,
                    record.period_date,
                    {
                        "revenue": record.revenue,
                        "cost": record.cost,
                        "fixed_expenses": record.fixed_expenses,
                        "variable_expenses": record.variable_expenses,
                        "tax": record.tax,
                    },
                    metric_ids,
                    unit_ids,
                )
            conn.commit()

    def delete_record(self, record_id: int) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                period_context = self._get_record_identity_from_value_id(cur, record_id)
                if period_context is None:
                    raise ValueError("Запись для удаления не найдена")
                enterprise_id, period_date = period_context
                cur.execute(
                    """
                    DELETE FROM enterprise_metric_value
                    WHERE enterprise_id = %s AND value_date = %s
                    """,
                    (enterprise_id, period_date),
                )
            conn.commit()

    def upsert_record(self, record: FinancialRecord) -> FinancialRecord:
        self.upsert_records([record])
        return self.get_record_by_date(record.enterprise_id, record.period_date) or record

    def upsert_records(self, records: list[FinancialRecord]) -> None:
        if not records:
            return

        with self.connection() as conn:
            with conn.cursor() as cur:
                metric_ids = self._get_metric_ids(cur, "enterprise_metric")
                unit_ids = self._get_unit_ids(cur)
                params: list[tuple[int, int, int, date, float]] = []
                for record in records:
                    for metric_small_name, value in (
                        ("revenue", record.revenue),
                        ("cost", record.cost),
                        ("fixed_expenses", record.fixed_expenses),
                        ("variable_expenses", record.variable_expenses),
                        ("tax", record.tax),
                    ):
                        params.append(
                            (
                                record.enterprise_id,
                                metric_ids[metric_small_name],
                                unit_ids[self._unit_symbol_for_metric(metric_small_name)],
                                record.period_date,
                                value,
                            )
                        )
                cur.executemany(
                    """
                    INSERT INTO enterprise_metric_value (
                        enterprise_id, metric_id, unit_id, value_date, value
                    )
                    VALUES (%s, %s, %s, %s, %s)
                    ON CONFLICT (enterprise_id, metric_id, value_date)
                    DO UPDATE SET
                        unit_id = EXCLUDED.unit_id,
                        value = EXCLUDED.value
                    """,
                    params,
                )
            conn.commit()

    def save_financial_report(
        self,
        report: FinancialReport,
        metric_values: dict[str, float],
    ) -> FinancialReport:
        with self.connection() as conn:
            with conn.cursor() as cur:
                unit_ids = self._get_unit_ids(cur)
                report_metric_ids = self._get_metric_ids(cur, "report_metric")
                cur.execute(
                    """
                    INSERT INTO financial_report (
                        enterprise_id, report_name, date_created, period_start, period_end
                    )
                    VALUES (%s, %s, %s, %s, %s)
                    RETURNING report_id
                    """,
                    (
                        report.enterprise_id,
                        report.name,
                        report.date_created,
                        report.period_start,
                        report.period_end,
                    ),
                )
                report.id = cur.fetchone()[0]
                cur.executemany(
                    """
                    INSERT INTO report_metric_value (report_id, metric_id, unit_id, value)
                    VALUES (%s, %s, %s, %s)
                    ON CONFLICT (report_id, metric_id)
                    DO UPDATE SET
                        unit_id = EXCLUDED.unit_id,
                        value = EXCLUDED.value
                    """,
                    [
                        (
                            report.id,
                            report_metric_ids[metric_small_name],
                            unit_ids[self._unit_symbol_for_report_metric(metric_small_name)],
                            metric_value,
                        )
                        for metric_small_name, metric_value in metric_values.items()
                        if metric_small_name in report_metric_ids
                    ],
                )
            conn.commit()
        return report

    def list_financial_reports(self, enterprise_id: int | None = None) -> list[FinancialReport]:
        params: list[object] = []
        query = """
            SELECT report_id, enterprise_id, report_name, date_created, period_start, period_end
            FROM financial_report
        """
        if enterprise_id is not None:
            query += " WHERE enterprise_id = %s"
            params.append(enterprise_id)
        query += " ORDER BY date_created DESC, report_id DESC"

        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(query, params)
                rows = cur.fetchall()

        return [
            FinancialReport(
                id=row[0],
                enterprise_id=row[1],
                name=row[2],
                date_created=row[3],
                period_start=row[4],
                period_end=row[5],
            )
            for row in rows
        ]

    def get_financial_report(self, report_id: int) -> FinancialReport | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT report_id, enterprise_id, report_name, date_created, period_start, period_end
                    FROM financial_report
                    WHERE report_id = %s
                    """,
                    (report_id,),
                )
                row = cur.fetchone()

        if not row:
            return None

        return FinancialReport(
            id=row[0],
            enterprise_id=row[1],
            name=row[2],
            date_created=row[3],
            period_start=row[4],
            period_end=row[5],
        )

    def delete_financial_report(self, report_id: int) -> None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM financial_report WHERE report_id = %s", (report_id,))
            conn.commit()

    def update_financial_report_name(self, report_id: int, new_name: str) -> FinancialReport | None:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE financial_report
                    SET report_name = %s
                    WHERE report_id = %s
                    RETURNING report_id, enterprise_id, report_name, date_created, period_start, period_end
                    """,
                    (new_name, report_id),
                )
                row = cur.fetchone()
            conn.commit()
        if not row:
            return None
        return FinancialReport(
            id=row[0],
            enterprise_id=row[1],
            name=row[2],
            date_created=row[3],
            period_start=row[4],
            period_end=row[5],
        )

    def get_financial_report_metric_values(self, report_id: int) -> dict[str, float]:
        with self.connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT rm.metric_small_name, rmv.value
                    FROM report_metric_value rmv
                    JOIN report_metric rm ON rm.metric_id = rmv.metric_id
                    WHERE rmv.report_id = %s
                    """,
                    (report_id,),
                )
                rows = cur.fetchall()
        return {row[0]: float(row[1]) for row in rows}

    def get_financial_report_records(self, report_id: int) -> list[FinancialRecord]:
        report = self.get_financial_report(report_id)
        if report is None:
            return []
        return self.get_records(report.enterprise_id, report.period_start, report.period_end)

    def _get_enterprise_by_name_with_cursor(self, cur: psycopg.Cursor, name: str) -> Enterprise | None:
        cur.execute(
            """
            SELECT enterprise_id, enterprise_name, enterprise_small_name
            FROM enterprise
            WHERE enterprise_name = %s
            """,
            (name,),
        )
        row = cur.fetchone()
        return Enterprise(id=row[0], name=row[1], small_name=row[2]) if row else None

    def _get_metric_ids(self, cur: psycopg.Cursor, table_name: str) -> dict[str, int]:
        cur.execute(f"SELECT metric_small_name, metric_id FROM {table_name}")
        return {row[0]: row[1] for row in cur.fetchall()}

    def _get_unit_ids(self, cur: psycopg.Cursor) -> dict[str, int]:
        cur.execute("SELECT unit_small_name, unit_id FROM unit")
        return {row[0]: row[1] for row in cur.fetchall()}

    def _insert_enterprise_metric_values(
        self,
        cur: psycopg.Cursor,
        enterprise_id: int,
        period_date: date,
        values: dict[str, float],
        metric_ids: dict[str, int],
        unit_ids: dict[str, int],
    ) -> int | None:
        inserted_record_id: int | None = None
        for metric_small_name, metric_value in values.items():
            cur.execute(
                """
                INSERT INTO enterprise_metric_value (
                    enterprise_id, metric_id, unit_id, value_date, value
                )
                VALUES (%s, %s, %s, %s, %s)
                RETURNING value_id
                """,
                (
                    enterprise_id,
                    metric_ids[metric_small_name],
                    unit_ids[self._unit_symbol_for_metric(metric_small_name)],
                    period_date,
                    metric_value,
                ),
            )
            value_id = cur.fetchone()[0]
            if metric_small_name == "revenue":
                inserted_record_id = value_id
        return inserted_record_id

    def _upsert_enterprise_metric_values(
        self,
        cur: psycopg.Cursor,
        enterprise_id: int,
        period_date: date,
        values: dict[str, float],
        metric_ids: dict[str, int],
        unit_ids: dict[str, int],
    ) -> None:
        cur.executemany(
            """
            INSERT INTO enterprise_metric_value (
                enterprise_id, metric_id, unit_id, value_date, value
            )
            VALUES (%s, %s, %s, %s, %s)
            ON CONFLICT (enterprise_id, metric_id, value_date)
            DO UPDATE SET
                unit_id = EXCLUDED.unit_id,
                value = EXCLUDED.value
            """,
            [
                (
                    enterprise_id,
                    metric_ids[metric_small_name],
                    unit_ids[self._unit_symbol_for_metric(metric_small_name)],
                    period_date,
                    metric_value,
                )
                for metric_small_name, metric_value in values.items()
            ],
        )

    def _get_record_identity_from_value_id(self, cur: psycopg.Cursor, value_id: int) -> tuple[int, date] | None:
        cur.execute(
            """
            SELECT enterprise_id, value_date
            FROM enterprise_metric_value
            WHERE value_id = %s
            """,
            (value_id,),
        )
        row = cur.fetchone()
        return (row[0], row[1]) if row else None

    @staticmethod
    def _unit_symbol_for_metric(metric_small_name: str) -> str:
        if metric_small_name in {"revenue", "cost", "fixed_expenses", "variable_expenses", "tax"}:
            return "₽"
        return "ед."

    @staticmethod
    def _unit_symbol_for_report_metric(metric_small_name: str) -> str:
        if metric_small_name in {"target_ros", "avg_ros", "std_ros", "min_ros", "max_ros"}:
            return "%"
        if metric_small_name == "avg_profit":
            return "₽"
        return "ед."
