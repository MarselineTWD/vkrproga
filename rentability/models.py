from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(slots=True)
class Enterprise:
    id: int
    name: str
    small_name: str | None = None


@dataclass(slots=True)
class FinancialReport:
    id: int | None
    enterprise_id: int
    name: str
    date_created: date
    period_start: date
    period_end: date


@dataclass(slots=True)
class FinancialRecord:
    id: int | None
    enterprise_id: int
    period_date: date
    revenue: float
    cost: float
    fixed_expenses: float
    variable_expenses: float
    tax: float

    @property
    def net_profit(self) -> float:
        return (
            self.revenue
            - self.cost
            - self.fixed_expenses
            - self.variable_expenses
            - self.tax
        )

    @property
    def ros(self) -> float:
        return (self.net_profit / self.revenue) * 100 if self.revenue else 0.0
