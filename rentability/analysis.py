from __future__ import annotations

import math
from typing import Iterable

import numpy as np


def t_test_one_sample(data: Iterable[float], mu: float) -> tuple[float, float]:
    values = list(data)
    sample_size = len(values)
    if sample_size < 2:
        raise ValueError("Для t-теста требуется как минимум 2 наблюдения")

    mean = sum(values) / sample_size
    variance = sum((item - mean) ** 2 for item in values) / (sample_size - 1)
    std_dev = math.sqrt(variance)
    sem = std_dev / math.sqrt(sample_size)
    if sem == 0:
        raise ValueError("Невозможно вычислить t-статистику при нулевой дисперсии")

    t_stat = (mean - mu) / sem
    p_value = t_cdf(t_stat, sample_size - 1)
    return t_stat, p_value


def t_cdf(t_value: float, degrees_of_freedom: int) -> float:
    if degrees_of_freedom > 30:
        return 0.5 * (1 + math.erf(t_value / math.sqrt(2)))

    x_value = (t_value + math.sqrt(t_value * t_value + degrees_of_freedom)) / (
        2 * math.sqrt(t_value * t_value + degrees_of_freedom)
    )
    return beta_cdf(x_value, degrees_of_freedom / 2, degrees_of_freedom / 2)


def beta_cdf(x_value: float, a_value: float, b_value: float) -> float:
    if x_value <= 0:
        return 0.0
    if x_value >= 1:
        return 1.0

    steps = 1000
    delta = x_value / steps
    integral = 0.0

    for index in range(steps):
        xi = index * delta + delta / 2
        if 0 < xi < 1:
            integral += (xi ** (a_value - 1)) * ((1 - xi) ** (b_value - 1)) * delta

    beta = math.gamma(a_value) * math.gamma(b_value) / math.gamma(a_value + b_value)
    return integral / beta


def summarize_ros(ros_values: Iterable[float]) -> tuple[float, float]:
    values = list(ros_values)
    if len(values) < 2:
        raise ValueError("Недостаточно данных для расчета статистики")
    return float(np.mean(values)), float(np.std(values, ddof=1))
