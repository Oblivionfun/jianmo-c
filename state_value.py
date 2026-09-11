"""Small, auditable building blocks for the C题 adversarial update.

The module deliberately keeps the settlement rule and the inventory-value
calculation independent from the long back-test driver.  This makes it harder
to accidentally change a billing convention while changing a forecast.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Iterable

import numpy as np


def settlement_b(plan: np.ndarray, adjusted: np.ndarray, price: np.ndarray) -> dict:
    """Return the Q3 bill under the wording in the statement.

    The planned quantity is billed at the normal price.  A reduction relative
    to the plan is charged at 0.5 times the transaction price and an increase
    is charged at 1.5 times that price.  The returned components are useful in
    the paper and prevent the common mistake of billing the adjusted quantity
    as if it were the original plan.
    """
    plan = np.asarray(plan, dtype=float)
    adjusted = np.asarray(adjusted, dtype=float)
    price = np.asarray(price, dtype=float)
    if not (plan.shape == adjusted.shape == price.shape):
        raise ValueError("plan, adjusted and price must have the same shape")
    reduction = np.maximum(plan - adjusted, 0.0)
    increase = np.maximum(adjusted - plan, 0.0)
    base = price * plan
    reduction_fee = 0.5 * price * reduction
    increase_fee = 1.5 * price * increase
    return {
        "base": float(base.sum()),
        "reduction_fee": float(reduction_fee.sum()),
        "increase_fee": float(increase_fee.sum()),
        "total": float((base + reduction_fee + increase_fee).sum()),
        "reduction_kwh": float(reduction.sum()),
        "increase_kwh": float(increase.sum()),
    }


def settlement_b_vector(plan: np.ndarray, adjusted: np.ndarray, price: np.ndarray) -> float:
    """Scalar convenience wrapper used by vectorized back-tests."""
    return settlement_b(plan, adjusted, price)["total"]


def forecast_load_decomposed(load, target: date, dates: Iterable[date], dt: float,
                             same_type_window: int = 8) -> np.ndarray:
    """Causal daily-level/shape load forecast.

    Only dates strictly before ``target`` are considered.  Recent days with
    the same weekday are preferred; if fewer than three exist, all recent days
    are used.  The daily energy level and normalized intraday shape are
    estimated separately, so a high-load day does not distort the shape of a
    low-load day.  ``load`` is a date-indexed dataframe in kW.
    """
    target = target if isinstance(target, date) else target.date()
    all_dates = list(dates)
    prior = [d for d in all_dates if d < target and d in load.index]
    if not prior:
        return np.full(load.shape[1], 6000.0)
    same = [d for d in prior if d.weekday() == target.weekday()][-same_type_window:]
    chosen = same if len(same) >= 3 else prior[-same_type_window:]
    arr = np.asarray(load.loc[chosen].values, dtype=float)
    daily = np.nansum(arr * dt, axis=1)
    safe = np.where(daily > 1e-9, daily, 1.0)
    shape = arr / safe[:, None] / dt
    level = float(np.nanmedian(daily))
    profile = np.nanmedian(shape, axis=0)
    profile = np.maximum(profile, 0.0)
    norm = float(np.nansum(profile * dt))
    if not np.isfinite(norm) or norm <= 1e-9:
        return np.nanmedian(arr, axis=0)
    return level * profile / norm


def finite_grid_value_dp(net_load_kwh: np.ndarray, price: np.ndarray, *,
                         initial: float, terminal: float, s_min: float,
                         s_max: float, eta: float, p_max_kwh: float,
                         step: float = 20.0) -> dict:
    """Approximate the deterministic inventory value function on a grid.

    The state is internal battery energy.  For each state transition only one
    of charge/discharge is used, so the recursion is physically explicit and
    avoids a hidden simultaneous-charge/discharge degree of freedom.  The
    result is an *independent numerical cross-check* for the LP, not a claim
    that the grid is an exact continuous DP.
    """
    n = len(net_load_kwh)
    net_load_kwh = np.asarray(net_load_kwh, dtype=float)
    price = np.asarray(price, dtype=float)
    if net_load_kwh.shape != price.shape:
        raise ValueError("net load and price must have the same length")
    grid = np.arange(s_min, s_max + 0.5 * step, step, dtype=float)
    if not np.any(np.isclose(grid, initial)) or not np.any(np.isclose(grid, terminal)):
        raise ValueError("initial and terminal must lie on the DP grid")
    terminal_idx = int(np.argmin(np.abs(grid - terminal)))
    value = np.full((n + 1, grid.size), np.inf, dtype=float)
    value[n, terminal_idx] = 0.0
    policy = np.full((n, grid.size), -1, dtype=int)
    max_delta = p_max_kwh * eta
    for t in range(n - 1, -1, -1):
        for i, s in enumerate(grid):
            lo = max(s_min, s - max_delta)
            hi = min(s_max, s + max_delta)
            js = np.flatnonzero((grid >= lo - 1e-9) & (grid <= hi + 1e-9))
            delta = grid[js] - s
            charge = np.maximum(delta, 0.0) / eta
            discharge = np.maximum(-delta, 0.0) * eta
            grid_cost = price[t] * np.maximum(net_load_kwh[t] + charge - discharge, 0.0)
            cand = grid_cost + value[t + 1, js]
            if cand.size:
                k = int(np.argmin(cand))
                value[t, i] = cand[k]
                policy[t, i] = int(js[k])
    initial_idx = int(np.argmin(np.abs(grid - initial)))
    return {
        "grid": grid,
        "value": value,
        "policy": policy,
        "initial_value": float(value[0, initial_idx]),
        "grid_step_kwh": float(step),
        "reachable": int(np.isfinite(value[0]).sum()),
    }


def dp_policy_rollout(dp: dict, net_load_kwh: np.ndarray, price: np.ndarray,
                      initial: float, eta: float) -> dict:
    """Roll out a grid-DP policy and return charge/discharge/purchase arrays."""
    grid = dp["grid"]
    idx = int(np.argmin(np.abs(grid - initial)))
    charge = np.zeros(len(net_load_kwh))
    discharge = np.zeros(len(net_load_kwh))
    purchase = np.zeros(len(net_load_kwh))
    states = np.zeros(len(net_load_kwh) + 1)
    states[0] = grid[idx]
    for t, net in enumerate(np.asarray(net_load_kwh, dtype=float)):
        nxt = int(dp["policy"][t, idx])
        if nxt < 0:
            raise RuntimeError("DP policy is infeasible at rollout state")
        delta = grid[nxt] - grid[idx]
        if delta >= 0:
            charge[t] = delta / eta
        else:
            discharge[t] = -delta * eta
        purchase[t] = max(0.0, net + charge[t] - discharge[t])
        states[t + 1] = grid[nxt]
        idx = nxt
    return {"q": purchase, "c": charge, "d": discharge, "s": states,
            "cost": float(np.dot(price, purchase))}
