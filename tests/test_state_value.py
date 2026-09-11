import numpy as np
import pandas as pd
from datetime import date, timedelta

from state_value import settlement_b, forecast_load_decomposed, finite_grid_value_dp, dp_policy_rollout


def test_q3_settlement_components_follow_statement():
    bill = settlement_b(np.array([10.0, 10.0]), np.array([5.0, 15.0]), np.array([2.0, 2.0]))
    assert bill["base"] == 40.0
    assert bill["reduction_fee"] == 5.0
    assert bill["increase_fee"] == 15.0
    assert bill["total"] == 60.0


def test_load_forecast_does_not_read_target_or_future_rows():
    dates = [date(2025, 1, 1) + timedelta(days=i) for i in range(5)]
    # The target-day/future rows are intentionally extreme.  They must not
    # enter a forecast for the target date.
    frame = pd.DataFrame(np.array([[100.0] * 2, [110.0] * 2, [120.0] * 2,
                                   [9999.0] * 2, [9999.0] * 2]), index=dates)
    pred = forecast_load_decomposed(frame, dates[3], dates, 1.0)
    assert np.all(pred < 1000.0)


def test_grid_dp_rollout_is_feasible_and_exclusive():
    net = np.array([2.0, 0.0, 2.0, 0.0])
    price = np.array([1.0, 1.0, 3.0, 3.0])
    dp = finite_grid_value_dp(net, price, initial=10.0, terminal=10.0,
                              s_min=0.0, s_max=20.0, eta=0.9,
                              p_max_kwh=5.0, step=1.0)
    result = dp_policy_rollout(dp, net, price, initial=10.0, eta=0.9)
    assert np.isfinite(dp["initial_value"])
    assert np.isfinite(result["cost"])
    assert np.all(np.minimum(result["c"], result["d"]) <= 1e-12)
    assert abs(result["s"][0] - 10.0) < 1e-9
    assert abs(result["s"][-1] - 10.0) < 1e-9
