from dataclasses import replace

import pytest

from valuation_engine.model import Drivers, ModelError, cost_of_capital, discount, fcff, project, value


@pytest.fixture
def drivers():
    return Drivers((0.,)*5, (.2,)*5, (.03,)*5, (.04,)*5, (.1,)*5, .25, .1, .02)


def test_hand_calculable_perpetuity():
    # Five years of 10; terminal year cash flow 10, g=0: EV exactly 10 / 10%.
    result = discount([10]*5, 10, .1, 0, 5, 20, diluted_shares=10)
    assert result.enterprise_value == pytest.approx(100)
    assert result.equity_value == pytest.approx(85)
    assert result.per_share == pytest.approx(8.5)


def test_driver_projection_and_stable_growth_transition(drivers):
    forecast, terminal = project(100, 10, drivers)
    assert forecast.fcff.to_list() == pytest.approx([14]*5)
    # Year 6: revenue 102, EBIT 20.4, tax 5.1, D&A 3.06, capex 4.08, dNWC .2.
    assert terminal == pytest.approx(14.08)
    assert terminal != pytest.approx(14*1.02)
    _, result = value(100, 10, drivers, dict(cash=5, debt=20))
    expected = sum(14/1.1**t for t in range(1,6)) + (14.08/.08)/1.1**5
    assert result.enterprise_value == pytest.approx(expected)
    assert result.per_share is None


def test_negative_cash_flow_and_no_tax_benefit(drivers):
    assert fcff(-10, .25, 3, 4, 2) == -13
    forecast, _ = project(100, 10, replace(drivers, margin=(-.1, .2, .2, .2, .2)))
    assert forecast.fcff.iloc[0] == -11
    assert discount([-10,10,10,10,10], 10, .1, 0, 0, 0).enterprise_value < 100


@pytest.mark.parametrize("wacc,g", [(.03,.03),(.03,.04),(0,0),(-.1,-.2),(.1,-1)])
def test_invalid_discount_rates(wacc,g):
    with pytest.raises(ModelError):
        discount([10]*5, 10, wacc, g, 0, 0)


def test_zero_shares_and_missing_data(drivers):
    with pytest.raises(ModelError, match="shares"):
        discount([10]*5, 10, .1, 0, 0, 0, diluted_shares=0)
    with pytest.raises(ModelError):
        project(None, 0, drivers)
    with pytest.raises(ModelError):
        project(0, 0, drivers)
    with pytest.raises(ModelError):
        project(100, 10, replace(drivers, growth=(0.,)*4))
    with pytest.raises(ModelError, match="terminal FCFF"):
        project(100, 10, replace(drivers, margin=(-.1,)*5))


def test_capm_and_wacc():
    equity, wacc = cost_of_capital(dict(risk_free_rate=.04, beta=1.2, equity_risk_premium=.05,
                                     cost_of_debt=.06, debt_weight=.2, tax_rate=.25))
    assert equity == pytest.approx(.1)
    assert wacc == pytest.approx(.089)
