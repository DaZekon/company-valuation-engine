"""Financial calculations. Monetary values and shares use millions throughout."""
from __future__ import annotations

from dataclasses import dataclass
from math import isfinite
from typing import Sequence

import numpy as np
import pandas as pd


class ModelError(ValueError):
    """An input cannot support the requested valuation."""


def finite(value: float, label: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ModelError(f"{label} must be numeric") from exc
    if not isfinite(number):
        raise ModelError(f"{label} must be finite")
    return number


def validate_rates(wacc: float, growth: float) -> None:
    finite(wacc, "WACC")
    finite(growth, "Terminal growth")
    if not 0 < wacc <= 1:
        raise ModelError("WACC must be greater than zero and at most 100%")
    if not -1 < growth < wacc:
        raise ModelError("Terminal growth must be greater than -100% and less than WACC")


def cost_of_capital(inputs: dict) -> tuple[float, float]:
    """CAPM cost of equity and market-value-weighted after-tax WACC."""
    names = ("risk_free_rate", "beta", "equity_risk_premium", "cost_of_debt", "debt_weight", "tax_rate")
    values = {name: finite(inputs[name], name) for name in names}
    if not 0 <= values["debt_weight"] <= 1 or not 0 <= values["tax_rate"] <= 1:
        raise ModelError("Debt weight and tax rate must be between 0 and 1")
    if min(values[n] for n in ("risk_free_rate", "equity_risk_premium", "cost_of_debt")) < 0:
        raise ModelError("This compact CAPM model requires nonnegative rate inputs")
    equity = values["risk_free_rate"] + values["beta"] * values["equity_risk_premium"]
    wacc = equity * (1 - values["debt_weight"]) + values["cost_of_debt"] * (1 - values["tax_rate"]) * values["debt_weight"]
    if equity < 0 or not 0 < wacc <= 1:
        raise ModelError("Cost of equity and WACC are outside supported ranges")
    return equity, wacc


@dataclass(frozen=True)
class Drivers:
    growth: tuple[float, ...]
    margin: tuple[float, ...]
    da_ratio: tuple[float, ...]
    capex_ratio: tuple[float, ...]
    nwc_ratio: tuple[float, ...]
    tax_rate: float
    wacc: float
    terminal_growth: float

    @classmethod
    def from_config(cls, case: dict, capital: dict) -> "Drivers":
        _, calculated_wacc = cost_of_capital(capital)
        # None means calculated WACC; an explicit zero is never treated as missing.
        override = case.get("wacc_override")
        wacc = calculated_wacc + finite(case.get("wacc_spread", 0), "WACC spread") if override is None else finite(override, "WACC override")
        return cls(**{n: tuple(finite(v, f"{n} year {i+1}") for i,v in enumerate(case[n])) for n in ("growth", "margin", "da_ratio", "capex_ratio", "nwc_ratio")},
                   tax_rate=finite(capital["tax_rate"], "Tax rate"), wacc=wacc,
                   terminal_growth=finite(case["terminal_growth"], "Terminal growth"))

    def validate(self) -> None:
        validate_rates(self.wacc, self.terminal_growth)
        if not 0 <= finite(self.tax_rate, "Tax rate") <= 1:
            raise ModelError("Tax rate must be between 0 and 1")
        for name in ("growth", "margin", "da_ratio", "capex_ratio", "nwc_ratio"):
            items = getattr(self, name)
            if len(items) != 5:
                raise ModelError(f"{name} requires exactly five annual assumptions")
            for value in items:
                value = finite(value, name)
                if name == "growth" and value <= -1:
                    raise ModelError("Revenue growth must exceed -100%")
                if name in ("da_ratio", "capex_ratio") and value < 0:
                    raise ModelError(f"{name} cannot be negative")
                if name in ("margin", "nwc_ratio") and not -1 <= value <= 1:
                    raise ModelError(f"{name} must be between -100% and 100%")


@dataclass(frozen=True)
class Valuation:
    enterprise_value: float
    equity_value: float
    per_share: float | None
    terminal_value: float
    terminal_fcff: float
    pv_terminal: float
    pv_explicit: float


def fcff(ebit: float, tax_rate: float, da: float, capex: float, delta_nwc: float) -> float:
    """No immediate cash tax benefit on losses; no NOL carryforward schedule."""
    return ebit - max(ebit, 0) * tax_rate + da - capex - delta_nwc


def project(revenue: float, nwc: float, drivers: Drivers) -> tuple[pd.DataFrame, float]:
    drivers.validate()
    revenue, nwc = finite(revenue, "Starting revenue"), finite(nwc, "Starting operating NWC")
    if revenue <= 0:
        raise ModelError("Starting revenue must be positive")
    rows = []
    for year in range(5):
        old_nwc = nwc
        revenue *= 1 + drivers.growth[year]
        ebit = revenue * drivers.margin[year]
        da, capex = revenue * drivers.da_ratio[year], revenue * drivers.capex_ratio[year]
        nwc = revenue * drivers.nwc_ratio[year]
        taxes = max(ebit, 0) * drivers.tax_rate
        rows.append(dict(year=year + 1, revenue=revenue, ebit=ebit, taxes=taxes,
                         nopat=ebit - taxes, da=da, capex=capex, nwc=nwc,
                         delta_nwc=nwc - old_nwc,
                         fcff=fcff(ebit, drivers.tax_rate, da, capex, nwc - old_nwc)))
    forecast = pd.DataFrame(rows)
    # Rebuild year 6 at stable growth: year 5 growth can differ from terminal g.
    revenue6 = revenue * (1 + drivers.terminal_growth)
    terminal_fcff = fcff(revenue6 * drivers.margin[-1], drivers.tax_rate,
                         revenue6 * drivers.da_ratio[-1], revenue6 * drivers.capex_ratio[-1],
                         revenue6 * drivers.nwc_ratio[-1] - nwc)
    if terminal_fcff <= 0:
        raise ModelError("Nonpositive terminal FCFF: perpetual-growth valuation requires a sustainable positive year 6; revise explicit assumptions")
    if not np.isfinite(forecast.select_dtypes(include="number").to_numpy()).all() or not isfinite(terminal_fcff):
        raise ModelError("Projection overflowed; reduce assumptions")
    return forecast, terminal_fcff


def discount(cash_flows: Sequence[float], terminal_fcff: float, wacc: float, growth: float,
             cash: float, debt: float, other_adjustments: float = 0,
             diluted_shares: float | None = None) -> Valuation:
    validate_rates(wacc, growth)
    if len(cash_flows) != 5:
        raise ModelError("This model requires five annual cash flows")
    flows = np.array([finite(x, "FCFF") for x in cash_flows], dtype=float)
    terminal_fcff = finite(terminal_fcff, "Terminal FCFF")
    if terminal_fcff <= 0:
        raise ModelError("Terminal FCFF must be positive for this perpetual-growth model")
    cash, debt = finite(cash, "Cash adjustment"), finite(debt, "Debt adjustment")
    other_adjustments = finite(other_adjustments, "Other equity bridge adjustments")
    if cash < 0 or debt < 0:
        raise ModelError("Cash and debt adjustments must be nonnegative")
    if diluted_shares is not None and finite(diluted_shares, "Diluted shares") <= 0:
        raise ModelError("Diluted shares must be positive or omitted")
    pv_explicit = float(np.sum(flows / np.power(1 + wacc, np.arange(1, 6))))
    terminal = terminal_fcff / (wacc - growth)
    pv_terminal = terminal / (1 + wacc) ** 5
    ev = pv_explicit + pv_terminal
    equity = ev + cash - debt + other_adjustments
    return Valuation(ev, equity, None if diluted_shares is None else equity / diluted_shares,
                     terminal, terminal_fcff, pv_terminal, pv_explicit)


def value(revenue: float, nwc: float, drivers: Drivers, bridge: dict) -> tuple[pd.DataFrame, Valuation]:
    forecast, terminal = project(revenue, nwc, drivers)
    return forecast, discount(forecast.fcff.to_list(), terminal, drivers.wacc,
                             drivers.terminal_growth, **bridge)
