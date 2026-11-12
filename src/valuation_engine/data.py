"""Annual CSV normalization and explicit review gates."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import logging

import numpy as np
import pandas as pd

from .model import ModelError, finite

LOG = logging.getLogger(__name__)
FIELDS = ("revenue", "net_income", "cfo", "capex", "cash", "debt", "ebit", "da", "operating_nwc", "diluted_weighted_avg_shares")
META = ("fiscal_start", "fiscal_end", "currency", "unit_scale", "source", "retrieved_at")


def load_history(path: str | Path) -> pd.DataFrame:
    frame = pd.read_csv(path, dtype={"source": str, "currency": str})
    missing = set(META + FIELDS) - set(frame.columns)
    if missing:
        raise ModelError(f"CSV missing columns: {', '.join(sorted(missing))}")
    if len(frame) < 3:
        raise ModelError("Provide at least three annual fiscal periods")
    try:
        for col in ("fiscal_start", "fiscal_end", "retrieved_at"):
            frame[col] = pd.to_datetime(frame[col], errors="raise")
    except (ValueError, TypeError) as exc:
        raise ModelError("Invalid fiscal or retrieval date") from exc
    if frame[list(META)].isna().any().any() or frame.source.str.strip().eq("").any():
        raise ModelError("Period, source, units and retrieval date cannot be missing")
    frame = frame.sort_values("fiscal_end").reset_index(drop=True)
    if frame.fiscal_end.duplicated().any():
        raise ModelError("Duplicate fiscal end dates")
    durations = (frame.fiscal_end - frame.fiscal_start).dt.days + 1
    if not durations.between(330, 380).all():
        raise ModelError("Only complete 330–380 day annual periods are supported")
    for i in range(1, len(frame)):
        if (frame.fiscal_start[i] - frame.fiscal_end[i-1]).days != 1:
            raise ModelError("Inconsistent fiscal dates: annual periods overlap or contain a gap")
    if frame.currency.nunique() != 1 or frame.currency.iloc[0] != "USD":
        raise ModelError("This version supports USD only; no implicit FX conversion")
    frame["unit_scale"] = pd.to_numeric(frame.unit_scale, errors="coerce")
    if not frame.unit_scale.isin([1, 1000, 1000000]).all():
        raise ModelError("unit_scale must be 1, 1000 or 1000000")
    for field in FIELDS:
        original = frame[field]
        parsed = pd.to_numeric(original, errors="coerce")
        if (original.notna() & parsed.isna()).any() or np.isinf(parsed).any():
            raise ModelError(f"Invalid numeric value in {field}")
        frame[field] = parsed * frame.unit_scale / 1_000_000
    frame["unit_scale"] = 1_000_000
    for field in ("revenue", "capex", "cash", "debt", "da", "diluted_weighted_avg_shares"):
        present = frame[field].dropna()
        if (present < 0).any():
            raise ModelError(f"{field} must be nonnegative; capex is a positive cash outflow")
    for field in ("revenue", "net_income", "cfo", "capex", "cash", "debt"):
        if frame[field].isna().any():
            LOG.warning("Historical %s is missing in some periods; no substitute was inserted", field)
    return frame


def validate_config(config: dict, history: pd.DataFrame) -> dict:
    for name in ("company", "ticker", "valuation_date", "review_note", "assumption_source", "assumption_date", "bridge"):
        if not config.get(name):
            raise ModelError(f"Missing config field: {name}")
    if config.get("reviewed") is not True:
        raise ModelError("Review annual inputs against filings, fill gaps manually, then set reviewed=true and supply review_note")
    if config.get("industry", "nonfinancial") != "nonfinancial":
        raise ModelError("This FCFF model supports nonfinancial operating companies; banks/insurers require a different method")
    end = history.fiscal_end.iloc[-1].date()
    try:
        if date.fromisoformat(config["valuation_date"]) != end:
            raise ModelError("Valuation date must equal the last fiscal end: this version is a retrospective year-end DCF")
        date.fromisoformat(config["assumption_date"])
    except (ValueError, TypeError) as exc:
        raise ModelError("Config dates must use YYYY-MM-DD") from exc
    latest = history.iloc[-1]
    needed = ("revenue", "ebit", "da", "capex", "operating_nwc", "cash", "debt")
    missing = [x for x in needed if pd.isna(latest[x])]
    if missing:
        raise ModelError("Missing latest-year inputs: " + ", ".join(missing) + ". Fill reviewed manual CSV inputs with sources.")
    if latest.revenue <= 0:
        raise ModelError("Latest revenue must be positive")
    # Share count is a separate reviewed capital-structure input, never weighted-average EPS shares.
    shares = config.get("diluted_shares")
    if shares is not None:
        if finite(shares, "Diluted shares") <= 0 or not config.get("shares_source"):
            raise ModelError("Positive point-in-time diluted shares (millions) and shares_source are required")
        if config.get("shares_as_of") != config["valuation_date"]:
            raise ModelError("Reviewed diluted shares must be as of the valuation date")
    bridge = config["bridge"]
    amounts={key:finite(bridge.get(key),key) for key in ("cash_adjustment", "debt_adjustment", "other_adjustments")}
    if amounts["cash_adjustment"] < 0 or amounts["debt_adjustment"] < 0:
        raise ModelError("Bridge cash and debt adjustments cannot be negative")
    if not bridge.get("source"):
        raise ModelError("Bridge adjustments require a source and explanation, including zero other adjustments")
    if set(config.get("scenarios", {})) != {"Bear", "Base", "Bull"}:
        raise ModelError("Provide exactly Bear, Base and Bull scenarios")
    return dict(cash=amounts["cash_adjustment"], debt=amounts["debt_adjustment"],
                other_adjustments=amounts["other_adjustments"], diluted_shares=None if shares is None else finite(shares,"Diluted shares"))


def historical_metrics(history: pd.DataFrame) -> pd.DataFrame:
    metrics = history.copy()
    prior = metrics.revenue.shift(1)
    metrics["revenue_growth"] = metrics.revenue.div(prior.where(prior.ne(0))) - 1
    metrics["operating_margin"] = metrics.ebit.div(metrics.revenue.where(metrics.revenue.ne(0)))
    metrics["cfo_less_capex"] = metrics.cfo - metrics.capex
    metrics["delta_nwc"] = metrics.operating_nwc.diff()
    return metrics
