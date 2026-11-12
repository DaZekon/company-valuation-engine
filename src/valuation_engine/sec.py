"""Conservative US-GAAP SEC importer. Every extracted fact keeps its provenance."""
from __future__ import annotations

from datetime import date, datetime, timezone
import hashlib
import json
import logging
from pathlib import Path
import re
import time

import pandas as pd
import requests

from .data import FIELDS, META
from .model import ModelError

LOG = logging.getLogger(__name__)
TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
TAGS = {
    "revenue": ("RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet"),
    "net_income": ("NetIncomeLoss", "ProfitLoss"),
    "cfo": ("NetCashProvidedByUsedInOperatingActivities",),
    "capex": ("PaymentsToAcquirePropertyPlantAndEquipment",),
    "cash": ("CashAndCashEquivalentsAtCarryingValue",),
    "ebit": ("OperatingIncomeLoss",),
    "da": ("DepreciationDepletionAndAmortization", "DepreciationDepletionAndAmortizationPropertyPlantAndEquipment"),
    "diluted_weighted_avg_shares": ("WeightedAverageNumberOfDilutedSharesOutstanding",),
}
INSTANT = {"cash", "debt_current", "debt_noncurrent", "short_term_borrowings"}


class SecError(RuntimeError):
    pass


class SecClient:
    """At most two requests/second per instance. Use one process per user/IP."""
    def __init__(self, user_agent: str, cache_dir: str | Path = ".cache/sec", cache_hours: int = 24):
        if not re.search(r"[^\s@]+@[^\s@]+\.[^\s@]+", user_agent) or len(user_agent.split()) < 2:
            raise ModelError("SEC_USER_AGENT must identify your application/person and real contact email")
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate"})
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.cache_hours = cache_hours
        self.last_request = 0.0

    def get_json(self, url: str) -> tuple[dict, str]:
        key = self.cache_dir / (hashlib.sha256(url.encode()).hexdigest() + ".json")
        if key.exists():
            cached = json.loads(key.read_text(encoding="utf-8"))
            age = datetime.now(timezone.utc) - datetime.fromisoformat(cached["retrieved_at"])
            if age.total_seconds() < self.cache_hours * 3600:
                LOG.info("Using cached SEC response retrieved %s", cached["retrieved_at"])
                return cached["data"], cached["retrieved_at"]
        for attempt in range(3):
            time.sleep(max(0.0, .5 - (time.monotonic() - self.last_request)))
            self.last_request = time.monotonic()
            try:
                response = self.session.get(url, timeout=(10, 40))
                # A 403 is not bypassed or retried with an impersonated identity.
                if response.status_code in (429, 500, 502, 503, 504) and attempt < 2:
                    retry = response.headers.get("Retry-After", "")
                    delay = float(retry) if retry.isdigit() else 2 ** (attempt + 1)
                    time.sleep(min(max(delay, 1), 60))
                    continue
                response.raise_for_status()
                payload = response.json()
            except (requests.RequestException, ValueError) as exc:
                raise SecError(f"SEC request failed: {exc}. Use the CSV fallback; no financial values were synthesized.") from exc
            retrieved = datetime.now(timezone.utc).isoformat()
            key.write_text(json.dumps(dict(retrieved_at=retrieved, data=payload)), encoding="utf-8")
            return payload, retrieved
        raise SecError("SEC retries exhausted. Use CSV fallback.")

    def resolve(self, ticker: str) -> tuple[int, str]:
        ticker = ticker.strip().upper()
        if not re.fullmatch(r"[A-Z0-9.\-]{1,12}", ticker):
            raise ModelError("Ticker contains unsupported characters")
        companies, _ = self.get_json(TICKERS_URL)
        matches = [r for r in companies.values() if r["ticker"].upper() == ticker]
        if len(matches) != 1:
            raise SecError(f"Ticker {ticker} is unavailable or ambiguous in the SEC ticker map")
        return int(matches[0]["cik_str"]), matches[0]["title"]

    def fetch(self, ticker: str) -> tuple[dict, str, int, str]:
        cik, company = self.resolve(ticker)
        facts, retrieved = self.get_json(FACTS_URL.format(cik=cik))
        return facts, retrieved, cik, company


def candidates(facts: dict, tags: tuple[str, ...], instant: bool = False, shares: bool = False) -> list[dict]:
    output = []
    for priority, tag in enumerate(tags):
        for fact in facts.get("facts", {}).get("us-gaap", {}).get(tag, {}).get("units", {}).get("shares" if shares else "USD", []):
            if fact.get("form") not in ("10-K", "10-K/A") or not fact.get("end") or not fact.get("filed"):
                continue
            if not instant:
                if not fact.get("start"):
                    continue
                duration = (date.fromisoformat(fact["end"]) - date.fromisoformat(fact["start"])).days + 1
                if not 330 <= duration <= 380:
                    continue
            output.append({**fact, "tag": tag, "priority": priority})
    return output


def pick(items: list[dict], start: str, end: str, instant: bool = False) -> dict | None:
    matched = [f for f in items if f["end"] == end and (instant or f["start"] == start)]
    if not matched:
        return None
    # Latest filing wins (restatements included); taxonomy preference breaks ties.
    latest_filed = max(x["filed"] for x in matched)
    matched = [x for x in matched if x["filed"] == latest_filed]
    priority = min(x["priority"] for x in matched)
    matched = [x for x in matched if x["priority"] == priority]
    if len({x["val"] for x in matched}) > 1:
        raise ModelError(f"Conflicting SEC facts for {end}: {matched[0]['tag']}. Resolve manually.")
    return matched[0]


def normalize_facts(facts: dict, retrieved_at: str, years: int = 3) -> tuple[pd.DataFrame, list[dict], list[str]]:
    if years < 3:
        raise ModelError("Retrieve at least three fiscal years")
    revenue_facts = candidates(facts, TAGS["revenue"])
    periods = sorted({(f["start"], f["end"]) for f in revenue_facts}, key=lambda p: p[1])
    # Reject ambiguous annual calendars instead of combining overlapping contexts.
    if len({e for _,e in periods}) != len(periods):
        raise ModelError("Conflicting annual revenue periods; use reviewed CSV input")
    periods = periods[-years:]
    if len(periods) < 3:
        raise ModelError("Fewer than three standard US-GAAP annual revenue periods; use CSV")
    warnings = ["SEC extraction is unreviewed. Reconcile every selected tag, period and scope to the filing.",
                "Operating income is an EBIT proxy; review nonoperating items. D&A may include depletion or omit amortization.",
                "Operating NWC is left blank: define operating assets less operating liabilities, excluding cash and financing.",
                "Capex includes cash PP&E purchases only. Review capitalized software, intangibles, leases and acquisitions.",
                "Weighted-average diluted EPS shares are historical context only; they are never used as current diluted shares."]
    rows, provenance = [], []
    debt_tags = {"debt_current": ("LongTermDebtCurrent",), "debt_noncurrent": ("LongTermDebtNoncurrent",),
                 "short_term_borrowings": ("ShortTermBorrowings",)}
    pools = {field: candidates(facts, tags, field in INSTANT, field == "diluted_weighted_avg_shares") for field, tags in {**TAGS, **debt_tags}.items()}
    cik = int(facts.get("cik", 0))
    for start, end in periods:
        row = {field: None for field in FIELDS}
        row.update(fiscal_start=start, fiscal_end=end, currency="USD", unit_scale=1,
                   source=FACTS_URL.format(cik=cik), retrieved_at=retrieved_at)
        selected = {}
        for field, pool in pools.items():
            fact = pick(pool, start, end, field in INSTANT)
            if fact is None:
                continue
            selected[field] = fact
            if field in row:
                row[field] = fact["val"]
            accn = fact.get("accn", "")
            provenance.append(dict(field=field, fiscal_start=start, fiscal_end=end, tag=fact["tag"],
                                   raw_value=fact["val"], unit="shares" if field == "diluted_weighted_avg_shares" else "USD",
                                   filed=fact["filed"], accession=accn, retrieved_at=retrieved_at,
                                   url=f"https://www.sec.gov/Archives/edgar/data/{cik}/{accn.replace('-', '')}/"))
        if all(k in selected for k in debt_tags):
            row["debt"] = sum(selected[k]["val"] for k in debt_tags)
            warnings.append(f"{end}: debt = current long-term debt + noncurrent long-term debt + short-term borrowings; review leases and completeness.")
        else:
            warnings.append(f"{end}: total debt is missing; no absent component was assumed zero.")
        for field in ("revenue", "net_income", "cfo", "capex", "cash", "ebit", "da"):
            if row[field] is None:
                warnings.append(f"{end}: missing {field}; provide manual data if needed.")
        rows.append(row)
    return pd.DataFrame(rows, columns=list(META)+list(FIELDS)), provenance, warnings
