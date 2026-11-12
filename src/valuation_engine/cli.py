"""Small command line interface; data download and valuation are separate steps."""
from __future__ import annotations

import argparse
from dataclasses import asdict
from datetime import datetime, timezone
import json
import logging
import os
from pathlib import Path
import sys

from .data import historical_metrics, load_history, validate_config
from .excel import make_workbook
from .model import Drivers, ModelError, cost_of_capital, value

LOG = logging.getLogger(__name__)


def run_valuation(history_path: str | Path, assumptions_path: str | Path, output: str | Path,
                  provenance_path: str | Path | None = None) -> dict:
    history = load_history(history_path)
    config = json.loads(Path(assumptions_path).read_text(encoding="utf-8"))
    bridge = validate_config(config, history)
    base = history.iloc[-1]
    results = {name:value(float(base.revenue),float(base.operating_nwc),Drivers.from_config(case,config["capital"]),bridge)
               for name,case in config["scenarios"].items()}
    output=Path(output)
    output.mkdir(parents=True,exist_ok=True)
    provenance = json.loads(Path(provenance_path).read_text(encoding="utf-8")) if provenance_path else None
    make_workbook(history,config,results,output/"valuation.xlsx",provenance)
    metrics=historical_metrics(history)
    metrics.to_csv(output/"normalized_history.csv",index=False)
    equity_cost,wacc=cost_of_capital(config["capital"])
    result=dict(company=config["company"],ticker=config["ticker"],illustrative=bool(config.get("illustrative")),
                valuation_date=config["valuation_date"],generated_at=datetime.now(timezone.utc).isoformat(),
                units="USD millions; shares in millions; per-share value in USD",cost_of_equity=equity_cost,
                calculated_wacc=wacc,recalculation="Not recalculated by openpyxl. Open and recalculate in Excel or LibreOffice; use verify_cached.py after saving.",
                scenarios={name:asdict(values[1]) for name,values in results.items()})
    (output/"valuation.json").write_text(json.dumps(result,indent=2,allow_nan=False),encoding="utf-8")
    for name,(forecast,_) in results.items():forecast.to_csv(output/f"forecast_{name.lower()}.csv",index=False)
    base_result=results["Base"][1]
    price="Unavailable: no reviewed point-in-time diluted share count." if base_result.per_share is None else f"${base_result.per_share:,.2f} per share (assumption-derived)"
    summary=f"""# {config['company']}

{'ILLUSTRATIVE TRAINING DATA. Fictional company and hypothetical assumptions.' if config.get('illustrative') else 'Reviewed CSV inputs; forecasts and WACC are analyst assumptions.'}

Retrospective valuation date: {config['valuation_date']}. Generated: {result['generated_at']}.

Base enterprise value: ${base_result.enterprise_value:,.2f} million. Equity value: ${base_result.equity_value:,.2f} million. {price}.

| Scenario | Enterprise value ($m) | Equity value ($m) | Value / share ($) |
|---|---:|---:|---:|
"""
    for name in ("Bear","Base","Bull"):
        v=results[name][1]
        summary+=f"| {name} | {v.enterprise_value:,.2f} | {v.equity_value:,.2f} | {'n.a.' if v.per_share is None else f'{v.per_share:,.2f}'} |\n"
    summary+="\nFCFF = EBIT − operating cash taxes + D&A − capex − increase in operating NWC. Cash taxes use MAX(EBIT,0) × tax rate. Explicit cash flows are discounted at year-end. Stable year-6 FCFF is rebuilt using terminal growth and year-5 operating ratios.\n\n"
    summary+="Equity = enterprise value + reviewed excess cash − reviewed debt + signed other adjustments. Shares are a reviewed point-in-time input; historical weighted-average EPS shares are not substituted.\n\n"
    summary+="Key limitations: no NOL, full PP&E, dilution or stub-period schedules. Terminal reinvestment and long-run ROIC require analyst judgment. This is not a current market price target. Recalculate the workbook in Excel/LibreOffice before relying on displayed formula results.\n\n"
    summary+="Input review: "+config["review_note"]+"\n\nAssumptions: "+config["assumption_source"]+"\n\nBridge: "+config["bridge"]["source"]+"\n"
    (output/"summary.md").write_text(summary,encoding="utf-8")
    LOG.info("Generated %s; Base enterprise value %.2f USD m",output/"valuation.xlsx",base_result.enterprise_value)
    return result


def fetch_sec(args) -> None:
    from .sec import SecClient, SecError, normalize_facts
    user_agent=os.getenv("SEC_USER_AGENT","")
    client=SecClient(user_agent,args.cache)
    facts,retrieved,cik,company=client.fetch(args.ticker)
    frame,provenance,warnings=normalize_facts(facts,retrieved,args.years)
    output=Path(args.out);output.mkdir(parents=True,exist_ok=True)
    frame.to_csv(output/"history_unreviewed.csv",index=False)
    (output/"companyfacts_raw.json").write_text(json.dumps(facts),encoding="utf-8")
    (output/"provenance.json").write_text(json.dumps(provenance,indent=2),encoding="utf-8")
    report=dict(ticker=args.ticker.upper(),cik=cik,company=company,retrieved_at=retrieved,warnings=warnings,
                next_step="Review tags and annual filings, fill blank operating_nwc/debt/other inputs with sources, copy assumptions template and set reviewed=true only after review.")
    (output/"review_required.json").write_text(json.dumps(report,indent=2),encoding="utf-8")
    print(json.dumps(report,indent=2))


def main(argv: list[str] | None = None) -> int:
    parser=argparse.ArgumentParser(description="Auditable SEC/manual FCFF valuation engine")
    parser.add_argument("--verbose",action="store_true")
    subs=parser.add_subparsers(dest="command",required=True)
    fetch=subs.add_parser("fetch-sec",help="Download unreviewed SEC history and provenance")
    fetch.add_argument("ticker");fetch.add_argument("--years",type=int,default=3)
    fetch.add_argument("--out",default="outputs/sec");fetch.add_argument("--cache",default=".cache/sec")
    val=subs.add_parser("value",help="Generate Python outputs and a formula-driven Excel DCF")
    val.add_argument("--history",required=True);val.add_argument("--assumptions",required=True)
    val.add_argument("--out",default="outputs/valuation");val.add_argument("--provenance")
    args=parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,format="%(levelname)s %(name)s: %(message)s")
    try:
        if args.command=="fetch-sec":fetch_sec(args)
        else:
            result=run_valuation(args.history,args.assumptions,args.out,args.provenance)
            print(json.dumps(result,indent=2))
    except (ModelError,ValueError,KeyError,OSError) as exc:
        LOG.error("%s",exc);return 2
    except RuntimeError as exc:
        LOG.error("%s",exc);return 2
    return 0
