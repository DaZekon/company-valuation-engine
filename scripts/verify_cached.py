"""Compare saved Excel formula caches to independently calculated Python results.

Run after recalculating/saving valuation.xlsx in Excel or LibreOffice.
Exit 2 means caches absent, errors present, or a mismatch. Never writes caches.
"""
from __future__ import annotations

import argparse
import json
from math import isclose
from pathlib import Path

from openpyxl import load_workbook


def verify(workbook_path: str | Path, json_path: str | Path) -> dict:
    wb=load_workbook(workbook_path,data_only=True)
    expected=json.loads(Path(json_path).read_text(encoding="utf-8"))
    checks=[]
    def check(sheet,cell,target):
        actual=wb[sheet][cell].value
        passed=(actual=="n.a.") if target is None else isinstance(actual,(int,float)) and isclose(actual,target,rel_tol=1e-9,abs_tol=1e-7)
        checks.append(dict(sheet=sheet,cell=cell,actual=actual,expected=target,passed=passed))
    for cell,key in (("C16","enterprise_value"),("C22","equity_value"),("C25","per_share"),("C8","terminal_fcff")):
        check("DCF",cell,expected["scenarios"]["Base"][key])
    for col,name in zip("DEF",("Bear","Base","Bull")):
        for row,key in ((16,"enterprise_value"),(17,"equity_value"),(18,"per_share")):
            check("Sensitivity",f"{col}{row}",expected["scenarios"][name][key])
    check("Sensitivity","F9",expected["scenarios"]["Base"]["enterprise_value"])
    errors=[f"{ws.title}!{cell.coordinate}: {cell.value}" for ws in wb for row in ws for cell in row if cell.data_type=="e"]
    result=dict(passed=all(c["passed"] for c in checks) and not errors,checks=checks,formula_errors=errors)
    if not result["passed"]:
        raise ValueError(json.dumps(result,indent=2)+"\nRecalculate and save in Excel/LibreOffice; check inputs if caches differ.")
    return result


if __name__=="__main__":
    parser=argparse.ArgumentParser();parser.add_argument("workbook");parser.add_argument("results_json")
    args=parser.parse_args()
    try:print(json.dumps(verify(args.workbook,args.results_json),indent=2))
    except ValueError as exc:parser.exit(2,str(exc)+"\n")
