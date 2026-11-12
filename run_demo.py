"""Run the complete offline example without an editable package installation."""
from __future__ import annotations

import argparse
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from valuation_engine.cli import run_valuation


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate the illustrative five-year Excel DCF")
    parser.add_argument("--open", action="store_true", help="Open the generated workbook in the Windows default spreadsheet app")
    args = parser.parse_args()
    output = ROOT / "outputs" / "demo"
    result = run_valuation(
        ROOT / "sample_data" / "illustrative_history.csv",
        ROOT / "sample_data" / "illustrative_assumptions.json",
        output,
    )
    print("\nILLUSTRATIVE TRAINING COMPANY: these are not real security values.")
    print("Amounts in USD millions, except value per share.\n")
    print(f"{'Scenario':<10} {'Enterprise value':>18} {'Equity value':>18} {'Value/share':>15}")
    for name in ("Bear", "Base", "Bull"):
        values = result["scenarios"][name]
        share_text = "n.a." if values["per_share"] is None else f"${values['per_share']:,.2f}"
        print(f"{name:<10} {values['enterprise_value']:>18,.2f} {values['equity_value']:>18,.2f} {share_text:>15}")
    workbook = output / "valuation.xlsx"
    print(f"\nWorkbook: {workbook}")
    print("Open and save in Excel to recalculate formulas. Input cells are blue.")
    if args.open:
        if sys.platform == "win32":
            os.startfile(str(workbook))
        else:
            print("Open the workbook using your spreadsheet application.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
