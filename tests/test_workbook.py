from copy import deepcopy
import json
from pathlib import Path

from openpyxl import load_workbook
import pytest

from valuation_engine.cli import main, run_valuation
from valuation_engine.data import load_history, validate_config
from valuation_engine.model import ModelError

ROOT=Path(__file__).resolve().parents[1]
SAMPLE=ROOT/"sample_data"


def test_full_pipeline_and_excel_formulas(tmp_path):
    results=run_valuation(SAMPLE/"illustrative_history.csv",SAMPLE/"illustrative_assumptions.json",tmp_path)
    assert results["scenarios"]["Bear"]["enterprise_value"] < results["scenarios"]["Base"]["enterprise_value"] < results["scenarios"]["Bull"]["enterprise_value"]
    wb=load_workbook(tmp_path/"valuation.xlsx",data_only=False)
    assert wb.sheetnames==["Summary","Inputs","Historical","Projection","DCF","Sensitivity","Sources"]
    assert wb["DCF"]["C16"].data_type=="f"
    assert "C6-C7" in wb["DCF"]["C13"].value
    assert wb["Projection"]["J20"].value=='=IF(COUNT(J19,H19)=2,J19-H19,"n.a.")'
    assert len(wb["Summary"]._charts)==2
    assert len(wb["Inputs"].data_validations.dataValidation)>10
    assert wb["Inputs"]["C10"].font.color.rgb=="000000FF"
    assert all("#REF!" not in str(c.value) for ws in wb for row in ws for c in row)
    assert not wb._external_links


def test_missing_shares_still_values_equity(tmp_path):
    cfg=json.loads((SAMPLE/"illustrative_assumptions.json").read_text())
    cfg["diluted_shares"]=None
    (tmp_path/"assumptions.json").write_text(json.dumps(cfg))
    results=run_valuation(SAMPLE/"illustrative_history.csv",tmp_path/"assumptions.json",tmp_path/"out")
    assert results["scenarios"]["Base"]["per_share"] is None
    assert results["scenarios"]["Base"]["equity_value"]>0


@pytest.mark.parametrize("change",[dict(reviewed=False),dict(valuation_date="2026-11-12"),dict(shares_as_of="2026-01-01"),dict(industry="bank"),dict(diluted_shares=0)])
def test_review_gates(change):
    cfg=json.loads((SAMPLE/"illustrative_assumptions.json").read_text())
    cfg.update(change)
    with pytest.raises(ModelError):validate_config(cfg,load_history(SAMPLE/"illustrative_history.csv"))


def test_manual_data_gap_is_specific():
    cfg=json.loads((SAMPLE/"illustrative_assumptions.json").read_text())
    history=load_history(SAMPLE/"illustrative_history.csv")
    history.loc[len(history)-1,"operating_nwc"]=float("nan")
    with pytest.raises(ModelError,match="operating_nwc"):
        validate_config(cfg,history)


def test_bad_path_returns_actionable_cli_error():
    assert main(["value","--history","does_not_exist.csv","--assumptions","none.json"])==2
