from copy import deepcopy

import pandas as pd
import pytest

from valuation_engine.data import FIELDS, load_history
from valuation_engine.model import ModelError
from valuation_engine.sec import SecClient, candidates, normalize_facts


@pytest.fixture
def csv_frame():
    return pd.DataFrame([dict(fiscal_start=f"{y}-01-01", fiscal_end=f"{y}-12-31", currency="USD",
                             unit_scale=1000000, source="Synthetic test fixture", retrieved_at="2026-11-12",
                             **{field:100 for field in FIELDS}) for y in (2023,2024,2025)])


def test_units_and_missing_values(tmp_path, csv_frame):
    csv_frame.loc[0,"ebit"] = None
    csv_frame.to_csv(tmp_path/"h.csv",index=False)
    result = load_history(tmp_path/"h.csv")
    assert result.revenue.iloc[0] == 100
    assert pd.isna(result.ebit.iloc[0])
    csv_frame["unit_scale"] = 1
    csv_frame.to_csv(tmp_path/"h.csv",index=False)
    assert load_history(tmp_path/"h.csv").revenue.iloc[0] == .0001


@pytest.mark.parametrize("field,value", [("fiscal_start","2024-02-01"),("fiscal_end","2025-12-31"),("currency","EUR"),("revenue","invalid"),("capex",-1),("unit_scale",100)])
def test_invalid_periods_and_values(tmp_path, csv_frame, field, value):
    csv_frame[field] = csv_frame[field].astype(object)
    csv_frame.loc[1,field] = value
    csv_frame.to_csv(tmp_path/"h.csv",index=False)
    with pytest.raises(ModelError):
        load_history(tmp_path/"h.csv")


def make_facts():
    facts = []
    for y in (2023,2024,2025):
        facts.append(dict(start=f"{y}-01-01",end=f"{y}-12-31",val=y*100,form="10-K",filed="2026-02-01",accn="0000000001-26-000001"))
    return {"cik":1,"facts":{"us-gaap":{"Revenues":{"units":{"USD":facts}}}}}


def test_sec_no_quarterly_or_fake_missing_values():
    facts = make_facts()
    facts["facts"]["us-gaap"]["Revenues"]["units"]["USD"].append(dict(start="2025-10-01",end="2025-12-31",val=999,form="10-K",filed="2026-02-01"))
    frame, sources, warnings = normalize_facts(facts,"2026-11-12T00:00:00+00:00")
    assert len(frame) == 3
    assert frame.revenue.iloc[-1] == 202500
    assert pd.isna(frame.debt.iloc[-1]) and pd.isna(frame.operating_nwc.iloc[-1])
    assert sources[-1]["tag"] == "Revenues"
    assert warnings


def test_restatement_and_conflicting_values():
    facts = make_facts()
    unit = facts["facts"]["us-gaap"]["Revenues"]["units"]["USD"]
    revised = {**unit[-1], "filed":"2026-03-01", "val":999}
    unit.append(revised)
    assert normalize_facts(facts,"2026-11-12")[0].revenue.iloc[-1] == 999
    unit.append({**revised,"val":1000})
    with pytest.raises(ModelError,match="Conflicting"):
        normalize_facts(facts,"2026-11-12")


def test_user_agent_required(tmp_path):
    with pytest.raises(ModelError,match="SEC_USER_AGENT"):
        SecClient("anonymous",tmp_path)
