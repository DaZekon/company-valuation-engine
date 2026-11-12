from datetime import datetime, timezone
from unittest.mock import Mock

import pytest
import requests

from valuation_engine.sec import SecClient, SecError


def test_cache_preserves_original_retrieval_date_and_no_network(tmp_path, monkeypatch):
    client=SecClient("AutomatedValuation contact@example.org",tmp_path)
    response=Mock(status_code=200)
    response.json.return_value={"example":1}
    client.session.get=Mock(return_value=response)
    first,timestamp=client.get_json("https://data.sec.gov/example")
    second,cached_timestamp=client.get_json("https://data.sec.gov/example")
    assert first==second=={"example":1}
    assert timestamp==cached_timestamp
    assert client.session.get.call_count==1


def test_403_fails_without_bypass(tmp_path):
    client=SecClient("AutomatedValuation contact@example.org",tmp_path)
    response=Mock(status_code=403)
    response.raise_for_status.side_effect=requests.HTTPError("403 denied")
    client.session.get=Mock(return_value=response)
    with pytest.raises(SecError,match="CSV fallback"):
        client.get_json("https://data.sec.gov/example")
    assert client.session.get.call_count==1


def test_rate_limit_retry_and_two_requests_per_second(tmp_path,monkeypatch):
    waits=[]
    monkeypatch.setattr("valuation_engine.sec.time.sleep",lambda seconds:waits.append(seconds))
    limited=Mock(status_code=429,headers={"Retry-After":"2"})
    success=Mock(status_code=200);success.json.return_value={"ok":True}
    client=SecClient("AutomatedValuation contact@example.org",tmp_path)
    client.session.get=Mock(side_effect=[limited,success])
    assert client.get_json("https://data.sec.gov/example")[0]=={"ok":True}
    assert 2 in waits
    assert len(waits)==3
