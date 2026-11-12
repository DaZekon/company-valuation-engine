import pytest

from valuation import (
    calculate_fcff,
    calculate_enterprise_value,
)


def test_fcff_calculation():
    result = calculate_fcff(20, 0.25, 3, 4, 2)
    assert result == pytest.approx(12)


def test_losses_receive_no_immediate_tax_refund():
    result = calculate_fcff(-10, 0.25, 3, 4, 2)
    assert result == pytest.approx(-13)


def test_hand_calculated_enterprise_value():
    result = calculate_enterprise_value(
        [10, 10, 10, 10, 10], 0.10, 0.00, 10
    )
    assert result == pytest.approx(100)


def test_terminal_growth_cannot_equal_wacc():
    with pytest.raises(ValueError):
        calculate_enterprise_value(
            [10, 10, 10, 10, 10], 0.10, 0.10, 10
        )