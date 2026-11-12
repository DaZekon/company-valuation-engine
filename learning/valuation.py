from math import isfinite


def calculate_fcff(
    ebit, tax_rate, depreciation, capex, change_in_nwc
):
    """Calculate cash available to lenders and shareholders."""
    inputs = (
        ebit, tax_rate, depreciation, capex, change_in_nwc
    )

    if not all(isfinite(x) for x in inputs):
        raise ValueError("All inputs must be finite numbers.")

    if not 0 <= tax_rate <= 1:
        raise ValueError("Tax rate must be between 0 and 1.")

    if depreciation < 0 or capex < 0:
        raise ValueError("Depreciation and capex cannot be negative.")

    # Losses receive no immediate cash tax refund in this model.
    cash_taxes = max(ebit, 0) * tax_rate

    return (
        ebit
        - cash_taxes
        + depreciation
        - capex
        - change_in_nwc
    )


def calculate_enterprise_value(
    fcffs, wacc, terminal_growth, terminal_fcff
):
    """Discount five annual FCFFs and a perpetual terminal value."""
    if len(fcffs) != 5:
        raise ValueError("Provide exactly five annual cash flows.")

    inputs = [*fcffs, wacc, terminal_growth, terminal_fcff]

    if not all(isfinite(x) for x in inputs):
        raise ValueError("All inputs must be finite numbers.")

    if not 0 < wacc <= 1:
        raise ValueError("WACC must be above zero and at most 100%.")

    if not -1 < terminal_growth < wacc:
        raise ValueError(
            "Terminal growth must exceed -100% and be below WACC."
        )

    if terminal_fcff <= 0:
        raise ValueError(
            "This method requires positive terminal-year FCFF."
        )

    pv_cash_flows = sum(
        cash_flow / (1 + wacc) ** year
        for year, cash_flow in enumerate(fcffs, start=1)
    )

    terminal_value = terminal_fcff / (wacc - terminal_growth)
    pv_terminal_value = terminal_value / (1 + wacc) ** 5

    return pv_cash_flows + pv_terminal_value


if __name__ == "__main__":
    # Illustrative training numbers, all amounts in USD millions.
    example_fcff = calculate_fcff(
        ebit=20,
        tax_rate=0.25,
        depreciation=3,
        capex=4,
        change_in_nwc=2,
    )

    example_value = calculate_enterprise_value(
        fcffs=[10, 10, 10, 10, 10],
        wacc=0.10,
        terminal_growth=0.00,
        terminal_fcff=10,
    )

    print(f"Practice FCFF: ${example_fcff:.2f} million")
    print(f"Practice enterprise value: ${example_value:.2f} million")