# Resume suggestions and interview preparation

These describe the completed project, including its scope. Use them only once you can explain the implementation. Do not claim a live company analysis, backtest performance, investment returns, time savings or professional experience you have not demonstrated.

## Three resume bullet suggestions

- Built a Python valuation engine that generates a formula-driven Excel DCF with five-year unlevered cash-flow forecasts, an enterprise-to-equity bridge, and Bear/Base/Bull analysis.
- Implemented SEC ticker-to-CIK resolution and conservative annual US-GAAP fact normalization with source provenance, caching, validation and a reviewed CSV fallback for missing data.
- Validated DCF calculations with pytest financial test cases and independent spreadsheet-engine comparisons across 25 WACC/terminal-growth combinations and scenario/input changes.

If asked, explain that AI assisted with coding and describe the financial choices you reviewed yourself. A useful interview demonstration is to change one assumption, predict the direction of the effect, and trace the resulting formula before showing the number.

## 1. Why use FCFF and WACC together?

**Model explanation:** FCFF is operating cash flow available to both lenders and shareholders, before interest and borrowing flows. WACC combines their required returns. Discounting FCFF gives operating enterprise value. I then add reviewed excess cash, subtract debt and adjust for other claims to get common equity value. FCFE is cash available to common shareholders after the relevant financing flows and uses cost of equity instead.

**Project detail:** historical CFO less capex appears separately and never feeds the FCFF forecast. My forecast uses EBIT, operating cash taxes, D&A, capex and changes in operating NWC.

## 2. How did you calculate WACC, and why is it uncertain?

**Model explanation:** cost of equity is risk-free rate plus levered beta times the equity risk premium. I weight that by the market-value equity share, then add pre-tax cost of debt times one minus the tax rate, weighted by debt's market-value share. WACC is an estimate, so all assumptions are editable and each scenario can use a spread or explicit override.

**Project detail:** the fictional example uses 4% risk-free, 1.1 beta, 5% equity premium, 5.5% debt cost, 20% debt weight and 25% tax. Cost of equity is 9.5%. WACC is `9.5%×80% + 5.5%×75%×20% = 8.425%`. I do not call those current market estimates.

## 3. Why not multiply year-five FCFF by one plus terminal growth?

**Model explanation:** that shortcut assumes the final explicit year is already in a stable state. My final explicit sales growth may differ from terminal growth, which changes the cash needed for working capital. I rebuild year-six revenue at terminal growth, keep year-five operating ratios, calculate the new NWC balance, subtract year-five NWC, then calculate terminal FCFF.

**Project detail:** a unit test uses revenue 100, margin 20%, 25% tax, D&A 3%, capex 4% and NWC 10%. Explicit FCFF is 14 with no growth. At terminal growth 2%, revenue becomes 102 and NWC investment is 0.2. FCFF becomes 14.08, not `14×1.02=14.28`.

## 4. What was difficult about SEC financial data?

**Model explanation:** the same metric can have several concept tags, duplicate contexts and restated comparative values. Filing fiscal-year labels can refer to the filing year rather than the observation's economic period. I anchor actual start/end dates to annual revenue, reject quarterly/YTD values and conflicts, match balance-sheet instants to fiscal end, and retain the accession and filing date for review.

**Project detail:** I do not infer operating NWC from a generic current-assets/current-liabilities subtraction, and I do not treat an absent debt component as zero. The SEC pipeline has fixture and mocked-HTTP tests. A complete live end-to-end company download was not demonstrated in this build, so I would not claim one without running and reviewing it myself.

## 5. How do you know the Excel workbook is correct?

**Model explanation:** I test the finance math with manually verifiable cases, then calculate the same model in a separate spreadsheet engine. I compare enterprise value, equity value, share value, case changes and all 25 sensitivity combinations to Python. I also test missing inputs, invalid growth/discount combinations and negative cash flows, and visually inspect all tabs.

**Project detail:** openpyxl does not calculate formulas. New exports need Excel or LibreOffice recalculation, then the included verifier checks the saved results against Python. The bundled example was recalculated using a compatible engine. Desktop Excel automation could not run in the build environment, so I distinguish that check from a native Excel test.

## A two-minute demo

1. Identify the fictional source data, fiscal dates and units.
2. Explain the Base FCFF build from EBIT to cash flow.
3. Change the selector from Base to Bear and predict the direction.
4. Raise WACC and show the lower present value.
5. Delete the diluted share input and show that equity value survives while per-share value becomes unavailable.
6. Set terminal growth equal to WACC and explain why perpetual-growth valuation becomes invalid.
