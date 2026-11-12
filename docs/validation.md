# Validation record

Build date: November 12, 2026. Example: fictional Illustrative Industrial Co., retrospective valuation date December 31, 2025.

## Automated financial and pipeline tests

`python -m pytest -q` passed **32 tests** in the build environment. Milestone runs progressed from 10 financial tests to 20 financial/data tests, 29 full-pipeline tests, then 32 including HTTP/cache behavior.

Tests cover:

- A hand-checkable perpetual 10 cash flow at 10% WACC and zero growth: enterprise value 100, equity 85, share value 8.50.
- Rebuilding year-six working-capital investment when terminal growth differs from explicit growth.
- Negative explicit cash flows and no immediate operating-tax refund for losses.
- Missing inputs, zero starting revenue, zero shares, nonpositive terminal FCFF, and invalid WACC/g combinations.
- Dollar/million normalization, missing values, duplicate/overlapping/gapped fiscal dates and invalid signs/units.
- SEC annual-versus-quarterly selection, restatements, conflicting observations, missing debt and NWC, User-Agent validation.
- Cache reuse with original retrieval time, rate-limit retry behavior and an unretried 403 failure.
- End-to-end offline CLI output, formula structure, formatting, charts, review gates and missing-share suppression.

## Independent spreadsheet-engine checks

The generated openpyxl workbook was imported into the compatible `@oai/artifact-tool` spreadsheet engine, recalculated and compared with independent Python results. **44 comparisons/input-change checks passed**, comprising:

- Active-case enterprise value, equity value and share price for Bear, Base and Bull: 9 checks.
- Every combination in the 5×5 WACC/terminal-growth grid: 25 checks, including the center matching the Base DCF.
- Blank unselected assumption leaves Base value intact; the affected Bear comparison becomes unavailable.
- Selecting the case with a blank driver makes its DCF unavailable.
- Missing shares suppress share value and preserve equity value; zero shares also suppress share value.
- Terminal growth equal to WACC suppresses valuation.
- Negative first-year EBIT remains in valuation with no assumed tax refund.
- An invalid case number suppresses valuation.
- Restoring all inputs reproduces the Base value.

Saved-cache checks additionally compare Base enterprise/equity/share values and terminal FCFF, all three scenario comparison outputs, and the sensitivity center against Python. Formula results were supplied by actual spreadsheet-engine calculation, not by copying Python valuation numbers into formula cells. The original openpyxl formulas, styles, data validation and native charts were preserved.

Base result before presentation rounding:

| Metric | Python / spreadsheet result |
|---|---:|
| WACC | 0.08425000000000002 |
| Year-six FCFF ($m) | 168.03978282000003 |
| PV of explicit FCFF ($m) | 546.1273760430206 |
| PV of terminal value ($m) | 1892.677428031556 |
| Enterprise value ($m) | 2438.8048040745766 |
| Equity value ($m) | 2378.8048040745766 |
| Share value ($) | 47.576096081491535 |

All seven sheets were rendered and reviewed. The Summary charts and sensitivity table were checked visually. No unexpected formula errors were found in the recalculated workbook. The formula-based sensitivity is not a native Excel What-If Data Table.

## Practical limits of this verification

Desktop Excel is installed in the environment, but its COM calculation automation returned access denied. No native Excel calculation or LibreOffice run is claimed. The supplied compatible-engine caches help previews display values; the workbook is still marked for automatic full recalculation on opening. New openpyxl exports have no calculated caches until a spreadsheet engine opens and saves them.

The SEC endpoint was reachable during a connectivity probe, but no complete live company download/normalization/filing reconciliation was performed because a real SEC contact identity was not provided. HTTP behavior and normalization were tested using mocked responses and clearly synthetic fixtures. A reviewed real-company use case remains work for the project's owner before claiming completed company research.

The sample is an illustrative workflow, not a linked three-statement model. Accounting reconciliation, capital-structure review and economically sustainable terminal assumptions remain analyst responsibilities.
