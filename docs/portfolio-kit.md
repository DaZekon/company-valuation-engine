# Portfolio kit

## Resume project entry

**Automated Company Valuation Engine | Python, Excel, Financial Modeling**

- Built a Python tool that generates a formula-driven Excel DCF with five-year unlevered cash-flow forecasts and an enterprise-to-equity bridge.
- Developed Bear/Base/Bull scenarios and a 25-combination WACC versus terminal-growth sensitivity analysis with editable operating assumptions.
- Tested valuation mathematics, data validation and Excel generation with pytest; documented SEC fact mappings and a manual-input fallback.

These bullets describe an independent project, not employment. AI assisted with implementation. Be able to run the model and explain the finance before discussing it in an interview. Do not claim investment returns, a fully automated research workflow or completed real-company research from the fictional example.

## LinkedIn Projects entry

Name: Automated Company Valuation Engine

Description: Built an educational Python-to-Excel valuation tool that projects five years of unlevered free cash flow, calculates enterprise and equity value, and compares Bear/Base/Bull assumptions. The model includes a 5x5 WACC/terminal-growth sensitivity grid, source documentation and automated financial tests. The demonstration uses clearly labeled fictional data. SEC importing is provided as a separate review-required workflow.

Skills: Financial Modeling, Discounted Cash Flow, Corporate Finance, Python, Microsoft Excel, Data Analysis, Git

Add the actual published repository URL after publication. Use the workbook preview in docs/summary.png as the project image.

## LinkedIn launch post draft

I built an Automated Company Valuation Engine as an independent finance project.

The tool uses Python to generate an Excel DCF with five-year unlevered cash-flow projections, an enterprise-to-equity bridge, Bear/Base/Bull scenarios, and a WACC versus terminal-growth sensitivity table.

My main focus was making the calculations traceable: separating FCFF from CFO less capex, keeping enterprise and equity value consistent, and testing how missing inputs and invalid assumptions affect the model.

The demonstration uses fictional training data. A separate SEC importer supports a reviewed real-company workflow. I used AI coding assistance and worked through the financial logic and tests as part of the project.

Repository: add the published GitHub link here.

I'd welcome feedback from people working in corporate finance, valuation or investment banking.

## Before adding the link to applications

1. Run run_demo.py and the tests on your computer.
2. Open the workbook and change the case selector from Base to Bear.
3. Explain why a higher WACC lowers present value.
4. Trace how revenue assumptions affect operating NWC and FCFF.
5. Explain why weighted-average EPS shares are not automatically a current diluted share count.
6. Read docs/interview-prep.md and distinguish what the model does from what still needs analyst review.
