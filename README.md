# Microsoft (MSFT) DCF Valuation & Sensitivity Analysis

Built a 5-year discounted cash flow model on Microsoft (MSFT) using Python-pulled financials via yfinance and a fully structured Excel model. Assumptions are sourced from Alpha Spread, yfinance, CNBC earnings reports, and CFO commentary from the Q1 FY2026 earnings call. First built in April 2026; refreshed on 10/26/2026 with FY2026 actuals (see [What Changed Since April](#what-changed-since-april)).

## Results (as of 10/26/2026)
- **Intrinsic Value Per Share:** $396.28
- **Current Market Price:** $532.60
- **Implied Downside:** ~26%
- **Conclusion:** MSFT trades well above fair value under base case assumptions. At $532.60 the market is pricing in either a ~7.9% discount rate (vs. 9.49% in the model) or a much faster payoff from AI capex than the model assumes.

## Model Structure
- **Assumptions Tab:** All inputs in one place — revenue growth rates, margin assumptions, WACC inputs (risk-free rate, ERP, beta, cost of debt), terminal growth rate, and valuation timing, each with a cited source
- **Forecast Tab:** FY2022–FY2026 historicals and FY2027E–FY2031E projections, including a full FCFF build from NOPAT through CapEx and NWC changes
- **DCF Tab:** Discounted FCFF with a stub period for the part of FY2027 still ahead of the valuation date, terminal value via Gordon Growth Model, enterprise value bridge, and implied intrinsic value per share
- **Sensitivity Tab:** 2-way sensitivity analysis across WACC (8.0%–10.5%) and terminal growth rate (2.0%–4.0%), calculated live from the DCF tab, with conditional formatting vs. current market price

## Key Assumptions
- WACC of 9.49% derived via CAPM using a risk-free rate of 4.99% (10Y Treasury), ERP of 4.18%, and beta of 1.099, with a market-value debt weight of 1.4%
- Terminal growth rate of 3.00% anchored to long-run nominal GDP
- CapEx modeled as a glide path from 30% of revenue in FY2027 down to 11% by FY2031, reflecting peak AI infrastructure spending normalizing as data center buildout matures
- Revenue growth tapering from 17% in FY2027 to 11% in FY2031, reflecting AI-driven cloud acceleration gradually normalizing
- Margins, tax rate, and D&A taken from FY2026 actuals (D&A as a 3-year average, FY2024–FY2026)

## What Changed Since April
The original April 2026 model is kept at [`excel/archive/MSFT_DCF_Model_2026-04.xlsx`](excel/archive/MSFT_DCF_Model_2026-04.xlsx) (git tag `v1-apr2026`).

**April forecast vs. FY2026 actuals**

| FY2026 | April forecast | Actual |
|---|---|---|
| Revenue growth | 16.7% | 17.8% |
| EBIT margin | 45.6% | 46.8% |
| CapEx / revenue | 30.0% | 34.9% |

Revenue and margins came in slightly ahead of the forecast. CapEx ran well above it (Q4 FY2026 alone was 39.8% of revenue), so the April glide path is shifted one year later rather than kept on its original schedule.

**Bridge from $380.14 to $396.28**

| Step | Value / share |
|---|---|
| April model as published | $380.14 |
| Fix: terminal value grew the final-year FCFF by (1+g) twice | $370.78 |
| Fix: count short-term investments as cash in the EV bridge | $379.44 |
| Fix: stub period (April valuation date was 72 days before FY2026 year end) | $399.96 |
| Roll forward to FY2026 actuals, FY2027E–FY2031E, 10/26/2026 inputs | $396.28 |

Most of the roll-forward effect is the discount rate: the 10Y Treasury rose from 4.34% to 4.99%, taking WACC from 8.80% to 9.49%. At the April WACC the refreshed model gives $446.87.

## Data Pipeline
1. `scripts/fetch_financials.py` pulls statements and the company profile from yfinance into `data/raw/`
2. `scripts/clean_data.py` normalizes them into `data/processed/MSFT_historicals.csv`
3. `scripts/export_to_excel.py` writes the Historicals and Profile tabs

Yahoo has since revised FY2024–FY2025 D&A to exclude "other" items (FY2025: $34.2B → $29.4B); the historicals use the revised series so all years are on one definition.

## Tools
- Python (yfinance, pandas, openpyxl)
- Microsoft Excel
