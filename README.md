# Microsoft (MSFT) DCF Valuation & Sensitivity Analysis

Built a 5-year discounted cash flow model on Microsoft (MSFT) using Python-pulled financials via yfinance and a fully structured Excel model. Assumptions are sourced from Alpha Spread, yfinance, CNBC earnings reports, and CFO commentary from the Q1 FY2026 earnings call.

## Results
- **Intrinsic Value Per Share:** $380.14
- **Current Market Price:** $424.16
- **Implied Downside:** ~10%
- **Conclusion:** MSFT appears to be trading at a modest premium to fair value under base case assumptions. The market appears to be pricing in a more optimistic AI monetization scenario than the model assumes.

## Model Structure
- **Assumptions Tab:** All inputs in one place — revenue growth rates, margin assumptions, WACC inputs (risk-free rate, ERP, beta, cost of debt), and terminal growth rate, each with a cited source
- **Forecast Tab:** FY2021–FY2025 historicals and FY2026E–FY2030E projections, including a full FCFF build from NOPAT through CapEx and NWC changes
- **DCF Tab:** Discounted FCFF, terminal value via Gordon Growth Model, enterprise value bridge, and implied intrinsic value per share
- **Sensitivity Tab:** 2-way sensitivity analysis across WACC (7.5%–10.0%) and terminal growth rate (2.0%–4.0%) with conditional formatting vs. current market price

## Key Assumptions
- WACC of 8.80% derived via CAPM using a risk-free rate of 4.34%, ERP of 4.18%, and beta of 1.107
- Terminal growth rate of 3.00% anchored to long-run nominal GDP
- CapEx modeled as a glide path from 30% of revenue in FY2026 down to 11% by FY2030, reflecting peak AI infrastructure spending normalizing as data center buildout matures
- Revenue growth tapering from 16.7% in FY2026 to 12% in FY2030, reflecting AI-driven cloud acceleration gradually normalizing

## Tools
- Python (yfinance, pandas)
- Microsoft Excel