# Microsoft (MSFT) DCF Valuation

A 5-year DCF on Microsoft. Python pulls and cleans the financials (yfinance, pandas, openpyxl), and the model itself is in Excel. I built it in April 2026 and updated it on 10/26/2026 after FY2026 results came out.

## Result

As of 10/26/2026 I get $395.95 per share against a market price of $532.60, about 26% downside. To justify $532.60, you'd need a discount rate around 7.9% instead of my 9.45%, or AI capex that pays off a lot faster than I assume.

## Assumptions

- WACC of 9.45% from CAPM: 4.99% 10Y Treasury, 4.18% ERP, 1.099 beta, and a 2.6% debt weight
- Revenue growth of 17% in FY2027, slowing to 11% by FY2031
- Capex at 30% of revenue in FY2027, falling to 11% by FY2031 as the data center buildout winds down
- Margins and tax rate from FY2026; D&A as a 3-year average
- 3% terminal growth (Gordon Growth)
- Debt, cash, and investments from the FY2026 10-K, with $66.6B of finance leases counted as debt

The workbook has Assumptions, Forecast, DCF, and Sensitivity tabs. The sensitivity grid runs WACC from 8.0% to 10.5% against terminal growth from 2% to 4%.

## What changed since April

The original model is in `excel/archive/MSFT_DCF_Model_2026-04.xlsx` (tag `v1-apr2026`). Here's how its FY2026 forecast held up:

| FY2026 | April forecast | Actual |
|---|---|---|
| Revenue growth | 16.7% | 17.8% |
| EBIT margin | 45.6% | 46.8% |
| Capex / revenue | 30.0% | 34.9% |

Revenue and margins came in a bit better than I expected. Capex didn't. It ran well above my estimate (39.8% of revenue in Q4 alone), so I pushed the capex glide path back a year.

Going from $380.14 to $395.95:

| Step | Per share |
|---|---|
| April model | $380.14 |
| Fixed the terminal value, which grew the last year's FCFF by (1+g) twice | $370.78 |
| Counted short-term investments as cash | $379.44 |
| Added a stub period (April was 72 days before fiscal year end) | $399.96 |
| Rolled forward to FY2026 actuals and current inputs | $396.28 |
| Took debt, cash, and investments from the 10-K | $395.95 |

Rates did most of the damage. The 10Y went from 4.34% to 4.99%, which pushed WACC from 8.80% to 9.49%. At the old WACC the updated model would give $446.87.

## Where I could be wrong

| If instead... | Per share |
|---|---|
| New finance lease assets count as capex ($24.6B in FY2026) | ~$361 |
| Mid-year discounting | ~$399 |
| The OpenAI stake is valued at ~$135B (Oct 2025 recapitalization), not book | ~$409 |
| A 10-year forecast fades growth from 17% to 6% first | ~$475 |
| ERP is 4.0% / 4.5% | ~$409 / ~$374 |

The weakest spot is D&A. I hold it at 10.2% of revenue, but the assets Microsoft has already built point to more like 14 to 16%. That doesn't fit well with a flat EBIT margin, or with terminal capex (11%) sitting below D&A. The terminal value is also 82% of enterprise value and works out to about 10x FY2031 EBITDA, while the stock trades near 18x forward EBITDA. Finally, the balance sheet is from 6/30/2026, so the Q1 FY2027 10-Q should be folded in once it's out.

## Data

`scripts/fetch_financials.py` pulls statements from yfinance, `clean_data.py` builds `data/processed/MSFT_historicals.csv`, and `export_to_excel.py` writes the Historicals and Profile tabs. Yahoo has since revised FY2024 and FY2025 D&A (FY2025 went from $34.2B to $29.4B), and I use the revised numbers so every year is on the same basis.
