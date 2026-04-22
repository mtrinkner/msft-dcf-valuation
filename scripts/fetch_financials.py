import yfinance as yf
import json
from pathlib import Path

RAW_DIR = Path("data/raw")
RAW_DIR.mkdir(parents=True, exist_ok=True)


def fetch_all(ticker: str = "MSFT") -> None:

    print(f"\n=== Fetching financials for {ticker} via yfinance ===\n")
    co = yf.Ticker(ticker)

    inc = co.financials.T
    inc.index = inc.index.astype(str)
    out = RAW_DIR / f"{ticker}_income.csv"
    inc.to_csv(out)
    print(f"  ✓ Income statement  → {out}  ({len(inc)} years)")

    bs = co.balance_sheet.T
    bs.index = bs.index.astype(str)
    out = RAW_DIR / f"{ticker}_balance.csv"
    bs.to_csv(out)
    print(f"  ✓ Balance sheet     → {out}  ({len(bs)} years)")

    cf = co.cashflow.T
    cf.index = cf.index.astype(str)
    out = RAW_DIR / f"{ticker}_cashflow.csv"
    cf.to_csv(out)
    print(f"  ✓ Cash flow         → {out}  ({len(cf)} years)")

    info = co.info
    out = RAW_DIR / f"{ticker}_profile.json"
    with open(out, "w") as f:
        json.dump(info, f, indent=2)
    print(f"  ✓ Profile           → {out}")
    print(f"\n✓ All raw data saved to data/raw/")


if __name__ == "__main__":
    fetch_all("MSFT")