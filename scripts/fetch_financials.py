import yfinance as yf
import json
from pathlib import Path

RAW_DIR = Path("data/raw")
RAW_DIR.mkdir(parents=True, exist_ok=True)


def fetch_all(ticker: str = "MSFT"):

    print(f"Downloading data for {ticker}")
    co = yf.Ticker(ticker)

    # Income statement
    inc = co.financials.T
    inc.index = inc.index.astype(str)
    out = RAW_DIR / f"{ticker}_income.csv"
    inc.to_csv(out)
    print(f"Saved income statement to {out}")

    # Balance sheet
    bs = co.balance_sheet.T
    bs.index = bs.index.astype(str)
    out = RAW_DIR / f"{ticker}_balance.csv"
    bs.to_csv(out)
    print(f"Saved balance sheet to {out}")

    # Cash flow statement
    cf = co.cashflow.T
    cf.index = cf.index.astype(str)
    out = RAW_DIR / f"{ticker}_cashflow.csv"
    cf.to_csv(out)
    print(f"Saved cash flow statement to {out}")

    # Company profile
    info = co.info
    out = RAW_DIR / f"{ticker}_profile.json"
    with open(out, "w") as f:
        json.dump(info, f, indent=2)
    print(f"Saved company profile to {out}")
    print(f"Finished saving raw files")


if __name__ == "__main__":
    fetch_all("MSFT")