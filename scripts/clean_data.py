import pandas as pd
from pathlib import Path

RAW_DIR  = Path("data/raw")
PROC_DIR = Path("data/processed")
PROC_DIR.mkdir(parents=True, exist_ok=True)


def _load(filename: str) -> pd.DataFrame:
    return pd.read_csv(RAW_DIR / filename, index_col=0)


def clean_income_statement(ticker: str) -> pd.DataFrame:
    df = _load(f"{ticker}_income.csv").copy()
    df = df.rename(columns={
        "Total Revenue": "revenue",
        "Gross Profit": "gross_profit",
        "Operating Income": "ebit",
        "Net Income": "net_income",
        "EBITDA": "ebitda",
        "Interest Expense": "interest_expense",
        "Tax Provision": "income_tax",
        "Pretax Income": "ebt",
    })
    df.index = pd.to_datetime(df.index)
    df["fiscal_year"] = df.index.year
    keep = [c for c in ["fiscal_year","revenue","gross_profit","ebit",
                          "net_income","ebitda","interest_expense",
                          "income_tax","ebt"] if c in df.columns]
    df = df[keep].sort_values("fiscal_year").reset_index(drop=True)
    df["gross_margin_pct"] = df["gross_profit"] / df["revenue"]
    df["ebit_margin_pct"]  = df["ebit"] / df["revenue"]
    if "income_tax" in df.columns and "ebt" in df.columns:
        df["tax_rate"] = df["income_tax"] / df["ebt"].replace(0, pd.NA)
    # yfinance doesn't provide info
    row_2021 = pd.DataFrame([{
    "fiscal_year": 2021,
    "revenue": 168088000000,
    "gross_profit": 115856000000,
    "ebit": 69916000000,
    "net_income": 61271000000,
    "ebitda": 81869000000,
    "interest_expense": 2346000000,
    "income_tax": 9831000000,
    "ebt": 71102000000,
    "gross_margin_pct": 115856000000 / 168088000000,
    "ebit_margin_pct": 69916000000  / 168088000000,
    "tax_rate": 9831000000   / 71102000000,
}])
    df = df[df["fiscal_year"] != 2021]
    df = pd.concat([row_2021, df], ignore_index=True)
    df = df.sort_values("fiscal_year").reset_index(drop=True)
    return df

def clean_cash_flow(ticker: str) -> pd.DataFrame:
    df = _load(f"{ticker}_cashflow.csv").copy()
    df = df.rename(columns={
        "Depreciation And Amortization": "depreciation_amortization",
        "Capital Expenditure": "capex",
        "Changes In Working Capital": "change_in_nwc",
        "Free Cash Flow": "free_cash_flow",
    })
    df.index = pd.to_datetime(df.index)
    df["fiscal_year"] = df.index.year
    keep = [c for c in ["fiscal_year","depreciation_amortization",
                          "capex","change_in_nwc","free_cash_flow"]
            if c in df.columns]
    df = df[keep].sort_values("fiscal_year").reset_index(drop=True)
    if "capex" in df.columns:
        df["capex"] = df["capex"].abs()
    # yfinance doesn't provide info
    row_2021_cf = pd.DataFrame([{
    "fiscal_year": 2021,
    "depreciation_amortization": 11689000000,
    "capex": 20622000000,
    "change_in_nwc": -5665000000,
    "free_cash_flow": 56118000000,
}])
    df = df[df["fiscal_year"] != 2021]
    df = pd.concat([row_2021_cf, df], ignore_index=True)
    df = df.sort_values("fiscal_year").reset_index(drop=True)
    return df


def clean_balance_sheet(ticker: str) -> pd.DataFrame:
    df = _load(f"{ticker}_balance.csv").copy()
    df = df.rename(columns={
        "Current Assets": "current_assets",
        "Current Liabilities": "current_liabilities",
        "Total Debt": "total_debt",
        "Cash And Cash Equivalents": "cash",
        "Stockholders Equity": "total_equity",
    })
    df.index = pd.to_datetime(df.index)
    df["fiscal_year"] = df.index.year
    keep = [c for c in ["fiscal_year","current_assets","current_liabilities",
                          "total_debt","cash","total_equity"]
            if c in df.columns]
    df = df[keep].sort_values("fiscal_year").reset_index(drop=True)
    if "current_assets" in df.columns and "current_liabilities" in df.columns:
        df["net_working_capital"] = df["current_assets"] - df["current_liabilities"]
    # yfinance doesn't provide info
    row_2021_bs = pd.DataFrame([{
    "fiscal_year": 2021,
    "current_assets": 184406000000,
    "current_liabilities": 88657000000,
    "total_debt": 67775000000,
    "cash": 14224000000,
    "total_equity": 141988000000,
    "net_working_capital": 95749000000,
}])
    df = df[df["fiscal_year"] != 2021]
    df = pd.concat([row_2021_bs, df], ignore_index=True)
    df = df.sort_values("fiscal_year").reset_index(drop=True)
    return df


def build_combined_historicals(ticker: str) -> pd.DataFrame:
    income = clean_income_statement(ticker)
    cf = clean_cash_flow(ticker)
    bs = clean_balance_sheet(ticker)
    combined = income.merge(cf, on="fiscal_year", how="outer")
    combined = combined.merge(bs, on="fiscal_year", how="outer")
    combined = combined.sort_values("fiscal_year").reset_index(drop=True)
    out = PROC_DIR / f"{ticker}_historicals.csv"
    combined.to_csv(out, index=False)
    print(f"✓ Saved combined historicals → {out}")
    return combined


if __name__ == "__main__":
    df = build_combined_historicals("MSFT")
    cols = [c for c in ["fiscal_year","revenue","ebit","ebit_margin_pct","free_cash_flow"]
            if c in df.columns]
    print(df[cols].to_string(index=False))