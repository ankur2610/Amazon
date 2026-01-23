import pandas as pd

INPUT_LOG = "data/output/daily_price_log.xlsx"
OUTPUT_ANALYTICS = "data/output/analytics_summary.xlsx"

def compute_analytics():
    df = pd.read_excel(INPUT_LOG)
    if df.empty:
        return

    # Ensure date column is datetime
    df["date"] = pd.to_datetime(df["date"])

    # Basic daily metrics
    df.sort_values(["product_name", "date"], inplace=True)
    df["price_change_abs"] = df.groupby("product_name")["price"].diff()
    df["price_change_pct"] = df.groupby("product_name")["price"].pct_change() * 100

    # 7-day rolling metrics
    df["rolling_avg_7d"] = df.groupby("product_name")["price"].rolling(7, min_periods=1).mean().reset_index(0, drop=True)
    df["rolling_min_7d"] = df.groupby("product_name")["price"].rolling(7, min_periods=1).min().reset_index(0, drop=True)
    df["rolling_max_7d"] = df.groupby("product_name")["price"].rolling(7, min_periods=1).max().reset_index(0, drop=True)

    # 30-day min and max
    df["rolling_min_30d"] = df.groupby("product_name")["price"].rolling(30, min_periods=1).min().reset_index(0, drop=True)
    df["rolling_max_30d"] = df.groupby("product_name")["price"].rolling(30, min_periods=1).max().reset_index(0, drop=True)

    # Buy signal: within 5% of 30-day low
    df["buy_signal"] = df["price"] <= df["rolling_min_30d"] * 1.05

    # Write analytics
    df.to_excel(OUTPUT_ANALYTICS, index=False)

if __name__ == "__main__":
    compute_analytics()
