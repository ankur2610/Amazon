import pandas as pd
import numpy as np
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

INPUT_FILE = "data/input/Price Comparison - 14Jan'26.xlsx"
OUTPUT_CLEAN = "data/output/clean_master.xlsx"
OUTPUT_LOG = "data/output/daily_price_log.xlsx"

def normalize_name(name):
    return str(name).strip().lower()

def get_amazon_price(page, query):
    # Search Amazon India for the query and return the price (if found)
    search_url = f"https://www.amazon.in/s?k={query.replace(' ', '+')}"
    page.goto(search_url)
    page.wait_for_timeout(4000)  # wait for results to load

    content = page.content()
    soup = BeautifulSoup(content, "lxml")

    price_span = soup.select_one("span.a-price-whole")
    if price_span:
        try:
            return int(price_span.text.replace(",", "").strip())
        except ValueError:
            return np.nan
    return np.nan

def scrape_prices(product_list):
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for product in product_list:
            query = product["normalized_name"]
            price = get_amazon_price(page, query)
            results.append({
                "product_name": product["original_name"],
                "date": datetime.now().strftime("%Y-%m-%d"),
                "price": price
            })
            # Optional delay to reduce request frequency
            time.sleep(2)

        browser.close()
    return results

def main():
    # Read input file
    df = pd.read_excel(INPUT_FILE)
    df["normalized_name"] = df.iloc[:,0].apply(normalize_name)
    df["original_name"] = df.iloc[:,0]

    # Save clean master file
    df.to_excel(OUTPUT_CLEAN, index=False)

    # Scrape prices
    products = df[["original_name", "normalized_name"]].to_dict(orient="records")
    daily_prices = scrape_prices(products)

    # Append or create daily price log
    try:
        log_df = pd.read_excel(OUTPUT_LOG)
        updated_log = pd.concat([log_df, pd.DataFrame(daily_prices)], ignore_index=True)
    except FileNotFoundError:
        updated_log = pd.DataFrame(daily_prices)

    updated_log.to_excel(OUTPUT_LOG, index=False)

if __name__ == "__main__":
    main()
