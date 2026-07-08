import math
from datetime import datetime

import pandas as pd
import yfinance as yf
import oracledb


# ---------------- Oracle Configuration ---------------- #
USERNAME = "system"
PASSWORD = "admin"
DSN = "localhost:1521/XEPDB1"

EXCEL_FILE = r"C:\Users\xyz\OneDrive\_FUND_VALUE.xlsx"


# ---------------- Read Portfolio ---------------- #
def read_portfolio(file_name):
    try:
        stocks=[]
]
        df = pd.read_excel(file_name, engine="openpyxl")
        print(f"{len(df)} stocks loaded from Excel.")
        for idx in range(min(len(df), len(stocks))):
            if df.loc[idx, "Company Name"] != stocks[idx]:
                df.loc[idx, "Company Name"] = stocks[idx]
        return df
    except Exception as e:
        print("Unable to read Excel:", e)
        raise


# ---------------- Fetch Live Prices ---------------- #
def fetch_stock_prices(df):

    prices = []
    previous_close = []
    change_percent = []

    for stock in df["Company Name"]:

        symbol = stock if stock in ["EPAM", "GOOG"] else stock + ".NS"

        try:

            ticker = yf.Ticker(symbol)

            current_price = ticker.fast_info.get("lastPrice")

            hist = ticker.history(period="5d").tail(2)

            if len(hist) >= 2:
                prev_close = hist["Close"].iloc[-2]
            else:
                prev_close = None

            if current_price is not None and prev_close is not None:
                change = round(((current_price - prev_close) / prev_close) * 100, 2)
            else:
                change = None

        except Exception as e:

            print(f"Unable to fetch {stock}: {e}")

            current_price = None
            prev_close = None
            change = None

        prices.append(current_price)
        previous_close.append(prev_close)
        change_percent.append(change)

    df["Current Price"] = prices
    df["Previous Close"] = previous_close
    df["Today's Change %"] = change_percent

    return df


# ---------------- Calculate Overall ---------------- #
def calculate_metrics(df):

    df["overall %"] = (
        df["Today's Change %"] *
        (df["weightage"] / 100)
    )

    df["Current Price"] = df["Current Price"].round(2)
    df["Previous Close"] = df["Previous Close"].round(2)
    df["Today's Change %"] = df["Today's Change %"].round(2)
    df["overall %"] = df["overall %"].round(4)

    return df


# ---------------- Load into Oracle ---------------- #
def load_to_oracle(df):

    connection = None
    cursor = None

    try:

        connection = oracledb.connect(
            user=USERNAME,
            password=PASSWORD,
            dsn=DSN
        )

        cursor = connection.cursor()

        # Replace NaN with None
        df = df.where(pd.notna(df), None)

        data = list(
            df[
                [
                    "Company Name",
                    "Current Price",
                    "Previous Close",
                    "Today's Change %",
                    "weightage",
                    "overall %"
                ]
            ].itertuples(index=False, name=None)
        )

        cursor.executemany("""
        INSERT INTO STOCKANALYSISIS
        (
            COMPANY_NAME,
            CURRENT_PRICE,
            PREVIOUS_CLOSE,
            TODAY_CHANGE,
            WEIGHTAGE,
            OVERALL,
            INSERTED_DATE,
            INSERTED_TIME
            
        )
        VALUES
        (
            :1,:2,:3,:4,:5,:6, SYSDATE,SYSTIMESTAMP
        )
        """, data)

        connection.commit()

        print(f"{cursor.rowcount} rows inserted successfully.")

    except Exception as e:

        if connection:
            connection.rollback()

        print("Oracle Error:", e)

    finally:

        if cursor:
            cursor.close()

        if connection:
            connection.close()


# ---------------- Export Excel ---------------- #
def export_excel(df):

    cols = ["weightage", "overall %"]

    total_row = df[cols].sum()

    total_row["Company Name"] = "Total"

    df.loc[len(df)] = total_row

    filename = f"Portfolio_{datetime.today().strftime('%Y%m%d')}.xlsx"

    df.to_excel(filename, index=False)

    print(f"Excel exported: {filename}")


# ---------------- Main ---------------- #
def main():

    df = read_portfolio(EXCEL_FILE)

    df = fetch_stock_prices(df)

    df = calculate_metrics(df)

    load_to_oracle(df)

    export_excel(df)

    print(df)


if __name__ == "__main__":
    main()
