import pandas as pd
from datetime import datetime
import os

URL = "https://pcic.pcc.gov.tw/pwc-web/api/service/opendata-file/document/大宗資材及其漲跌幅彙整表.csv"
FILE = "history.xlsx"

def fetch_data():
    df = pd.read_csv(URL, encoding="big5")
    df.columns = df.columns.str.strip()
    return df

def transform_data(df):
    today = datetime.today()
    df["更新年月"] = today.strftime("%Y-%m")

    # 找價格欄
    price_col = None
    for col in df.columns:
        if "價格" in col:
            price_col = col
            break

    if price_col is None:
        raise Exception("❌ 找不到價格欄位")

    df["價格"] = pd.to_numeric(df[price_col], errors="coerce")

    # 補必要欄位（避免缺欄位爆掉）
    for col in ["調查項目", "調查地區", "單位"]:
        if col not in df.columns:
            df[col] = ""

    df = df[["更新年月", "調查項目", "調查地區", "單位", "價格"]]

    return df

def update_history(df_new):
    if os.path.exists(FILE):
        try:
            df_old = pd.read_excel(FILE)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except:
            print("⚠️ 舊檔讀取失敗，重新建立")
            df_all = df_new
    else:
        df_all = df_new

    df_all = df_all.drop_duplicates(
