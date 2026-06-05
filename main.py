import pandas as pd
from datetime import datetime
import os

# ✅ 工程會資料來源（官方CSV）
URL = "https://pcic.pcc.gov.tw/pwc-web/api/service/opendata-file/document/大宗資材及其漲跌幅彙整表.csv"

# ✅ 歷史資料檔
FILE = "history.xlsx"


# ✅ 1. 抓資料
def fetch_data():
    df = pd.read_csv(URL, encoding="big5")
    df.columns = df.columns.str.strip()
    return df


# ✅ 2. 整理資料
def transform_data(df):
    today = datetime.today()

    # 加上更新年月（關鍵）
    df["更新年月"] = today.strftime("%Y-%m")

    # 找價格欄（自動）
    price_col = None
    for col in df.columns:
        if "價格" in col:
            price_col = col
            break

    if price_col is None:
        raise Exception("❌ 找不到價格欄位")

    df["價格"] = pd.to_numeric(df[price_col], errors="coerce")

    # 確保欄位存在（避免爆錯）
