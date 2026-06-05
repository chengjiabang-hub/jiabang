import pandas as pd
from datetime import datetime
import os

URL = "https://pcic.pcc.gov.tw/pwc-web/api/service/opendata-file/document/大宗資材及其漲跌幅彙整表.csv"
FILE = "history.xlsx"

# 抓資料
df_new = pd.read_csv(URL, encoding="big5")
df_new.columns = df_new.columns.str.strip()

# 加時間
today = datetime.today()
df_new["更新年月"] = today.strftime("%Y-%m")

# 找價格欄
price_col = [c for c in df_new.columns if "價格" in c][0]
df_new["價格"] = pd.to_numeric(df_new[price_col], errors="coerce")

# 合併歷史
if os.path.exists(FILE):
    df_old = pd.read_excel(FILE)
    df_all = pd.concat([df_old, df_new], ignore_index=True)

    df_all = df_all.drop_duplicates(
        subset=["調查項目", "調查地區", "更新年月"],
        keep="last"
    )
else:
    df_all = df_new

# 存檔
df_all.to_excel(FILE, index=False)

print("✅ 完成更新")
``
