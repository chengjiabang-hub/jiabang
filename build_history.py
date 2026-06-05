import pdfplumber
import pandas as pd
import glob

# ✅ 民國轉西元
def convert_minguo(ym):
    y, m = ym.split(".")
    return f"{int(y)+1911}-{int(m):02d}"


# ✅ 解析 PDF
def parse_pdf(file_path):
    ym_text = file_path.split("/")[-1].replace(".pdf","")
    ym = convert_minguo(ym_text)

    data = []

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                current_item = None

                for row in table:
                    if not row:
                        continue

                    row = [str(i) if i else "" for i in row]

                    # 材料名稱
                    if len(row[0]) > 5:
                        current_item = row[0]

                    if len(row) >= 3:
                        region = row[1]
                        price = row[-1]

                        region_map = {
                            "北": "北部",
                            "中": "中部",
                            "南": "南部",
                            "花": "花東",
                            "東": "花東"
                        }

                        try:
                            price = float(price.replace(",", ""))
                        except:
                            continue

                        if region in region_map:
                            data.append({
                                "更新年月": ym,
                                "調查項目": current_item,
                                "調查地區": region_map[region],
                                "單位": "",
                                "價格": price
                            })

    return pd.DataFrame(data)


# ✅ 主程式
def main():
    print("開始建立歷史資料庫")

    files = glob.glob("data/*.pdf")

    if not files:
        print("❌ 沒有找到 PDF，請建立 data 資料夾並放入 PDF")
        return

    all_data = []

    for file in files:
        print(f"處理 {file}")
        df = parse_pdf(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    df_all = df_all.drop_duplicates(
        subset=["調查項目","調查地區","更新年月"]
    )

    df_all.to_excel("history.xlsx", index=False)

    print("✅ 已建立 history.xlsx")


if __name__ == "__main__":
    main()
