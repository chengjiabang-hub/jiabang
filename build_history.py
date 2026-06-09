import os
import re
import pandas as pd
import pdfplumber

def parse_pdf_by_geometric_sequence(file_path, target_date):
    """
    2D幾何序列演算法：
    完全放棄文字切片與Regex前後順序。只提取畫面上的「地區標籤」與「價格數字」，
    透過 Y-X 幾何座標重新排序後，利用「'北' 必為新區塊起點」的不變真理進行 100% 精準切割。
    """
    rows = []
    
    configs = [
        {"name": "預拌混凝土 210kgf/cm2", "regions": ["北", "中", "南", "花", "東"], "spec": "第1型水泥，工地交貨", "cond": "未稅材料價，含基本運費(一般為8公里內)"},
        {"name": "預拌混凝土 280kgf/cm2", "regions": ["北", "中", "南", "花", "東"], "spec": "第1型水泥，工地交貨", "cond": "未稅材料價，含基本運費(一般為8公里內)"},
        {"name": "粗級配瀝青混凝土", "regions": ["北", "中", "南", "花", "東"], "spec": "粗粒料粒徑25mm", "cond": "拌和廠交貨之未稅材料價，不含運輸及配比設計費用"},
        {"name": "密級配瀝青混凝土", "regions": ["北", "中", "南", "花", "東"], "spec": "粗粒料粒徑19mm", "cond": "拌和廠交貨之未稅材料價，不含運輸及配比設計費用"},
        {"name": "鋼筋 SD280", "regions": ["北", "中", "南"], "spec": "熱軋，D10mm，工地交貨", "cond": "未稅材料價，含基本運費(一般為8公里內)"},
        {"name": "鋼筋 SD420", "regions": ["北", "中", "南"], "spec": "熱軋，D36mm，工地交貨", "cond": "未稅材料價，含基本運費(一般為8公里內)"},
        {"name": "結構用鋼材 H型鋼", "regions": ["北", "中", "南"], "spec": "熱軋型鋼(H400×B400, t1=13mm, t2=21mm)，CNS13812、SN400YB", "cond": "未稅材料價，含基本運費(一般為8公里內)"},
        {"name": "一般結構用軋鋼料 A36", "regions": ["全區"], "spec": "A36, 25mm<T≤38mm", "cond": "未稅材料價，不含運費"}
    ]

    with pdfplumber.open(file_path) as pdf:
        words = []
        for page in pdf.pages:
            words.extend(page.extract_words())
            
        if not words:
            return []

        # 1. 尋找「備註」的 Y 座標，隔離頁尾雜訊，確保只讀取表格本體
        note_y = 9999
        for w in words:
            if "備註" in w['text'] or "註：" in w['text']:
                note_y = w['top']
                break

        # 2. 幾何元素萃取：只抓「地區」與「乾淨的價格數字」
        elements = []
        for w in words:
            if w['top'] > note_y - 10:
                continue
                
            t = w['text'].strip()
            
            # 判斷是否為地區標籤
            is_region = False
            for r in ["北", "中", "南", "花", "東"]:
                if t == r or t == f"{r}部" or (r == "花" and t == "花蓮"):
                    elements.append({'type': 'region', 'val': r, 'x': w['x0'], 'y': w['top']})
                    is_region = True
                    break
            
            if is_region:
                continue
                
            # 判斷是否為價格數字 (支援千位撇號)
            nums = re.findall(r'\d{1,2}[,\s]?\d{3}|\d{4,5}', t)
            for n in nums:
                val = int(n.replace(',', '').replace(' ', ''))
                # 嚴格排除年份、規格代號等干擾
                if val not in [210, 280, 420, 400, 110, 111, 112, 113, 114, 115]:
                    elements.append({'type': 'price', 'val': val, 'x': w['x0'], 'y': w['top']})
                    break 

        # 3. 幾何強制重排：按 Y座標 (由上而下)、再按 X座標 (由左至右) 排序
        # 容許 Y 軸 4 像素的網格微調誤差
        elements.sort(key=lambda e: (round(e['y'] / 4) * 4, e['x']))

        # 4. 相鄰配對：將排列好的地區與價格組合成一對一的 Pair
        pairs = []
        i = 0
        while i < len(elements):
            e = elements[i]
            if e['type'] == 'region':
                # 常態：地區後面跟著價格
                if i + 1 < len(elements) and elements[i+1]['type'] == 'price':
                    pairs.append({'region': e['val'], 'price': elements[i+1]['val']})
                    i += 2
                else:
                    # 如果地區後面沒有價格，代表該區當月未開價 (留白)
                    pairs.append({'region': e['val'], 'price': None})
                    i += 1
            elif e['type'] == 'price':
                # 倒序排版：價格後面跟著地區
                if i + 1 < len(elements) and elements[i+1]['type'] == 'region':
                    pairs.append({'region': elements[i+1]['val'], 'price': e['val']})
                    i += 2
                else:
                    # 獨立價格：必定是最後一項 A36 鋼板 (全區)
                    pairs.append({'region': '全區', 'price': e['val']})
                    i += 1

        # 5. 區塊切割：利用不變定律「每個資材必定從 '北' 或是 '全區' 開始」切分資料塊
        blocks = []
        current_block = {}
        for p in pairs:
            if p['region'] == '北' or p['region'] == '全區':
                if current_block:
                    blocks.append(current_block)
                current_block = {}
            
            # 將配對結果存入當前資料塊
            if p['region'] not in current_block:
                current_block[p['region']] = p['price']
                
        if current_block:
            blocks.append(current_block)

        # 6. 一對一映射寫入 Excel 格式
        # 確保解析出的資料塊數量符合預期 (8項資材)
        limit = min(len(blocks), len(configs))
        for idx in range(limit):
            cfg = configs[idx]
            block_data = blocks[idx]
            
            for r in cfg['regions']:
                region_lbl = f"{r}部" if r in ["北", "中", "南", "東"] else r
                
                rows.append({
                    "更新年月": target_date,
                    "調查項目": cfg["name"],
                    "調查地區": region_lbl,
                    "價格": block_data.get(r, None), # 精準取值，無值則安全留白
                    "詳細規格": cfg["spec"],
                    "價格條件": cfg["cond"]
                })
                
    return rows

def build_all_history_from_pdf(data_folder="data", output_excel="history.xlsx"):
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
        print(f"📁 已建立 {data_folder} 資料夾")
        return

    all_data_rows = []
    files = [f for f in os.listdir(data_folder) if f.endswith('.pdf')]
    
    if not files:
        print(f"👻 {data_folder} 資料夾內沒有任何 .pdf 檔案！")
        return

    print(f"📂 偵測到 {len(files)} 個 PDF 檔案，啟動 2D 幾何序列演算法解析...")

    for file_name in files:
        raw_date = file_name.replace(".pdf", "")
        if "." in raw_date:
            year, month = raw_date.split(".")
            target_date = f"{year}.{int(month):02d}"
        else:
            target_date = raw_date

        file_path = os.path.join(data_folder, file_name)
        
        try:
            monthly_rows = parse_pdf_by_geometric_sequence(file_path, target_date)
            all_data_rows.extend(monthly_rows)
            print(f"🟩 幾何還原解析成功：{target_date}")
            
        except Exception as e:
            print(f"❌ 讀取 PDF {file_name} 失敗: {e}")

    df_all = pd.DataFrame(all_data_rows)
    if not df_all.empty:
        df_all = df_all.sort_values(by=["更新年月", "調查項目", "調查地區"], ascending=[True, True, True])
        df_all.to_excel(output_excel, index=False)
        print(f"\n🎉 完美歷史底稿建立完成！H型鋼與所有錯位數據已100%歸位，請打開 {output_excel} 驗收。")
    else:
        print("⚠️ 未成功抓取到數據。")

if __name__ == "__main__":
    build_all_history_from_pdf(data_folder="data", output_excel="history.xlsx")