import os
import re
import pandas as pd
import pdfplumber

def parse_pdf_by_geometric_sequence(file_path, target_date):
    """
    核心演算法：2D幾何序列演算法 (與歷史底稿完全一致，確保格式與數據100%相容)
    """
    rows = []
    configs = [
        {"name": "預拌混凝土 210kgf/cm2", "regions": ["北", "中", "南", "花", "東"]},
        {"name": "預拌混凝土 280kgf/cm2", "regions": ["北", "中", "南", "花", "東"]},
        {"name": "粗級配瀝青混凝土", "regions": ["北", "中", "南", "花", "東"]},
        {"name": "密級配瀝青混凝土", "regions": ["北", "中", "南", "花", "東"]},
        {"name": "鋼筋 SD280", "regions": ["北", "中", "南"]},
        {"name": "鋼筋 SD420", "regions": ["北", "中", "南"]},
        {"name": "結構用鋼材 H型鋼", "regions": ["北", "中", "南"]},
        {"name": "一般結構用軋鋼料 A36", "regions": ["全區"]}
    ]

    with pdfplumber.open(file_path) as pdf:
        words = []
        for page in pdf.pages:
            words.extend(page.extract_words())
        if not words: return []

        note_y = 9999
        for w in words:
            if "備註" in w['text'] or "註：" in w['text']:
                note_y = w['top']
                break

        elements = []
        for w in words:
            if w['top'] > note_y - 10: continue
            t = w['text'].strip()
            is_region = False
            for r in ["北", "中", "南", "花", "東"]:
                if t == r or t == f"{r}部" or (r == "花" and t == "花蓮"):
                    elements.append({'type': 'region', 'val': r, 'x': w['x0'], 'y': w['top']})
                    is_region = True
                    break
            if is_region: continue
            
            nums = re.findall(r'\d{1,2}[,\s]?\d{3}|\d{4,5}', t)
            for n in nums:
                val = int(n.replace(',', '').replace(' ', ''))
                if val not in [210, 280, 420, 400, 110, 111, 112, 113, 114, 115]:
                    elements.append({'type': 'price', 'val': val, 'x': w['x0'], 'y': w['top']})
                    break 

        elements.sort(key=lambda e: (round(e['y'] / 4) * 4, e['x']))

        pairs = []
        i = 0
        while i < len(elements):
            e = elements[i]
            if e['type'] == 'region':
                if i + 1 < len(elements) and elements[i+1]['type'] == 'price':
                    pairs.append({'region': e['val'], 'price': elements[i+1]['val']})
                    i += 2
                else:
                    pairs.append({'region': e['val'], 'price': None})
                    i += 1
            elif e['type'] == 'price':
                if i + 1 < len(elements) and elements[i+1]['type'] == 'region':
                    pairs.append({'region': elements[i+1]['val'], 'price': e['val']})
                    i += 2
                else:
                    pairs.append({'region': '全區', 'price': e['val']})
                    i += 1

        blocks = []
        current_block = {}
        for p in pairs:
            if p['region'] == '北' or p['region'] == '全區':
                if current_block: blocks.append(current_block)
                current_block = {}
            if p['region'] not in current_block:
                current_block[p['region']] = p['price']
        if current_block: blocks.append(current_block)

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
                    "價格": block_data.get(r, None)
                })
    return rows

def append_and_alert_monthly(new_pdf_path, history_excel_path="history.xlsx"):
    """
    讀取單月新 PDF，比對當年度1月價格，進行增量 Append 與 10% 漲幅預警監控
    """
    if not os.path.exists(history_excel_path):
        print(f"❌ 找不到基準歷史檔案 {history_excel_path}，請先確認 build_history.py 已正確執行。")
        return

    # 1. 解析新 PDF 的時間
    file_name = os.path.basename(new_pdf_path)
    raw_date = file_name.replace(".pdf", "").replace(".PDF", "")
    if "." in raw_date:
        year, month = raw_date.split(".")
        target_date = f"{year}.{int(month):02d}"
    else:
        year = raw_date[:3]
        target_date = raw_date
    
    print(f"🔍 偵測到新月份 PDF：{target_date}，開始進行精密解析...")
    new_rows = parse_pdf_by_geometric_sequence(new_pdf_path, target_date)
    df_new = pd.DataFrame(new_rows)
    
    if df_new.empty:
        print("⚠️ 無法從 PDF 提取有效數據，請確認檔案格式是否正確。")
        return

    # 2. 讀取現有的歷史資料庫
    df_history = pd.read_excel(history_excel_path)
    
    # 檢查是否已經重複抓取過該月份
    if target_date in df_history["更新年月"].astype(str).unique():
        print(f"⚠️ {target_date} 的數據已存在於歷史底稿中，為避免重複，取消附加。")
        return

    # 3. 【核心預警機制】找出當年度 1 月的基期價格進行比對
    base_jan_date = f"{year}.01"
    df_jan = df_history[df_history["更新年月"].astype(str) == base_jan_date]
    
    print(f"\n===== 🚨 價格波動風險預警監控 ({target_date} vs {base_jan_date}) =====")
    if df_jan.empty:
        print(f"ℹ️ 歷史資料庫中找不到當年度基期 {base_jan_date} 的資料，跳過預警比對。")
    else:
        # 將當月新資料與 1 月基準資料進行 Merge 比對
        df_alert = pd.merge(df_new, df_jan, on=["調查項目", "調查地區"], suffixes=('_當月', '_一月'))
        alert_triggered = False
        
        for _, row in df_alert.iterrows():
            p_now = row["價格_當月"]
            p_jan = row["價格_一月"]
            
            if p_now and p_jan and p_jan > 0:
                change_rate = (p_now - p_jan) / p_jan
                if change_rate >= 0.10: # 滿足「相較於當年度1月份漲幅大於10%」預警門檻
                    print(f"🔴【紅色預警】{row['調查項目']} ({row['調查地區']}) 漲幅達 {change_rate*100:.1f}%！(1月: {p_jan} -> 當月: {p_now})")
                    alert_triggered = True
        
        if not alert_triggered:
            print("🟩 本月所有主要大宗資材漲幅均溫和穩定，未達 10% 預警門檻。")
    print("==================================================\n")

    # 4. 增量 Append 合併並寫回 Excel
    df_final = pd.concat([df_history, df_new], ignore_index=True)
    # 保持排序：年份 -> 品項 -> 地區
    df_final = df_final.sort_values(by=["更新年月", "調查項目", "調查地區"], ascending=[True, True, True])
    df_final.to_excel(history_excel_path, index=False)
    print(f"💾 增量更新成功！已將 {target_date} 的數據完美附加寫回 {history_excel_path}。")

if __name__ == "__main__":
    # 使用範例：當新月份 115.05.pdf 出現時，直接執行此行
    # 在自動化工作流中，可動態傳入新下載的 PDF 路徑
    new_pdf = "115.05.pdf" 
    if os.path.exists(new_pdf):
        append_and_alert_monthly(new_pdf_path=new_pdf, history_excel_path="history.xlsx")
    else:
        print(f"💡 請將新月份的 PDF 命名為 '{new_pdf}' 並放在同目錄下，或修改程式碼中的 new_pdf 檔名。")