import sys
import csv
import openpyxl
from openpyxl.styles import numbers
import datetime

# 根據實際 CSV 值的格式，寫對應的 strptime 格式字串
# 假設「詐騙網站接獲日期」原本 CSV 內容是 "2023-03-01 12:34:56"
# 假設「停止解析日期」原本 CSV 內容是 "20230302"
# 如果你的原 CSV 有其他格式，請自行修改。
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"  # ex: 2023-03-01 12:34:56
DATE_FORMAT = "%Y%m%d"                # ex: 20230302

def convert_csv_to_xlsx(input_csv, output_xlsx):
    """
    將 CSV 轉為 XLSX，並針對
      1) 「詐騙網站接獲日期」 → yyyy-mm-ddTHH:mm:ss
      2) 「停止解析日期」     → yyyymmdd
    做特定格式設定
    """
    # 讀取 CSV 到記憶體
    with open(input_csv, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        print("[警告] CSV 為空，無法轉檔。")
        return

    # 建立 Workbook 與 Worksheet
    wb = openpyxl.Workbook()
    ws = wb.active

    # 將資料寫入 Worksheet
    for row_data in rows:
        ws.append(row_data)

    # 偵測 header
    header = rows[0]
    # 找出「詐騙網站接獲日期」「停止解析日期」欄位的索引 (0-based)
    idx_received = None
    idx_stopped = None
    try:
        idx_received = header.index("詐騙網站接獲日期")
    except ValueError:
        pass
    try:
        idx_stopped = header.index("停止解析日期")
    except ValueError:
        pass

    # 從第 2 列開始處理（第 1 列是表頭）
    for r in range(2, ws.max_row + 1):
        # 如果有「詐騙網站接獲日期」欄位
        if idx_received is not None:
            cell_received = ws.cell(row=r, column=idx_received + 1)  # openpyxl 是 1-based
            val = cell_received.value
            # 嘗試轉成 datetime
            if val:
                try:
                    dt = datetime.datetime.strptime(val, DATETIME_FORMAT)
                    cell_received.value = dt
                    # 設定儲存格格式 → yyyy-mm-ddTHH:mm:ss
                    # 其中 "T" 必須要用跳脫或雙引號，Excel 才能顯示
                    # 可嘗試 "yyyy-mm-dd\\THH:mm:ss" 讓 T 顯示為文字
                    cell_received.number_format = r'yyyy-mm-dd"T"HH:mm:ss'
                except ValueError:
                    pass  # 解析失敗就保持原文字

        # 如果有「停止解析日期」欄位
        if idx_stopped is not None:
            cell_stopped = ws.cell(row=r, column=idx_stopped + 1)
            val = cell_stopped.value
            if val:
                try:
                    dt2 = datetime.datetime.strptime(val, DATE_FORMAT)
                    # 只存日期
                    cell_stopped.value = dt2.date()
                    # 設定格式為 yyyymmdd
                    cell_stopped.number_format = "yyyymmdd"
                except ValueError:
                    pass

    wb.save(output_xlsx)
    print(f"[INFO] 已輸出 XLSX：{output_xlsx}")

def main():
    if len(sys.argv) < 3:
        print("用法：python csv_to_xlsx.py <input.csv> <output.xlsx>")
        sys.exit(1)
    input_csv = sys.argv[1]
    output_xlsx = sys.argv[2]
    convert_csv_to_xlsx(input_csv, output_xlsx)

if __name__ == "__main__":
    main()
