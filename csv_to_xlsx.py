import sys
import csv
import openpyxl
import chardet

def detect_encoding(file_path, sample_size=1000):
    """檢測 CSV 檔案的編碼"""
    with open(file_path, "rb") as f:
        raw_data = f.read(sample_size)
    return chardet.detect(raw_data)["encoding"]

def convert_csv_to_xlsx(input_csv, output_xlsx):
    # 嘗試偵測 CSV 檔案的編碼
    detected_encoding = detect_encoding(input_csv)
    print(f"[INFO] 偵測到的 CSV 編碼: {detected_encoding}")

    # 讀取 CSV 時強制轉為 UTF-8
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    with open(input_csv, "r", encoding=detected_encoding, errors="replace") as f:
        reader = csv.reader(f)
        for row in reader:
            ws.append(row)

    # 儲存 XLSX
    wb.save(output_xlsx)
    print(f"[INFO] 已輸出 XLSX：{output_xlsx}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法：python csv_to_xlsx.py <input.csv> <output.xlsx>")
        sys.exit(1)

    input_csv = sys.argv[1]
    output_xlsx = sys.argv[2]
    convert_csv_to_xlsx(input_csv, output_xlsx)
