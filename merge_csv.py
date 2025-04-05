import os
import sys
import csv
import argparse
import chardet
import tldextract
from urllib.parse import urlparse

def extract_main_domain(url: str) -> str:
    """
    使用 tldextract 精準抽取主網域:
      - 例如：輸入 "https://en.wikipedia.org"，輸出 "wikipedia.org"
    """
    ext = tldextract.extract(url)
    domain = ext.registered_domain
    return domain.lower() if domain else url

def merge_csv(input_dir: str, output_file: str, url_col: int = 2):
    """
    合併 input_dir 下所有 CSV 檔案，並輸出至 output_file。
    流程：
      1. 只保留第一個 CSV 的表頭，後續檔案略過表頭。
      2. 依據指定的網址欄 (預設第 3 欄) 去重（同一網址只保留第一筆）。
      3. 重新編號第一欄（從第二行起依序為 1,2,3,...）。
      4. 若輸出檔名包含 domain.csv，則將網址欄轉換成主網域，並自動補上 https:// 前綴。
      
    注意：其他欄位保持不變，欄位結構與原始格式相同。
    """
    if not os.path.isdir(input_dir):
        print(f"[警告] 找不到資料夾：{input_dir}")
        sys.exit(1)

    all_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".csv")]
    if not all_files:
        print(f"[警告] {input_dir} 中沒有任何 CSV 檔案。")
        sys.exit(1)

    seen_urls = set()
    merged_rows = []
    header_saved = False
    header = None

    # 依檔名排序處理，避免隨機順序
    for file_name in sorted(all_files):
        file_path = os.path.join(input_dir, file_name)
        # 偵測檔案編碼 (讀取前 2KB)
        with open(file_path, "rb") as rb:
            raw_data = rb.read(2048)
        enc = chardet.detect(raw_data)["encoding"] or "utf-8"

        try:
            with open(file_path, "r", encoding=enc, errors="replace") as f:
                reader = csv.reader(f)
                file_header = next(reader, None)
                if not file_header:
                    continue

                # 只在第一個檔保留表頭
                if not header_saved:
                    header = file_header
                    merged_rows.append(file_header)
                    header_saved = True

                for row in reader:
                    if len(row) <= url_col:
                        continue
                    current_url = row[url_col].strip()
                    if current_url in seen_urls:
                        continue
                    seen_urls.add(current_url)
                    merged_rows.append(row)
        except Exception as e:
            print(f"[錯誤] 讀取 {file_name} 失敗：{e}")

    if len(merged_rows) < 2:
        print("[警告] 合併後沒有有效資料。")
        sys.exit(1)

    # 重新編號：假設第一欄 (index=0) 為編號
    data_rows = merged_rows[1:]
    for i, row in enumerate(data_rows, start=1):
        row[0] = str(i)

    final_rows = [merged_rows[0]] + data_rows

    # 如果輸出檔名含有 domain.csv，就修改網址欄 (保留其他欄位不變)
    if "domain.csv" in os.path.basename(output_file).lower():
        for row in final_rows[1:]:
            original_url = row[url_col]
            main_dom = extract_main_domain(original_url)
            if main_dom:
                if not main_dom.startswith("www."):
                    row[url_col] = f"https://www.{main_dom}"
                else:
                    row[url_col] = f"https://{main_dom}"
            else:
                row[url_col] = ""


    # 確保目錄存在
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    try:
        with open(output_file, "w", encoding="utf-8", newline="") as fout:
            writer = csv.writer(fout)
            for row in final_rows:
                writer.writerow(row)
        print(f"[INFO] 成功輸出：{output_file}")
    except Exception as e:
        print(f"[錯誤] 寫入 {output_file} 失敗：{e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(
        description="合併多個 CSV（依據指定的網址欄去重並重新編號），輸出 total.csv 或 domain.csv。"
    )
    parser.add_argument("--input-dir", required=True, help="CSV 檔案來源資料夾")
    parser.add_argument("--output-file", required=True, help="合併後輸出的 CSV 檔案完整路徑")
    parser.add_argument("--url-col", type=int, default=2, help="網址欄的 0-based index，預設為 2 (第 3 欄)")
    args = parser.parse_args()

    merge_csv(
        input_dir=args.input_dir,
        output_file=args.output_file,
        url_col=args.url_col
    )

if __name__ == "__main__":
    main()
