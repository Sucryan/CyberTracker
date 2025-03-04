import os
import sys
import csv
import argparse
import chardet

def merge_csv(input_dir, output_file, domain_col=4):
    """
    合併 input_dir 下所有 CSV 檔案，並以 UTF-8 (無 BOM) 輸出至 output_file。
    流程：
      1. 從第一個 CSV 讀取第一行當作標頭，後續 CSV 均跳過第一行，避免重複。
      2. 若 domain_col >= 0，根據該欄位進行網域去重（只保留第一筆）。
      3. 合併後，不做排序，而是完全重新編號：從第二行開始將第一欄依序設為 1, 2, 3, ...
    注意：請確保每個 CSV 的第一行為標頭，而非資料。
    """
    if not os.path.isdir(input_dir):
        print(f"[警告] 找不到資料夾：{input_dir}")
        sys.exit(1)

    all_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".csv")]
    if not all_files:
        print(f"[警告] {input_dir} 中沒有任何 CSV 檔案。")
        sys.exit(1)

    do_deduplicate = (domain_col >= 0)
    seen_domain = set() if do_deduplicate else None

    merged_rows = []
    header_saved = False

    for file_name in sorted(all_files):
        file_path = os.path.join(input_dir, file_name)
        with open(file_path, "rb") as rb:
            raw_data = rb.read(2048)
        enc = chardet.detect(raw_data)["encoding"] or "utf-8"

        try:
            with open(file_path, "r", encoding=enc, errors="replace") as f:
                reader = csv.reader(f)
                file_header = next(reader, None)
                if not file_header:
                    continue

                if not header_saved:
                    merged_rows.append(file_header)  # 保存標頭
                    header_saved = True

                for row in reader:
                    if do_deduplicate:
                        if len(row) <= domain_col:
                            continue
                        dom = row[domain_col].strip()
                        if dom in seen_domain:
                            continue
                        else:
                            seen_domain.add(dom)
                    merged_rows.append(row)
        except Exception as e:
            print(f"[錯誤] 讀取 {file_name} 失敗：{e}")

    if len(merged_rows) < 2:
        print("[警告] 合併後沒有有效資料。")
        sys.exit(1)

    # 完全重新編號：將資料列（從第二行開始）的第一欄設為 1,2,3,...
    header = merged_rows[0]
    data_rows = merged_rows[1:]
    for i, row in enumerate(data_rows, start=1):
        if len(row) == 0:
            continue
        row[0] = str(i)
    merged_rows = [header] + data_rows

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    try:
        with open(output_file, "w", encoding="utf-8", newline="") as out:
            writer = csv.writer(out)
            for row in merged_rows:
                writer.writerow(row)
        print(f"[INFO] 合併及重新編號完成，輸出：{output_file} (UTF-8, 無 BOM)")
    except Exception as e:
        print(f"[錯誤] 寫入合併結果失敗：{e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(
        description="合併多個 CSV 檔案 (保留第一個 CSV 的標頭)，"
                    "根據指定的網域欄位去重，並完全重新編號（第一欄）。"
    )
    parser.add_argument("--input-dir", required=True, help="CSV 檔案來源資料夾")
    parser.add_argument("--output-file", required=True, help="合併後輸出的 CSV 檔案完整路徑")
    parser.add_argument("--domain-col", type=int, default=4,
                        help="網域欄位的 0-based index (預設=4，即第5欄) 進行去重；若 <0 則不去重")
    args = parser.parse_args()

    merge_csv(
        input_dir=args.input_dir,
        output_file=args.output_file,
        domain_col=args.domain_col
    )

if __name__ == "__main__":
    main()
