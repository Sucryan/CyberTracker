import os
import sys
import csv
import argparse
import chardet

def merge_csv(input_dir, output_file, domain_col=4):
    """
    合併 input_dir 下所有 .csv, 以 UTF-8 (無 BOM) 寫出 output_file。
    預設會以 domain_col 做網域去重：
      - domain_col >= 0 表示要根據該欄位去重
      - domain_col <  0 表示不做去重

    :param input_dir:   要合併的 CSV 資料夾路徑
    :param output_file: 合併後的檔案路徑 (UTF-8 無 BOM)
    :param domain_col:  網域欄位的 0-based index，預設=4 (第 5 欄) 去重。若為 -1 則不去重。
    """

    if not os.path.isdir(input_dir):
        print(f"[警告] 找不到資料夾：{input_dir}")
        return

    all_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".csv")]
    if not all_files:
        print(f"[警告] {input_dir} 中沒有任何 .csv 檔案，跳過合併。")
        return

    # 若需要去重，這裡存已出現的 domain
    seen_domain = set() if domain_col >= 0 else None  

    merged_data = []
    header_saved = False

    for file_name in all_files:
        file_path = os.path.join(input_dir, file_name)

        # 偵測原檔編碼
        with open(file_path, "rb") as rb:
            raw_data = rb.read(2048)
        enc = chardet.detect(raw_data)["encoding"] or "utf-8"

        try:
            with open(file_path, "r", encoding=enc, errors="replace") as f:
                reader = csv.reader(f)
                file_header = next(reader, None)  # 讀第一行當作 header
                if not file_header:
                    continue

                # 第一個檔案保留 header
                if not header_saved:
                    merged_data.append(file_header)
                    header_saved = True

                # 後續資料
                for row in reader:
                    if domain_col >= 0:
                        # 若要去重 domain
                        if len(row) <= domain_col:
                            continue
                        dom = row[domain_col].strip()
                        if dom in seen_domain:
                            continue
                        else:
                            seen_domain.add(dom)
                    # 加入合併結果
                    merged_data.append(row)

        except Exception as e:
            print(f"[錯誤] 讀取 {file_name} 失敗：{e}")

    if len(merged_data) < 2:
        # 可能只有 header 或是空檔
        print("[警告] 合併後沒有任何資料。")
        return

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    try:
        with open(output_file, "w", encoding="utf-8", newline="") as out:
            writer = csv.writer(out)
            for row in merged_data:
                writer.writerow(row)
        print(f"[INFO] 合併完成：{output_file} (UTF-8, 無 BOM)")
    except Exception as e:
        print(f"[錯誤] 寫入合併結果失敗：{e}")


def main():
    parser = argparse.ArgumentParser(
        description="合併多個 CSV 檔，預設根據第5欄 (domain_col=4) 去重。"
    )
    parser.add_argument("--input-dir", required=True, help="要合併的 CSV 資料夾路徑")
    parser.add_argument("--output-file", required=True, help="合併後輸出的檔案路徑")
    parser.add_argument("--domain-col", type=int, default=4,
                        help="網域欄位的 0-based index，預設=4 (第5欄) 進行去重；若指定 -1 則不去重")
    args = parser.parse_args()

    merge_csv(
        input_dir=args.input_dir,
        output_file=args.output_file,
        domain_col=args.domain_col
    )

if __name__ == "__main__":
    main()
