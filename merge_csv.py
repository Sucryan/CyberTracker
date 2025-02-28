import os
import sys
import glob
import pandas as pd
import chardet

def detect_encoding(file_path, num_bytes=1024):
    """偵測檔案編碼，預設讀取前 1024 bytes"""
    with open(file_path, 'rb') as f:
        rawdata = f.read(num_bytes)
    result = chardet.detect(rawdata)
    return result['encoding']

def merge_csv_files():
    # 取得目前執行檔所在的資料夾（適用於打包成 exe 時）
    cur_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    input_folder = os.path.join(cur_dir, "all_csv")
    output_folder = os.path.join(cur_dir, "csv_stuff")

    # 若輸出資料夾不存在則建立
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"建立資料夾: {output_folder}")

    # 搜尋所有 CSV 檔案
    csv_files = glob.glob(os.path.join(input_folder, "*.csv"))
    if not csv_files:
        print("找不到任何 CSV 檔案！")
        return None

    df_list = []
    for file in csv_files:
        encoding = detect_encoding(file)
        print(f"{file} 偵測到編碼: {encoding}")
        try:
            df = pd.read_csv(file, encoding=encoding)
        except Exception as e:
            print(f"讀取 {file} 時發生錯誤: {e}")
            continue
        df_list.append(df)
        print(f"成功讀取: {file}")

    if not df_list:
        print("沒有成功讀取任何 CSV 檔案")
        return None

    # 合併所有 DataFrame
    merged_df = pd.concat(df_list, ignore_index=True)

    # 調整「編號」欄位，依序重新編號（若該欄存在）
    if "編號" in merged_df.columns:
        merged_df["編號"] = range(1, len(merged_df) + 1)
        print("已重新調整「編號」欄位的序號。")

    # 輸出結果到 CSV，用 utf-8-sig 以利 Excel 正確開啟中文
    output_path = os.path.join(output_folder, "total.csv")
    merged_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"CSV 合併完成，結果儲存於: {output_path}")
    return output_path

if __name__ == "__main__":
    merge_csv_files()
