import os
import sys
import csv
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import time
import winshell  # 需先 pip install winshell

# 取得目前執行檔所在的資料夾
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

# 取得正確的桌面路徑 (例如 C:\Users\a0987\OneDrive\桌面)
DESKTOP_DIR = winshell.desktop()

# 定義輸出主資料夾，以及各子資料夾：HTML、PNG、XLSX（若有需要報告檔另存）
OUTPUT_DIR = os.path.join(DESKTOP_DIR, "CyberTrackerOutput")
HTML_DIR = os.path.join(OUTPUT_DIR, "html")
PNG_DIR = os.path.join(OUTPUT_DIR, "png")
XLSX_DIR = os.path.join(OUTPUT_DIR, "xlsx")  # 若需要 CSV→XLSX 轉檔

# -------------------------------------
# 輔助函式：寫 log
def log(msg, text_widget):
    text_widget.insert(tk.END, msg + "\n")
    text_widget.see(tk.END)

def thread_safe_log(msg, text_widget, root):
    root.after(0, lambda: log(msg, text_widget))

# -------------------------------------
# 輔助函式：從 CSV 第二行提取品牌名稱
def extract_brand_from_row(row):
    # 嘗試檢查 row 中每個欄位，找出包含「假冒」或「偽冒」的字串，取後面兩個字
    for cell in row:
        if "假冒" in cell:
            pos = cell.find("假冒")
            if len(cell) >= pos + 4:
                return cell[pos+2:pos+4]
        if "偽冒" in cell:
            pos = cell.find("偽冒")
            if len(cell) >= pos + 4:
                return cell[pos+2:pos+4]
    return "電商"  # 預設值

# -------------------------------------
# 複製 total.csv 並重新命名為 YYYYMMDD.申報.[品牌](N筆).csv
def extract_brand_from_row(row):
    """
    從第二行 (row) 的各欄位連接後的字串，判斷是否符合下列規則：
      1) 若開頭是「疑似假冒」 → 取 text[4:6] 當品牌
      2) 若開頭兩字是「偽冒」或「假冒」 → 取 text[2:4] 當品牌
      3) 否則 → '電商'
    """
    text = ''.join(row).strip()  # 把該行所有欄位串起來做檢查

    # 規則 1：若字串開頭是「疑似假冒」
    if text.startswith("疑似假冒"):
        if len(text) >= 6:
            return text[4:6]  # 取疑似假冒後面兩字
        else:
            return "電商"

    # 規則 2：若前兩字是「偽冒」或「假冒」
    front2 = text[:2]
    if front2 in ["偽冒", "假冒"]:
        if len(text) >= 4:
            return text[2:4]  # 取後兩字
        else:
            return "電商"

    # 都不符合則回傳「電商」
    return "電商"


def copy_total_csv_report(total_csv_path, output_folder):
    """
    讀取 total.csv，計算總筆數 (第一列為表頭，第二列起為資料)。
    1) 從第二行 (第一筆資料) 解析品牌名稱。
    2) 命名為 YYYYMMDD.申報.[品牌](N筆).csv
    3) 以 UTF-8 寫出。
    """
    total_count = 0
    brand = "電商"  # 預設
    second_row = None

    try:
        with open(total_csv_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)  # 第一行為表頭

            # 嘗試讀取第二行 (若有資料)
            try:
                second_row = next(reader)
            except StopIteration:
                second_row = None

            # 計算總筆數：若 second_row 存在就先 +1
            if second_row:
                total_count += 1

            # 後續每列
            for row in reader:
                if row and row[0].strip():
                    total_count += 1

    except Exception as e:
        raise Exception(f"讀取 total.csv 失敗：{e}")

    # 解析品牌
    if second_row:
        brand = extract_brand_from_row(second_row)

    # 產生檔名
    today_str = time.strftime("%Y%m%d", time.localtime())
    new_filename = f"{today_str}.申報.{brand}({total_count}筆).csv"
    new_file_path = os.path.join(output_folder, new_filename)

    # 複製 total.csv → 產生最終報告檔 (UTF-8)
    try:
        with open(total_csv_path, "r", encoding="utf-8") as src, \
             open(new_file_path, "w", encoding="utf-8", newline="") as dst:
            shutil.copyfileobj(src, dst)
    except Exception as e:
        raise Exception(f"寫入報告檔失敗：{e}")

    return new_file_path

# -------------------------------------
# 清空 all_csv 資料夾
def clear_all_csv(text_widget, root):
    target_dir = os.path.join(BASE_DIR, "all_csv")
    if os.path.exists(target_dir):
        try:
            for filename in os.listdir(target_dir):
                file_path = os.path.join(target_dir, filename)
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            log("all_csv 資料夾已清空。", text_widget)
        except Exception as e:
            log(f"清空 all_csv 失敗：{e}", text_widget)
    else:
        os.makedirs(target_dir, exist_ok=True)
        log("all_csv 資料夾不存在，已建立。", text_widget)

# -------------------------------------
# 匯入資料：將 CSV 複製到 all_csv 資料夾
def import_data(text_widget, root):
    csv_file = filedialog.askopenfilename(
        title="請選擇 CSV 檔案",
        filetypes=[("CSV files", "*.csv")]
    )
    if not csv_file:
        return
    target_dir = os.path.join(BASE_DIR, "all_csv")
    os.makedirs(target_dir, exist_ok=True)
    try:
        shutil.copy(csv_file, target_dir)
        log(f"成功匯入：{os.path.basename(csv_file)} 至 {target_dir}", text_widget)
    except Exception as e:
        log(f"匯入失敗：{e}", text_widget)

# -------------------------------------
# CSV → XLSX 轉檔 (單獨使用)
def convert_csv_button(text_widget, root):
    csv_file = filedialog.askopenfilename(
        title="請選擇要轉為 XLSX 的 CSV 檔案",
        filetypes=[("CSV files", "*.csv")]
    )
    if not csv_file:
        return
    base_name = os.path.splitext(os.path.basename(csv_file))[0]
    os.makedirs(XLSX_DIR, exist_ok=True)
    output_xlsx = os.path.join(XLSX_DIR, base_name + ".xlsx")
    csv_to_xlsx_exe = os.path.join(BASE_DIR, "csv_to_xlsx.exe")
    if not os.path.exists(csv_to_xlsx_exe):
        messagebox.showerror("錯誤", "找不到 csv_to_xlsx.exe 執行檔！")
        return

    def run_conversion():
        try:
            thread_safe_log(
                f"開始轉檔：{os.path.basename(csv_file)} → {os.path.basename(output_xlsx)}",
                text_widget, root
            )
            subprocess.run([csv_to_xlsx_exe, csv_file, output_xlsx], check=True)
            thread_safe_log(f"轉檔完成：{output_xlsx}", text_widget, root)
        except Exception as e:
            thread_safe_log(f"轉檔失敗：{e}", text_widget, root)
    t = threading.Thread(target=run_conversion)
    t.start()

# -------------------------------------
# 產生通報 (TXT 報告)
def generate_domain_report_txt(csv_file, output_txt):
    domains = set()
    try:
        with open(csv_file, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header:
                return "[錯誤] CSV 檔案為空！"
            try:
                domain_index = header.index("網域")
            except ValueError:
                return "[錯誤] 找不到「網域」欄位，請確認 CSV 格式！"
            for row in reader:
                if len(row) > domain_index:
                    domain = row[domain_index].strip()
                    if domain:
                        domains.add(domain)
    except Exception as e:
        return f"[錯誤] 讀取 CSV 失敗：{e}"
    formatted_text = "Hi @蘇昱丞 Ethan Su\n\n今天預計通報\n\n"
    formatted_text += "\n".join(sorted(domains))
    formatted_text += "\n\n截圖如下:\n\n"
    formatted_text += "此可以判斷為詐騙的原因:\n\n"
    try:
        with open(output_txt, "w", encoding="utf-8") as f:
            f.write(formatted_text)
    except Exception as e:
        return f"[錯誤] 寫入報告失敗：{e}"
    return f"已生成通報文件: {output_txt}"

def generate_report(text_widget, root):
    merged_csv = os.path.join(BASE_DIR, "csv_stuff", "total.csv")
    output_txt = os.path.join(OUTPUT_DIR, "report.txt")
    result = generate_domain_report_txt(merged_csv, output_txt)
    log(result, text_widget)

# -------------------------------------
# 嵌入圖片至 HTML：
# 處理 HTML 資料夾 (HTML_DIR) 中的 HTML 檔，
# 從 PNG 資料夾 (PNG_DIR) 中找到對應圖片檔後複製到 HTML_DIR，
# 並在 HTML 檔中插入 <img src="[圖片檔名]"> 標籤
def embed_image_in_html(html_folder, png_folder, root, text_widget):
    try:
        html_files = [f for f in os.listdir(html_folder) if f.lower().endswith(".html")]
        if not html_files:
            thread_safe_log("找不到 HTML 檔案以嵌入圖片。", text_widget, root)
            return

        for html_file in html_files:
            base = os.path.splitext(html_file)[0]
            png_file = base + ".png"
            png_src = os.path.join(png_folder, png_file)
            if os.path.exists(png_src):
                png_dest = os.path.join(html_folder, png_file)
                shutil.copy2(png_src, png_dest)
                html_path = os.path.join(html_folder, html_file)
                with open(html_path, "r", encoding="utf-8") as f:
                    content = f.read()
                img_tag = f'\n<img src="{png_file}" alt="Screenshot">\n'
                insert_pos = content.lower().rfind("</body>")
                if insert_pos == -1:
                    new_content = content + img_tag
                else:
                    new_content = content[:insert_pos] + img_tag + content[insert_pos:]
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(new_content)
                thread_safe_log(f"在 {html_file} 中嵌入圖片 {png_file}", text_widget, root)
            else:
                thread_safe_log(f"找不到對應於 {html_file} 的 PNG 檔案。", text_widget, root)
    except Exception as e:
        thread_safe_log(f"嵌入圖片到 HTML 時發生錯誤：{e}", text_widget, root)

# -------------------------------------
# 一鍵完成流程：
# 1. 執行 merge_csv.exe 產生 total.csv (在 BASE_DIR\csv_stuff)
# 2. 呼叫 web_capture.exe 分別輸出：
#    - 截圖 (.png) 至 PNG_DIR
#    - HTML (.html) 至 HTML_DIR
# 3. 生成報告 TXT (通報檔)，並複製 total.csv 重新命名為 YYYYMMDD.申報.[品牌](N筆).csv
def run_long_task(text_widget, root, progressbar):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(HTML_DIR, exist_ok=True)
    os.makedirs(PNG_DIR, exist_ok=True)
    # 確保 csv_stuff 資料夾存在
    csv_stuff_dir = os.path.join(BASE_DIR, "csv_stuff")
    os.makedirs(csv_stuff_dir, exist_ok=True)

    merge_exe = os.path.join(BASE_DIR, "merge_csv.exe")
    web_capture_exe = os.path.join(BASE_DIR, "web_capture.exe")

    if not os.path.exists(merge_exe):
        thread_safe_log("錯誤：找不到 merge_csv.exe！", text_widget, root)
        root.after(0, progressbar.stop)
        return

    # 呼叫 merge_csv.exe 時傳入參數
    input_dir = os.path.join(BASE_DIR, "all_csv")
    output_csv = os.path.join(csv_stuff_dir, "total.csv")
    thread_safe_log("開始執行 merge_csv...", text_widget, root)
    try:
        subprocess.run([
            merge_exe,
            "--input-dir", input_dir,
            "--output-file", output_csv
        ], check=True)
    except Exception as e:
        thread_safe_log(f"merge_csv 執行失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    if not os.path.isfile(output_csv):
        thread_safe_log(f"錯誤：找不到合併後的 total.csv：{output_csv}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    thread_safe_log(f"合併後 CSV 檔案：{output_csv}", text_widget, root)

    if not os.path.exists(web_capture_exe):
        thread_safe_log("錯誤：找不到 web_capture.exe！", text_widget, root)
        root.after(0, progressbar.stop)
        return

    # 呼叫 web_capture.exe 輸出截圖至 PNG_DIR
    thread_safe_log("啟動 web_capture [screenshot] 模式...", text_widget, root)
    proc_screenshot = subprocess.Popen([
        web_capture_exe, 
        "screenshot", 
        "--csv", output_csv,
        "--output", PNG_DIR
    ])
    # 呼叫 web_capture.exe 輸出 HTML 至 HTML_DIR
    thread_safe_log("啟動 web_capture [html] 模式...", text_widget, root)
    proc_html = subprocess.Popen([
        web_capture_exe, 
        "html", 
        "--csv", output_csv,
        "--output", HTML_DIR
    ])
    proc_screenshot.wait()
    proc_html.wait()
    thread_safe_log("web_capture 截圖/HTML 已完成。", text_widget, root)

    # 嵌入圖片：將 PNG 從 PNG_DIR 複製到 HTML_DIR，並在 HTML 檔中插入 <img> 標籤
    # embed_image_in_html(HTML_DIR, PNG_DIR, root, text_widget)

    # 生成報告 TXT
    total_count = 0
    try:
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            for row in reader:
                if row and row[0].strip():
                    total_count += 1
    except Exception as e:
        thread_safe_log(f"讀取 total.csv 失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    today_str = time.strftime("%Y%m%d", time.localtime())
    report_txt_name = f"LineReport.txt"
    report_txt_path = os.path.join(OUTPUT_DIR, report_txt_name)
    thread_safe_log(f"準備產生報告 TXT => {report_txt_name}", text_widget, root)
    result = generate_domain_report_txt(output_csv, report_txt_path)
    thread_safe_log(result, text_widget, root)

    # 複製 total.csv 並重新命名為 YYYYMMDD.申報.[品牌](N筆).csv
    try:
        report_csv_path = copy_total_csv_report(output_csv, OUTPUT_DIR)
        thread_safe_log(f"已複製 total.csv 並生成報告 CSV：{os.path.basename(report_csv_path)}", text_widget, root)
    except Exception as e:
        thread_safe_log(f"複製 total.csv 失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    thread_safe_log("所有任務已完成。", text_widget, root)
    root.after(0, progressbar.stop)

def one_click_complete(text_widget, root, progressbar):
    progressbar.start(10)
    t = threading.Thread(target=run_long_task, args=(text_widget, root, progressbar))
    t.start()

# -------------------------------------
# 建立美化後的 UI (駭客風：炭黑 + 駭客螢光綠)
# -------------------------------------
def main_ui():
    root = tk.Tk()
    root.title("CyberTracker Control Panel")
    root.geometry("720x550")

    HACKER_BLACK = "#1E1E1E"
    HACKER_GREEN = "#00FF00"

    root.configure(bg=HACKER_BLACK)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TFrame", background=HACKER_BLACK)
    style.configure("TLabel", background=HACKER_BLACK, foreground=HACKER_GREEN, font=("Courier New", 16))
    style.configure("TButton", background=HACKER_BLACK, foreground=HACKER_GREEN, font=("Courier New", 14, "bold"))
    style.configure("green.Horizontal.TProgressbar", troughcolor=HACKER_BLACK, bordercolor=HACKER_BLACK,
                    background=HACKER_GREEN, lightcolor=HACKER_GREEN, darkcolor=HACKER_GREEN)

    header = ttk.Label(root, text="CYBERTRACKER CONTROL PANEL")
    header.pack(pady=10)

    btn_frame = ttk.Frame(root)
    btn_frame.pack(pady=5)

    def create_hacker_text(parent):
        txt = ScrolledText(parent, wrap=tk.WORD, width=85, height=18,
                           font=("Courier New", 12), bg=HACKER_BLACK, fg=HACKER_GREEN,
                           insertbackground=HACKER_GREEN)
        return txt

    text_widget = create_hacker_text(root)
    text_widget.pack(padx=15, pady=5)

    btn_import = ttk.Button(btn_frame, text="匯入資料 (CSV)",
                            command=lambda: import_data(text_widget, root))
    btn_import.grid(row=0, column=0, padx=15, pady=10)

    btn_clear = ttk.Button(btn_frame, text="清空 all_csv",
                           command=lambda: clear_all_csv(text_widget, root))
    btn_clear.grid(row=0, column=1, padx=15, pady=10)

    btn_convert = ttk.Button(btn_frame, text="CSV → XLSX 轉檔",
                             command=lambda: convert_csv_button(text_widget, root))
    btn_convert.grid(row=0, column=2, padx=15, pady=10)

    btn_complete = ttk.Button(btn_frame, text="一鍵完成",
                              command=lambda: one_click_complete(text_widget, root, progressbar))
    btn_complete.grid(row=0, column=3, padx=15, pady=10)

    btn_report = ttk.Button(btn_frame, text="生成通報 (TXT)",
                            command=lambda: generate_report(text_widget, root))
    btn_report.grid(row=0, column=4, padx=15, pady=10)

    progressbar = ttk.Progressbar(root, style="green.Horizontal.TProgressbar",
                                  orient=tk.HORIZONTAL, mode="indeterminate", length=520)
    progressbar.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main_ui()
