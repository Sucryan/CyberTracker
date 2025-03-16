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
import winshell  # pip install winshell

# --------------------------
# 全域路徑設定
# --------------------------
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
DESKTOP_DIR = winshell.desktop()  # 例如 "C:\Users\YourName\OneDrive\桌面"
OUTPUT_DIR = os.path.join(DESKTOP_DIR, "CyberTrackerOutput")
# 分別存放桌面版與手機版輸出結果
LAPTOP_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "laptop")
MOBILE_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "mobile")
# 桌面版中：分別存放 HTML 與 PNG
LAP_HTML_DIR = os.path.join(LAPTOP_OUTPUT_DIR, "html")
LAP_PNG_DIR = os.path.join(LAPTOP_OUTPUT_DIR, "png")
# 手機版中：分別存放 HTML 與 PNG
MOB_HTML_DIR = os.path.join(MOBILE_OUTPUT_DIR, "html")
MOB_PNG_DIR = os.path.join(MOBILE_OUTPUT_DIR, "png")
# 若有 XLSX 轉檔功能
XLSX_DIR = os.path.join(OUTPUT_DIR, "xlsx")

# --------------------------
# 輔助函式
# --------------------------
def log(msg, text_widget):
    text_widget.insert(tk.END, msg + "\n")
    text_widget.see(tk.END)

def thread_safe_log(msg, text_widget, root):
    root.after(0, lambda: log(msg, text_widget))

def extract_brand_from_row(row):
    """
    從第二行的「網站」欄（假設為 row[1]）提取品牌：
      1) 若內容以「疑似假冒」開頭，則取後面兩個字。
      2) 若以「偽冒」或「假冒」開頭，則取後兩個字。
      3) 否則回傳 "電商"
    """
    if len(row) < 2:
        return "電商"
    text = row[1].strip()
    if text.startswith("疑似假冒"):
        return text[4:6] if len(text) >= 6 else "電商"
    if text[:2] in ["偽冒", "假冒"]:
        return text[2:4] if len(text) >= 4 else "電商"
    return "電商"

def copy_total_csv_report(total_csv_path, output_folder):
    """
    讀取 total.csv，計算總筆數 (第一列為表頭，從第二列起為資料)。
    從 total.csv 中解析每一筆資料的品牌（呼叫 extract_brand_from_row），
    並命名為 YYYYMMDD.申報.(廠商1+廠商2+...)(N筆).csv，以 UTF-8 輸出。
    """
    total_count = 0
    vendor_set = set()

    try:
        with open(total_csv_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)  # 略過表頭
            for row in reader:
                # 如果該列有資料 (row[0] 不是空的)，就計數並提取品牌
                if row and row[0].strip():
                    total_count += 1
                    brand = extract_brand_from_row(row)  # 使用先前定義好的函式
                    vendor_set.add(brand)
    except Exception as e:
        raise Exception(f"讀取 total.csv 失敗：{e}")

    # 如果有收集到品牌，則以 "+" 連接，否則預設為 "電商"
    vendor_str = "+".join(sorted(vendor_set)) if vendor_set else "電商"

    today_str = time.strftime("%Y%m%d", time.localtime())
    new_filename = f"{today_str}.申報.{vendor_str}({total_count}筆).csv"
    new_file_path = os.path.join(output_folder, new_filename)

    # 將 total.csv 原始內容直接複製到新檔名中
    try:
        with open(total_csv_path, "r", encoding="utf-8") as src, \
             open(new_file_path, "w", encoding="utf-8", newline="") as dst:
            shutil.copyfileobj(src, dst)
    except Exception as e:
        raise Exception(f"寫入報告檔失敗：{e}")

    return new_file_path


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

# --------------------------
# 移除 embed_image_in_html 功能（不再執行）
# --------------------------

def run_long_task(text_widget, root, progressbar):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(LAPTOP_OUTPUT_DIR, exist_ok=True)
    os.makedirs(MOBILE_OUTPUT_DIR, exist_ok=True)
    os.makedirs(LAP_HTML_DIR, exist_ok=True)
    os.makedirs(LAP_PNG_DIR, exist_ok=True)
    os.makedirs(MOB_HTML_DIR, exist_ok=True)
    os.makedirs(MOB_PNG_DIR, exist_ok=True)
    csv_stuff_dir = os.path.join(BASE_DIR, "csv_stuff")
    os.makedirs(csv_stuff_dir, exist_ok=True)

    merge_exe = os.path.join(BASE_DIR, "merge_csv.exe")
    web_capture_exe = os.path.join(BASE_DIR, "web_capture.exe")

    if not os.path.exists(merge_exe):
        thread_safe_log("錯誤：找不到 merge_csv.exe！", text_widget, root)
        root.after(0, progressbar.stop)
        return

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

    def run_capture(cmd, task_name):
        try:
            subprocess.run(cmd, check=True)
            thread_safe_log(f"{task_name} 完成。", text_widget, root)
        except Exception as e:
            thread_safe_log(f"{task_name} 執行失敗：{e}", text_widget, root)

    lap_screenshot_cmd = [
        web_capture_exe,
        "screenshot",
        "--csv", output_csv,
        "--output", LAP_PNG_DIR
    ]
    lap_html_cmd = [
        web_capture_exe,
        "html",
        "--csv", output_csv,
        "--output", LAP_HTML_DIR
    ]
    mob_screenshot_cmd = [
        web_capture_exe,
        "screenshot",
        "--csv", output_csv,
        "--output", MOB_PNG_DIR,
        "--mobile"
    ]
    mob_html_cmd = [
        web_capture_exe,
        "html",
        "--csv", output_csv,
        "--output", MOB_HTML_DIR,
        "--mobile"
    ]

    threads = []
    for cmd, task in [(lap_screenshot_cmd, "桌面截圖"),
                      (lap_html_cmd, "桌面 HTML"),
                      (mob_screenshot_cmd, "手機截圖"),
                      (mob_html_cmd, "手機 HTML")]:
        t = threading.Thread(target=run_capture, args=(cmd, task))
        t.start()
        threads.append(t)
    for t in threads:
        t.join()
    thread_safe_log("web_capture 所有任務已完成。", text_widget, root)

    total_count = 0
    try:
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)
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

# --------------------------
# 主 UI 程式
# --------------------------
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

    # 當使用者按下視窗右上角 X 時，直接終止整個程式
    def on_closing():
        if messagebox.askokcancel("離開", "確定要退出程式嗎？"):
            # 終止所有背景工作與子程序，直接退出
            os._exit(0)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main_ui()
