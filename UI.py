
import datetime
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
import traceback
import re
import os
today_str = datetime.datetime.now().strftime("%Y%m%d")
now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def on_closing():
    if messagebox.askokcancel("離開", "確定要退出程式嗎？"):
        # 設置中止旗標，告知所有 thread 停止工作
        global stop_all_threads
        stop_all_threads = True
        # 可以等待各 thread 收工或直接強制退出
        os._exit(0)

# --------------------------
# 全域handler設定
# --------------------------

# 假設 BASE_DIR 為你程式所在的資料夾
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def global_exception_handler(exctype, value, tb):
    # 建立 error log 資料夾
    error_log_folder = os.path.join(BASE_DIR, "error log")
    os.makedirs(error_log_folder, exist_ok=True)
    
    # 定義詳細錯誤與摘要檔案路徑
    detailed_log_file = os.path.join(error_log_folder, "detailed_error.log")
    summary_log_file = os.path.join(error_log_folder, "error_summary.txt")
    
    # 將完整 traceback 寫入 detailed_error.log
    with open(detailed_log_file, "w", encoding="utf-8") as f:
        traceback.print_exception(exctype, value, tb, file=f)
    
    # 產生錯誤摘要：從 traceback 中找出所有網址
    error_trace = "".join(traceback.format_exception(exctype, value, tb))
    # 使用正規表示式找出 http:// 或 https:// 開頭的網址
    url_pattern = r'(https?://[^\s]+)'
    urls = re.findall(url_pattern, error_trace)
    
    with open(summary_log_file, "w", encoding="utf-8") as f:
        if urls:
            for url in urls:
                f.write(url + "\n")
        else:
            f.write("未找到錯誤中的網址。請參考 detailed_error.log 獲得完整資訊。\n")
    
    # 顯示簡單提示給使用者
    print("發生錯誤，詳細資訊請參閱 error log 資料夾。")

# 設定全域例外處理器
sys.excepthook = global_exception_handler



# --------------------------
# 全域路徑設定
# --------------------------
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
DESKTOP_DIR = winshell.desktop()  # 例如 "C:\Users\YourName\OneDrive\桌面"
OUTPUT_DIR = os.path.join(DESKTOP_DIR, f"CyberTrackerOutput_{now_str}")
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
    # 如果讀到「數產」，直接回傳「電商」
    if "數產" in text:
        return "電商"
    if "台灣樂天" in text:
        return "台灣樂天"
    if text.startswith("疑似假冒"):
        return text[4:6] if len(text) >= 6 else "電商"
    if text[:2] in ["偽冒", "假冒"]:
        return text[2:4] if len(text) >= 4 else "電商"
    return "電商"

# 取得品牌資訊與總筆數的輔助函式
def get_vendor_info(total_csv_path):
    total_count = 0
    vendor_set = set()
    try:
        with open(total_csv_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)  # 略過表頭
            for row in reader:
                if row and row[0].strip():
                    total_count += 1
                    # 使用 extract_brand_from_row 函式取得品牌（請確保此函式已定義）
                    brand = extract_brand_from_row(row)
                    vendor_set.add(brand)
    except Exception as e:
        raise Exception(f"讀取 total.csv 失敗：{e}")
    vendor_str = "+".join(sorted(vendor_set)) if vendor_set else "電商"
    return total_count, vendor_str

# 修改後的複製 total.csv 產生 reference 檔案
def copy_total_csv_report(total_csv_path, output_folder):
    total_count = 0
    vendor_set = set()

    try:
        with open(total_csv_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)  # 略過表頭
            for row in reader:
                if row and row[0].strip():
                    total_count += 1
                    brand = extract_brand_from_row(row)
                    vendor_set.add(brand)
    except Exception as e:
        raise Exception(f"讀取 total.csv 失敗：{e}")

    # 舊有邏輯產生 vendor_str 已不再使用於檔名中
    today_str = time.strftime("%Y%m%d", time.localtime())
    new_filename = f"reference_{today_str}.csv"  # 改成 reference_{日期}.csv
    new_file_path = os.path.join(output_folder, new_filename)

    try:
        with open(total_csv_path, "r", encoding="utf-8") as src, \
             open(new_file_path, "w", encoding="big5", newline="") as dst:
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
    output_xlsx = os.path.join(XLSX_DIR, f"通報TWNIC詐騙網址彙整表_{today_str}.xlsx")
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
# ========== 郵件相關函式 (改用單純 print 而非 thread_safe_log) ==========
# Bolin已在網頁後端完成。
'''
def load_email_config():
    config_path = os.path.join(BASE_DIR, "email_config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        return config["account"], config["password"]
    except Exception as e:
        raise Exception(f"讀取 email_config.json 失敗：{e}")

SENDER_EMAIL, SENDER_PASSWORD = load_email_config()
SMTP_SERVER = "smtp.gmail.com"  # 如需使用其他 SMTP，請修改
SMTP_PORT = 587

def send_email_with_attachment(to_email: str, subject: str, body: str, attachment_path: str):
    print(f"[DEBUG] 正在寄給：{to_email}")
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        filename = os.path.basename(attachment_path)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        print(f"[Email] 已寄送給：{to_email}")
    except Exception as e:
        print(f"[Email] 寄送給 {to_email} 時發生錯誤：{e}")
'''
'''
def create_email_report(input_csv_path: str, output_report_path: str):
    try:
        # 以 UTF-8 讀取原始 CSV 檔（假設原檔含表頭）
        df = pd.read_csv(input_csv_path, encoding="utf-8")
        # 取出第 5 欄與第 7 欄（欄位索引分別為 4 與 6）
        df_report = df.iloc[:, [4, 6]].copy()
        # 重新命名欄位
        df_report.columns = ["域名", "含子域名"]

        # 匯出 CSV，不輸出 index，並指定 Windows 換行格式
        df_report.to_csv(output_report_path, index=False, encoding="big5", line_terminator="\r\n")
        print(f"[Email Report] 已產生報告：{os.path.basename(output_report_path)}")
    except Exception as e:
        print(f"[Email Report] 產生失敗：{e}")
        raise
'''
'''
def send_email_reports(json_path: str, total_csv_path: str, output_dir: str):
    if not os.path.exists(json_path):
        print(f"[Email] 找不到收件者名單：{json_path}")
        return

    with open(json_path, "r", encoding="utf-8") as f:
        email_dict = json.load(f)

    today_str = time.strftime("%Y%m%d", time.localtime())
    report_csv = os.path.join(output_dir, f"email_report_example_{today_str}.csv")

    # 產生報表
    create_email_report(total_csv_path, report_csv)

    subject = "封鎖通報報表"
    body = (
        "您好：\n\n"
        "附件為今日產生的封鎖通報報表（包含「域名」及「含子域名」欄位）。\n"
        "請查閱附件內容，謝謝！\n\n"
        "此致\n敬上"
    )

    # 逐一寄送給 emails.json 裡的收件者
    for _, recipient_email in email_dict.items():
        send_email_with_attachment(recipient_email, subject, body, report_csv)

def send_email_report_ui():
    """
    簡單版本：不傳 text_widget, root，
    只用 print(...) 顯示寄送訊息於終端機。
    """
    csv_stuff_dir = os.path.join(BASE_DIR, "csv_stuff")
    total_csv = os.path.join(csv_stuff_dir, "total.csv")
    if not os.path.exists(total_csv):
        print("[Email] 錯誤：找不到 total.csv！")
        return

    email_json_path = os.path.join(BASE_DIR, "emails.json")
    try:
        send_email_reports(email_json_path, total_csv, OUTPUT_DIR)
        print("[Email] email report 寄送流程完成。")
    except Exception as e:
        print(f"[Email] 寄送 email report 時發生錯誤：{e}")
'''
# --------------------------
# 移除 embed_image_in_html 功能（不再執行）
# --------------------------

def run_long_task(text_widget, root, progressbar):
    # 建立必要資料夾
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
        print("錯誤：找不到 merge_csv.exe！")
        thread_safe_log("錯誤：找不到 merge_csv.exe！", text_widget, root)
        root.after(0, progressbar.stop)
        return

    input_dir = os.path.join(BASE_DIR, "all_csv")
    output_csv = os.path.join(csv_stuff_dir, "total.csv")
    output_csv2 = os.path.join(csv_stuff_dir, "domain.csv")
    print("開始執行 merge_csv...")
    thread_safe_log("開始執行 merge_csv...", text_widget, root)
    try:
        subprocess.run([merge_exe, "--input-dir", input_dir, "--output-file", output_csv], check=True)
    except Exception as e:
        print(f"merge_csv 執行失敗：{e}")
        thread_safe_log(f"merge_csv 執行失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    if not os.path.isfile(output_csv):
        print(f"錯誤：找不到合併後的 total.csv：{output_csv}")
        thread_safe_log(f"錯誤：找不到合併後的 total.csv：{output_csv}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    print(f"合併後 CSV 檔案：{output_csv}")
    thread_safe_log(f"合併後 CSV 檔案：{output_csv}", text_widget, root)
    
    try:
        subprocess.run([merge_exe, "--input-dir", input_dir, "--output-file", output_csv2], check=True)
    except Exception as e:
        print(f"merge_csv 執行失敗：{e}")
        thread_safe_log(f"merge_csv 執行失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    if not os.path.isfile(output_csv2):
        print(f"錯誤：找不到合併後的 domain.csv：{output_csv2}")
        thread_safe_log(f"錯誤：找不到合併後的 domain.csv：{output_csv2}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    print(f"合併後 CSV 檔案：{output_csv2}")
    thread_safe_log(f"合併後 CSV 檔案：{output_csv2}", text_widget, root)

    if not os.path.exists(web_capture_exe):
        print("錯誤：找不到 web_capture.exe！")
        thread_safe_log("錯誤：找不到 web_capture.exe！", text_widget, root)
        root.after(0, progressbar.stop)
        return

    def run_capture(cmd, task_name):
        try:
            subprocess.run(cmd, check=True)
            print(f"{task_name} 完成。")

            thread_safe_log(f"{task_name} 完成。", text_widget, root)
        except Exception as e:
            print(f"{task_name} 執行失敗：{e}")
            thread_safe_log(f"{task_name} 執行失敗：{e}", text_widget, root)

    # 桌面版 (laptop) 指令 (不帶 --mobile)
    lap_screenshot_total_cmd = [web_capture_exe, "screenshot", "--csv", output_csv, "--output", LAP_PNG_DIR]
    lap_html_total_cmd = [web_capture_exe, "html", "--csv", output_csv, "--output", LAP_HTML_DIR]
    # 手機版 (mobile) 指令 (帶 --mobile)
    mob_screenshot_total_cmd = [web_capture_exe, "screenshot", "--csv", output_csv, "--output", MOB_PNG_DIR, "--mobile"]
    mob_html_total_cmd = [web_capture_exe, "html", "--csv", output_csv, "--output", MOB_HTML_DIR, "--mobile"]

    # 桌面版 (laptop) 指令 (不帶 --mobile)
    lap_screenshot_domain_cmd = [web_capture_exe, "screenshot", "--csv", output_csv2, "--output", LAP_PNG_DIR]
    lap_html_domain_cmd = [web_capture_exe, "html", "--csv", output_csv2, "--output", LAP_HTML_DIR]
    # 手機版 (mobile) 指令 (帶 --mobile)
    mob_screenshot_domain_cmd = [web_capture_exe, "screenshot", "--csv", output_csv2, "--output", MOB_PNG_DIR, "--mobile"]
    mob_html_domain_cmd = [web_capture_exe, "html", "--csv", output_csv2, "--output", MOB_HTML_DIR, "--mobile"]

    threads = []
    for cmd, task in [(lap_screenshot_total_cmd, "桌面截圖(subdomain)"),
                  (lap_html_total_cmd, "桌面 HTML(subdomain)"),
                  (mob_screenshot_total_cmd, "手機截圖(subdomain)"),
                  (mob_html_total_cmd, "手機 HTML(subdomain)"),
                  (lap_screenshot_domain_cmd, "桌面截圖(domain)"),
                  (lap_html_domain_cmd, "桌面 HTML(domain)"),
                  (mob_screenshot_domain_cmd, "手機截圖(domain)"),
                  (mob_html_domain_cmd, "手機 HTML(domain)"),
                  ]:
        t = threading.Thread(target=run_capture, args=(cmd, task))
        t.start()
        threads.append(t)
    for t in threads:
        t.join()
    print("web_capture 所有任務已完成。")
    thread_safe_log("web_capture 所有任務已完成。", text_widget, root)

    # 生成報告 TXT
    total_count = 0
    try:
        with open(output_csv, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if row and row[0].strip():
                    total_count += 1
    except Exception as e:
        print(f"讀取 total.csv 失敗：{e}")
        thread_safe_log(f"讀取 total.csv 失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    today_str = time.strftime("%Y%m%d", time.localtime())
    report_txt_name = "LineReport.txt"
    report_txt_path = os.path.join(OUTPUT_DIR, report_txt_name)
    print(f"準備產生報告 TXT => {report_txt_name}")
    thread_safe_log(f"準備產生報告 TXT => {report_txt_name}", text_widget, root)
    result = generate_domain_report_txt(output_csv, report_txt_path)
    print(result)
    thread_safe_log(result, text_widget, root)

    try:
        # 產生 reference CSV，原本 copy_total_csv_report 的回傳檔案
        report_csv_path = copy_total_csv_report(output_csv, OUTPUT_DIR)
        print(f"已複製 total.csv 並生成報告 CSV：{os.path.basename(report_csv_path)}")
        thread_safe_log(f"已複製 total.csv 並生成報告 CSV：{os.path.basename(report_csv_path)}", text_widget, root)
    except Exception as e:
        print(f"複製 total.csv 失敗：{e}")
        thread_safe_log(f"複製 total.csv 失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return


    # ---------------------------
    # 新增：自動產生並寄送 email report
    # ---------------------------
    # 產生 email report CSV，存放於 OUTPUT_DIR 下，檔名格式：email_report_example_{YYYYMMDD}.csv
    
    try:
        total_count, vendor_str = get_vendor_info(output_csv)
    except Exception as e:
        print(f"取得 vendor 資訊失敗：{e}")
        thread_safe_log(f"取得 vendor 資訊失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return

    # email report 檔案名稱改為申報命名法則：{YYYYMMDD}.申報.{vendor_str}({total_count}筆).csv
    email_report_name = f"{today_str}.申報.{vendor_str}({total_count}筆).csv"
    email_report = os.path.join(OUTPUT_DIR, email_report_name)

    try:
        with open(output_csv, "r", encoding="utf-8") as fin, \
            open(email_report, "w", encoding="big5", newline="") as fout:
            reader = csv.reader(fin)
            writer = csv.writer(fout)
            writer.writerow(["域名", "含子域名"])
            next(reader, None)
            for row in reader:
                if len(row) >= 7:
                    domain_val = row[4].strip()
                    subdomain_val = row[6].strip()
                    writer.writerow([domain_val, subdomain_val])
        print(f"已產生 email report：{os.path.basename(email_report)}")
        thread_safe_log(f"已產生 email report：{os.path.basename(email_report)}", text_widget, root)
    except Exception as e:
        print(f"生成 email report 失敗：{e}")
        thread_safe_log(f"生成 email report 失敗：{e}", text_widget, root)
        root.after(0, progressbar.stop)
        return
    '''
    # 從 emails.json 中讀取收件者並寄送報告
    email_json_path = os.path.join(BASE_DIR, "emails.json")

    if not os.path.exists(email_json_path):
        print(f"錯誤：找不到 emails.json：{email_json_path}")
    else:
        try:
            with open(email_json_path, "r", encoding="utf-8") as f:
                email_dict = json.load(f)
            subject = "封鎖通報報表"
            
            # 定義 TXT 檔案路徑並讀取內容
            body_path = os.path.join(BASE_DIR, "email_body.txt")
            if os.path.exists(body_path):
                with open(body_path, "r", encoding="utf-8") as body_file:
                    body = body_file.read()
            else:
                body = "錯誤：郵件正文檔案不存在，請檢查 email_body.txt"
                print(f"錯誤：找不到 email_body.txt：{body_path}")
                thread_safe_log(f"錯誤：找不到 email_body.txt：{body_path}", text_widget, root)
            
            for _, recipient in email_dict.items():
                send_email_with_attachment(recipient, subject, body, email_report)
            thread_safe_log("寄送成功!", text_widget, root)
        except Exception as e:
            print(f"寄送 email report 時發生錯誤：{e}")
            thread_safe_log(f"寄送 email report 時發生錯誤：{e}", text_widget, root)
    '''
    print("所有任務已完成。")
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
    '''
    btn_report = ttk.Button(btn_frame, text="生成通報 (TXT)",
                            command=lambda: generate_report(text_widget, root))
    btn_report.grid(row=0, column=4, padx=15, pady=10)
    '''
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