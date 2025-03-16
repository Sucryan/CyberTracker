import os
import sys
import csv
import time
import argparse
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from PIL import Image, ImageDraw, ImageFont

def safe_filename(name):
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "_")
    return name

# 預設手機模擬使用的裝置設定（以 Pixel 2 為例）
MOBILE_EMULATION = { "deviceName": "Pixel 2" }

def add_url_banner(screenshot_path, url, banner_height=50, banner_color="#f0f0f0", text_color="#000"):
    """
    讀取截圖後，於圖片下方新增一塊布幕，並在布幕上加入 URL 文字
    """
    try:
        img = Image.open(screenshot_path)
        width, height = img.size

        # 建立一個新圖片，整體高度增加 banner_height
        new_img = Image.new('RGB', (width, height + banner_height), color=banner_color)
        new_img.paste(img, (0, 0))

        draw = ImageDraw.Draw(new_img)
        try:
            font = ImageFont.truetype("arial.ttf", 20)
        except IOError:
            font = ImageFont.load_default()

        # 設定文字繪製位置 (可根據需求調整邊距)
        text_x = 10
        text_y = height + (banner_height - 20) // 2  # 垂直置中
        draw.text((text_x, text_y), url, fill=text_color, font=font)

        new_img.save(screenshot_path)
        print(f"已加入 URL 橫幅，檔案更新為：{screenshot_path}")
    except Exception as e:
        print(f"[add_url_banner][錯誤] {str(e)}")

class ScreenshotTaker:
    def __init__(self, csv_file, output_dir, headless=True, is_mobile=False, window_width=1280, window_height=2000):
        self.csv_file = csv_file
        self.output_dir = output_dir
        self.headless = headless
        self.is_mobile = is_mobile

        self.options = Options()
        if self.headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--no-sandbox')
        if self.is_mobile:
            self.options.add_experimental_option("mobileEmulation", MOBILE_EMULATION)
        else:
            self.options.add_argument(f"--window-size={window_width},{window_height}")
        self.driver = webdriver.Chrome(options=self.options)

    def run(self, zoom=80, load_wait=3):
        # 錯誤 log 檔案 (存放於輸出資料夾中)
        error_log_file = os.path.join(self.output_dir, "error_log.txt")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        with open(self.csv_file, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            for row in reader:
                if len(row) < 5:
                    continue
                url = row[2].strip()
                domain = row[4].strip()
                if not url:
                    continue
                try:
                    self.driver.get(url)
                    time.sleep(load_wait)
                    page_title = self.driver.title
                    timestamp = time.strftime("%m%d%H%M")
                    safe_dom = safe_filename(domain)
                    safe_tit = safe_filename(page_title)
                    screenshot_filename = f"{timestamp}_{safe_dom}_{safe_tit}.png"
                    screenshot_path = os.path.join(self.output_dir, screenshot_filename)

                    # 設定縮放比例（可依需求調整）
                    self.driver.execute_script(f"document.body.style.zoom='{zoom}%'")
                    time.sleep(2)

                    self.driver.save_screenshot(screenshot_path)
                    print(f"已截圖: {screenshot_path}")

                    # 截圖後進行後處理，加入下方 URL 布幕
                    add_url_banner(screenshot_path, url)
                except Exception as e:
                    error_msg = f"錯誤 - 網域: {domain} / URL: {url} / 錯誤內容: {str(e)}\n"
                    print(f"[ScreenshotTaker][錯誤] {error_msg}")
                    with open(error_log_file, "a", encoding="utf-8") as error_file:
                        error_file.write(error_msg)
        print("[ScreenshotTaker] CSV 中所有網址截圖完成！")
        self.driver.quit()

class HTMLDownloader:
    def __init__(self, csv_file, output_dir, headless=True, is_mobile=False, window_width=1280, window_height=2000):
        self.csv_file = csv_file
        self.output_dir = output_dir
        self.headless = headless
        self.is_mobile = is_mobile

        self.options = Options()
        if self.headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--no-sandbox')
        if self.is_mobile:
            self.options.add_experimental_option("mobileEmulation", MOBILE_EMULATION)
        else:
            self.options.add_argument(f"--window-size={window_width},{window_height}")
        self.driver = webdriver.Chrome(options=self.options)

    def run(self, load_wait=3):
        error_log_file = os.path.join(self.output_dir, "error_log.txt")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        with open(self.csv_file, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            for row in reader:
                if len(row) < 5:
                    continue
                url = row[2].strip()
                domain = row[4].strip()
                if not url:
                    continue
                try:
                    self.driver.get(url)
                    time.sleep(load_wait)
                    page_title = self.driver.title
                    timestamp = time.strftime("%m%d%H%M")
                    safe_dom = safe_filename(domain)
                    safe_tit = safe_filename(page_title)
                    html_filename = f"{timestamp}_{safe_dom}_{safe_tit}.html"
                    html_path = os.path.join(self.output_dir, html_filename)
                    with open(html_path, "w", encoding="utf-8") as html_file:
                        html_file.write(self.driver.page_source)
                    print(f"已存檔網頁原始碼: {html_path}")
                except Exception as e:
                    error_msg = f"錯誤 - 網域: {domain} / URL: {url} / 錯誤內容: {str(e)}\n"
                    print(f"[HTMLDownloader][錯誤] {error_msg}")
                    with open(error_log_file, "a", encoding="utf-8") as error_file:
                        error_file.write(error_msg)
        print("[HTMLDownloader] CSV 中所有網址原始檔下載完成！")
        self.driver.quit()

def main():
    try:
        parser = argparse.ArgumentParser(description="Web Capture Tool")
        parser.add_argument("mode", choices=["screenshot", "html"], help="選擇功能: screenshot 或 html")
        parser.add_argument("--csv", type=str, default=None, help="CSV 檔案路徑 (預設：./csv_stuff/total.csv)")
        parser.add_argument("--output", type=str, default=None, help="輸出資料夾 (預設依模式設定)")
        parser.add_argument("--no-headless", action="store_false", dest="headless", help="停用 headless 模式")
        parser.add_argument("--mobile", action="store_true", help="啟用手機模擬模式")
        args = parser.parse_args()

        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        csv_file = args.csv if args.csv else os.path.join(base_dir, "csv_stuff", "total.csv")
        if args.output:
            out_dir = args.output
        else:
            if args.mode == "screenshot":
                out_dir = os.path.join(base_dir, "output_screenshot")
            else:
                out_dir = os.path.join(base_dir, "output_html")

        if args.mode == "screenshot":
            taker = ScreenshotTaker(csv_file, out_dir, headless=args.headless, is_mobile=args.mobile)
            taker.run()
        else:
            downloader = HTMLDownloader(csv_file, out_dir, headless=args.headless, is_mobile=args.mobile)
            downloader.run()
    except Exception as e:
        error_log_file = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "global_error_log.txt")
        error_msg = f"全域錯誤: {str(e)}\n"
        print(error_msg)
        with open(error_log_file, "a", encoding="utf-8") as error_file:
            error_file.write(error_msg)

if __name__ == "__main__":
    main()
