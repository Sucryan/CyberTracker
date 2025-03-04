import os
import sys
import csv
import time
import argparse
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def safe_filename(name):
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "_")
    return name

# 預設手機模擬使用的裝置設定（這裡以 Pixel 2 為例）
MOBILE_EMULATION = { "deviceName": "Pixel 2" }

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
            # 使用 Chrome 的 mobileEmulation 功能
            self.options.add_experimental_option("mobileEmulation", MOBILE_EMULATION)
            # 不用設定固定視窗大小，Chrome 會自動套用裝置解析度
        else:
            self.options.add_argument(f"--window-size={window_width},{window_height}")

        self.driver = webdriver.Chrome(options=self.options)

    def run(self, zoom=80, load_wait=3):
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

                    self.driver.execute_script(f"""
                        var urlBanner = document.createElement('div');
                        urlBanner.style.position = 'fixed';
                        urlBanner.style.top = '0';
                        urlBanner.style.left = '0';
                        urlBanner.style.width = '100%';
                        urlBanner.style.padding = '10px';
                        urlBanner.style.backgroundColor = '#f0f0f0';
                        urlBanner.style.color = '#000';
                        urlBanner.style.zIndex = '999999';
                        urlBanner.style.fontSize = '20px';
                        urlBanner.style.fontWeight = 'bold';
                        urlBanner.style.fontFamily = 'sans-serif';
                        urlBanner.innerText = '{url}';
                        document.body.appendChild(urlBanner);
                    """)
                    self.driver.execute_script(f"document.body.style.zoom='{zoom}%'")
                    time.sleep(2)

                    self.driver.save_screenshot(screenshot_path)
                    print(f"已截圖: {screenshot_path}")
                except Exception as e:
                    print(f"[ScreenshotTaker][錯誤] 網域: {domain} / URL: {url}")
                    print("詳細錯誤：", e)
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
                    print(f"[HTMLDownloader][錯誤] 網域: {domain} / URL: {url}")
                    print("詳細錯誤：", e)
        print("[HTMLDownloader] CSV 中所有網址原始檔下載完成！")
        self.driver.quit()


def main():
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

if __name__ == "__main__":
    main()
