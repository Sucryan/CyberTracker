import os
import sys
import csv
import time
import argparse
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def safe_filename(name):
    """移除檔名中不合法的字元"""
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "_")
    return name

class ScreenshotTaker:
    def __init__(self, csv_file, output_dir, headless=True, window_width=1280, window_height=2000):
        self.csv_file = csv_file
        self.output_dir = output_dir
        self.headless = headless

        self.options = Options()
        if self.headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--no-sandbox')

        self.driver = webdriver.Chrome(options=self.options)
        self.driver.set_window_size(window_width, window_height)

    def run(self, zoom=80, load_wait=3):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        with open(self.csv_file, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)  # 略過第一行

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

                    safe_domain = safe_filename(domain)
                    safe_title = safe_filename(page_title)
                    screenshot_filename = f"{timestamp}_{safe_domain}_{safe_title}.png"
                    screenshot_path = os.path.join(self.output_dir, screenshot_filename)

                    # 插入簡易 banner
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
                        urlBanner.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
                        urlBanner.style.textAlign = 'center';
                        urlBanner.innerText = '{url}';
                        document.body.appendChild(urlBanner);
                    """)
                    # 縮放
                    self.driver.execute_script(f"document.body.style.zoom='{zoom}%'")
                    time.sleep(2)

                    self.driver.save_screenshot(screenshot_path)
                    print(f"已截圖: {screenshot_path}")

                except Exception as e:
                    print("[ScreenshotTaker][錯誤]", e)

        self.driver.quit()


class HTMLDownloader:
    def __init__(self, csv_file, output_dir, headless=True, window_width=1280, window_height=2000):
        self.csv_file = csv_file
        self.output_dir = output_dir
        self.headless = headless

        self.options = Options()
        if self.headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--no-sandbox')

        self.driver = webdriver.Chrome(options=self.options)
        self.driver.set_window_size(window_width, window_height)

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

                    safe_domain = safe_filename(domain)
                    safe_title = safe_filename(page_title)
                    html_filename = f"{timestamp}_{safe_domain}_{safe_title}.html"
                    html_path = os.path.join(self.output_dir, html_filename)

                    with open(html_path, "w", encoding="utf-8") as f_html:
                        f_html.write(self.driver.page_source)
                    print(f"已存檔網頁原始碼: {html_path}")

                except Exception as e:
                    print("[HTMLDownloader][錯誤]", e)

        self.driver.quit()


def main():
    parser = argparse.ArgumentParser(description="Web Capture Tool")
    parser.add_argument("mode", choices=["screenshot", "html"], help="選擇功能: screenshot 或 html")
    parser.add_argument("--csv", type=str, default=None, help="CSV 路徑 (預設=./csv_stuff/total.csv)")
    parser.add_argument("--output", type=str, default=None, help="輸出資料夾 (預設=./output_screenshot 或 ./output_html)")
    parser.add_argument("--no-headless", action="store_false", dest="headless",
                        help="停用 headless 模式，預設為 headless=True")
    args = parser.parse_args()

    # 預設路徑
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    csv_path = args.csv if args.csv else os.path.join(base_dir, "csv_stuff", "total.csv")

    if args.output:
        out_dir = args.output
    else:
        if args.mode == "screenshot":
            out_dir = os.path.join(base_dir, "output_screenshot")
        else:
            out_dir = os.path.join(base_dir, "output_html")

    # 預設 headless = True，只有使用者加上 --no-headless 才會顯示瀏覽器
    headless = args.headless

    if args.mode == "screenshot":
        taker = ScreenshotTaker(csv_file=csv_path, output_dir=out_dir, headless=headless)
        taker.run()
    else:
        downloader = HTMLDownloader(csv_file=csv_path, output_dir=out_dir, headless=headless)
        downloader.run()


if __name__ == "__main__":
    main()
