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

# 模擬手機 UA
MOBILE_UA = ("Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) "
             "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 "
             "Mobile/15E148 Safari/604.1")

class ScreenshotTaker:
    def __init__(self, csv_file, output_dir, headless=True, is_mobile=False):
        self.csv_file = csv_file
        self.output_dir = output_dir

        self.options = Options()
        if headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--no-sandbox')

        # 若要模擬手機
        if is_mobile:
            self.options.add_argument(f'--user-agent={MOBILE_UA}')
            # iPhone 12 視窗大小
            self.window_width = 390
            self.window_height = 844
        else:
            self.window_width = 1280
            self.window_height = 2000

        self.driver = webdriver.Chrome(options=self.options)
        self.driver.set_window_size(self.window_width, self.window_height)

    def run(self, zoom=80, load_wait=3):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir, exist_ok=True)

        with open(self.csv_file, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader, None)
            for row in reader:
                if len(row) < 5:
                    continue
                url = row[2].strip()  # 第 3 欄
                domain = row[4].strip()  # 第 5 欄
                if not url:
                    continue

                try:
                    self.driver.get(url)
                    time.sleep(load_wait)

                    title = self.driver.title
                    timestamp = time.strftime("%m%d%H%M")

                    safe_dom = safe_filename(domain)
                    safe_tit = safe_filename(title)
                    png_name = f"{timestamp}_{safe_dom}_{safe_tit}.png"
                    png_path = os.path.join(self.output_dir, png_name)

                    # Banner + 縮放
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

                    self.driver.save_screenshot(png_path)
                    print(f"[Screenshot] 已截圖：{png_path}")
                except Exception as e:
                    print(f"[ScreenshotTaker][錯誤] {url} - {e}")

        self.driver.quit()

class HTMLDownloader:
    def __init__(self, csv_file, output_dir, headless=True, is_mobile=False):
        self.csv_file = csv_file
        self.output_dir = output_dir

        self.options = Options()
        if headless:
            self.options.add_argument('--headless')
            self.options.add_argument('--no-sandbox')

        if is_mobile:
            self.options.add_argument(f'--user-agent={MOBILE_UA}')
            self.window_width = 390
            self.window_height = 844
        else:
            self.window_width = 1280
            self.window_height = 2000

        self.driver = webdriver.Chrome(options=self.options)
        self.driver.set_window_size(self.window_width, self.window_height)

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

                    title = self.driver.title
                    timestamp = time.strftime("%m%d%H%M")
                    safe_dom = safe_filename(domain)
                    safe_tit = safe_filename(title)

                    html_name = f"{timestamp}_{safe_dom}_{safe_tit}.html"
                    html_path = os.path.join(self.output_dir, html_name)
                    with open(html_path, "w", encoding="utf-8") as hf:
                        hf.write(self.driver.page_source)
                    print(f"[HTML] 已存檔：{html_path}")
                except Exception as e:
                    print(f"[HTMLDownloader][錯誤] {url} - {e}")

        self.driver.quit()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=["screenshot", "html"])
    parser.add_argument("--csv", default=None, help="CSV 檔案路徑")
    parser.add_argument("--output", default=None, help="輸出資料夾")
    parser.add_argument("--no-headless", action="store_false", dest="headless")
    parser.add_argument("--mobile", action="store_true", help="是否模擬手機")
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    csv_file = args.csv or os.path.join(base_dir, "csv_stuff", "total.csv")

    if args.output:
        out_dir = args.output
    else:
        out_dir = os.path.join(base_dir, "output_html" if args.mode=="html" else "output_png")

    if args.mode == "screenshot":
        taker = ScreenshotTaker(csv_file, out_dir, headless=args.headless, is_mobile=args.mobile)
        taker.run()
    else:
        downloader = HTMLDownloader(csv_file, out_dir, headless=args.headless, is_mobile=args.mobile)
        downloader.run()

if __name__=="__main__":
    main()
