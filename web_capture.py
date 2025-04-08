import os
import sys
import csv
import time
import argparse
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from PIL import Image, ImageDraw, ImageFont

# 自訂例外與其他輔助函式保持不變
class FacebookPagesException(Exception):
    facebook_urls = []
    def __init__(self, url, status_code):
        self.url = url
        self.status_code = status_code
        super().__init__(f"Facebook page returned status {status_code} for URL: {url}")
        FacebookPagesException.facebook_urls.append(url)

    @classmethod
    def output_urls(cls, output_file):
        with open(output_file, "w", encoding="utf-8") as f:
            for url in cls.facebook_urls:
                f.write(url + "\n")

class HTTPStatusError(Exception):
    def __init__(self, url, status_code):
        self.url = url
        self.status_code = status_code
        super().__init__(f"HTTP error {status_code} for URL: {url}")

def safe_filename(name):
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "_")
    return name

MOBILE_EMULATION = {"deviceName": "Pixel 2"}

MOBILE_EMULATION = {"deviceName": "Pixel 2"}

def add_url_banner(screenshot_path, url, banner_height=50, banner_color="#f0f0f0", text_color="#000"):
    try:
        img = Image.open(screenshot_path)
        width, height = img.size
        new_img = Image.new('RGB', (width, height + banner_height), color=banner_color)
        new_img.paste(img, (0, banner_height))
        draw = ImageDraw.Draw(new_img)
        try:
            font = ImageFont.truetype("arial.ttf", 20)
        except IOError:
            font = ImageFont.load_default()
        text_x = 10
        text_y = (banner_height - 20) // 2
        draw.text((text_x, text_y), url, fill=text_color, font=font)
        new_img.save(screenshot_path)
        print(f"已加入 URL 橫幅，檔案更新為：{screenshot_path}")
    except Exception as e:
        print(f"[add_url_banner][錯誤] {str(e)}")

def get_status_with_retry(url, retries=3, backoff=5):
    attempt = 0
    while attempt < retries:
        try:
            r = requests.head(url, timeout=10)
            if r.status_code == 429:
                print(f"HTTP 429 encountered for URL {url} (attempt {attempt+1}/{retries}), waiting {backoff} seconds...")
                attempt += 1
                if attempt < retries:
                    time.sleep(backoff)
                    continue
            return r
        except Exception as e:
            attempt += 1
            if attempt < retries:
                print(f"HTTP HEAD request failed for URL {url} (attempt {attempt}/{retries}), waiting {backoff} seconds...")
                time.sleep(backoff)
            else:
                raise e
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
        SLEEP_DELAY = 5
        error_log_file = os.path.join(self.output_dir, "error_log.txt")
        os.makedirs(self.output_dir, exist_ok=True)

        # 判斷是否讀取的是 domain CSV（檔名包含 "domain"）
        is_domain_csv = "domain" in os.path.basename(self.csv_file).lower()

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
                    r = get_status_with_retry(url, retries=3, backoff=5)
                    if r.status_code not in [200, 429]:
                        if "facebook.com" in url.lower():
                            raise FacebookPagesException(url, r.status_code)
                        else:
                            raise HTTPStatusError(url, r.status_code)
                    self.driver.get(url)
                    time.sleep(load_wait)
                    page_title = self.driver.title
                    # 改用更精確的時間戳（含秒）
                    timestamp = time.strftime("%m%d%H%M")
                    safe_dom = safe_filename(domain)
                    safe_tit = safe_filename(page_title)
                    if is_domain_csv:
                        base_filename = f"{timestamp}_1_{safe_dom}_{safe_tit}.png"
                    else:
                        base_filename = f"{timestamp}_{safe_dom}_{safe_tit}.png"
                    # 取得唯一的檔案路徑，避免覆蓋
                    screenshot_path = os.path.join(self.output_dir, base_filename)

                    self.driver.execute_script(f"document.body.style.zoom='{zoom}%'")
                    time.sleep(2)
                    self.driver.save_screenshot(screenshot_path)
                    print(f"已截圖: {screenshot_path}")
                    add_url_banner(screenshot_path, url)
                except Exception as e:
                    error_msg = f"錯誤 - 網域: {domain} / URL: {url} / 錯誤內容: {str(e)}\n"
                    print(f"[ScreenshotTaker][錯誤] {error_msg}")
                    with open(error_log_file, "a", encoding="utf-8") as error_file:
                        error_file.write(error_msg)
                finally:
                    time.sleep(SLEEP_DELAY)
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
        SLEEP_DELAY = 5
        error_log_file = os.path.join(self.output_dir, "error_log.txt")
        os.makedirs(self.output_dir, exist_ok=True)

        is_domain_csv = "domain" in os.path.basename(self.csv_file).lower()

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
                    r = get_status_with_retry(url, retries=3, backoff=5)
                    if r.status_code not in [200, 429]:
                        if "facebook.com" in url.lower():
                            raise FacebookPagesException(url, r.status_code)
                        else:
                            raise HTTPStatusError(url, r.status_code)
                    self.driver.get(url)
                    time.sleep(load_wait)
                    page_title = self.driver.title
                    timestamp = time.strftime("%m%d%H%M")
                    safe_dom = safe_filename(domain)
                    safe_tit = safe_filename(page_title)
                    if is_domain_csv:
                        base_filename = f"{timestamp}_1_{safe_dom}_{safe_tit}.html"
                    else:
                        base_filename = f"{timestamp}_{safe_dom}_{safe_tit}.html"
                    html_path = os.path.join(self.output_dir, base_filename)
                    with open(html_path, "w", encoding="utf-8") as html_file:
                        html_file.write(self.driver.page_source)
                    print(f"已存檔網頁原始碼: {html_path}")
                except Exception as e:
                    error_msg = f"錯誤 - 網域: {domain} / URL: {url} / 錯誤內容: {str(e)}\n"
                    print(f"[HTMLDownloader][錯誤] {error_msg}")
                    with open(error_log_file, "a", encoding="utf-8") as error_file:
                        error_file.write(error_msg)
                finally:
                    time.sleep(SLEEP_DELAY)
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
        global_error_log = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "global_error_log.txt")
        error_msg = f"全域錯誤: {str(e)}\n"
        print(error_msg)
        with open(global_error_log, "a", encoding="utf-8") as error_file:
            error_file.write(error_msg)

if __name__ == "__main__":
    main()
