import argparse
from web_capture import ScreenshotTaker, HTMLDownloader
import sys

def main():
    parser = argparse.ArgumentParser(description="Web Capture Tool")
    parser.add_argument("mode", choices=["screenshot", "html"], help="選擇要執行的功能: screenshot 或 html")
    parser.add_argument("--csv", type=str, help="CSV 檔案路徑", default=None)
    parser.add_argument("--output", type=str, help="輸出資料夾", default=None)

    # 預設 headless = True，但允許使用者用 --no-headless 來關閉
    parser.add_argument("--no-headless", action="store_false", dest="headless", help="停用 headless 模式，顯示瀏覽器視窗")
    
    args = parser.parse_args()

    if args.mode == "screenshot":
        taker = ScreenshotTaker(csv_file=args.csv, output_dir=args.output, headless=args.headless)
        taker.run()
        taker.close()
    elif args.mode == "html":
        downloader = HTMLDownloader(csv_file=args.csv, output_dir=args.output, headless=args.headless)
        downloader.run()
        downloader.close()

if __name__ == "__main__":
    main()
