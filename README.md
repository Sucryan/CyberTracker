##  流程圖/樹狀圖

### CyberTracker程式本身(範例)
```bash=
CyberTracker
|   csv_to_xlsx.exe  # 執行通報彙整檔轉換的程式
|   CyberTracker.exe # UI程式
|   merge_csv.exe    # 將all_csv的程式碼彙整成csv_stuff中的total.csv
|   web_capture.exe  # 執行網頁截圖與網頁原始檔下載
|
+---all_csv # (透過匯入按鈕輸入的檔案會儲存在這裡)
|       2025-02-08.申報.樂信創意 (1).csv
|       2025-02-08.申報.樂信創意 (2).csv
|       2025-02-08.申報.樂信創意 (3).csv
|       2025-02-08.申報.蝦皮.csv
|
\---csv_stuff 
        total.csv #(將上述檔案合併的檔案)
```

### CyberTrackerOutput
```bash=
CyberTrackerOutput
|   20250317.申報.蝦皮+電商(2筆).csv    # 彙整檔案(轉換成UTF=8的csv檔)
|   LineReport.txt                   # 可以拿來傳送給蘇昱辰的line通報檔
+---laptop # 儲存電腦版內容
|   +---html
|   |       03170210_ovuts.top_Google.html 
|   |       03170210_xinb.fokd.store_Yahoo奇摩.html
|   |       error_log.txt            # 會儲存error的原因及其url，方便審查員進行手動處理。
|   |
|   \---png
|           03170210_ovuts.top_Google.png
|           03170210_xinb.fokd.store_Yahoo奇摩.png
|           error_log.txt
|
\---mobile # 儲存手機版內容
    +---html
    |       03170210_ovuts.top_Google.html
    |       03170210_xinb.fokd.store_Yahoo奇摩.html
    |       error_log.txt
    |
    \---png
            03170210_ovuts.top_Google.png
            03170210_xinb.fokd.store_Yahoo奇摩.png
            error_log.txt

CyberTrackerOutput
|   20250317.申報.樂信+蝦皮(2筆).csv  # 彙整檔案(轉換成UTF=8的csv檔)
|   LineReport.txt                 # 可以拿來傳送給蘇昱辰的line通報檔
|
+---laptop # 儲存電腦版內容
|   +---html
|   |       03170245_ovuts.top_Google.html
|   |       03170245_xinb.fokd.store_Yahoo奇摩.html
|   |       error_log.txt            # 會儲存error的原因及其url，方便審查員進行手動處理。
|   |
|   \---png
|           03170245_ovuts.top_Google.png
|           03170245_xinb.fokd.store_Yahoo奇摩.png
|           error_log.txt            # 會儲存error的原因及其url，方便審查員進行手動處理。
|
+---mobile # 儲存手機版內容
|   +---html
|   |       03170245_ovuts.top_Google.html
|   |       03170245_xinb.fokd.store_Yahoo奇摩.html
|   |       error_log.txt            # 會儲存error的原因及其url，方便審查員進行手動處理。
|   |
|   \---png
|           03170245_ovuts.top_Google.png
|           03170245_xinb.fokd.store_Yahoo奇摩.png
|          error_log.txt            # 會儲存error的原因及其url，方便審查員進行手動處理。
|
\---xlsx # 儲存需要轉檔的內容
        通報TWNIC詐騙網址彙整表_(0317).xlsx
```


![CyberTracker執行路徑.drawio](https://hackmd.io/_uploads/SylPfscE3Jx.png)