# CyberTracker 專案說明

這份 README 提供了 CyberTracker 程式的架構與執行路徑說明，並以 Mermaid 流程圖呈現檔案結構。

## 1. CyberTracker 程式本身

這裡展示的是 CyberTracker 程式的主要目錄結構，包含主要執行檔與相關資料夾。

```mermaid
flowchart TB
    A((CyberTracker))
    A --> B["csv_to_xlsx.exe\n 執行通報彙整檔轉換的程式"]
    A --> C["CyberTracker.exe\n UI 程式"]
    A --> D["merge_csv.exe\n 將 all_csv 的檔案合併成 csv_stuff/total.csv"]
    A --> E["web_capture.exe\n 執行網頁截圖與原始檔下載"]
    A --> F((all_csv))
    A --> G((csv_stuff))

    F --> F1["2025-02-08.申報.樂信創意 (1).csv"]
    F --> F2["2025-02-08.申報.樂信創意 (2).csv"]
    F --> F3["2025-02-08.申報.樂信創意 (3).csv"]
    F --> F4["2025-02-08.申報.蝦皮.csv"]

    G --> G1["total.csv\n (將上述檔案合併)"]
```

## 2. CyberTrackerOutput 結構

以下流程圖描述 CyberTracker 執行後輸出的目錄結構，包含 CSV、Line 報告以及不同平台（laptop 與 mobile）下的 HTML 與 PNG 資料夾：

```mermaid
flowchart TB
    A((CyberTrackerOutput))
    A --> B["20250317.申報.樂信+蝦皮(2筆).csv\n 彙整檔案 (UTF-8 CSV)"]
    A --> C["LineReport.txt\n 用於 Line 通報"]
    A --> D((laptop))
    A --> E((mobile))
    A --> F((xlsx))

    D --> D1((html))
    D --> D2((png))
    D1 --> D1a["03170245_ovuts.top_Google.html"]
    D1 --> D1b["03170245_xinb.fokd.store_Yahoo奇摩.html"]
    D1 --> D1c["error_log.txt\n 紀錄錯誤原因及 URL"]
    D2 --> D2a["03170245_ovuts.top_Google.png"]
    D2 --> D2b["03170245_xinb.fokd.store_Yahoo奇摩.png"]
    D2 --> D2c["error_log.txt\n 同上"]

    E --> E1((html))
    E --> E2((png))
    E1 --> E1a["03170245_ovuts.top_Google.html"]
    E1 --> E1b["03170245_xinb.fokd.store_Yahoo奇摩.html"]
    E1 --> E1c["error_log.txt"]
    E2 --> E2a["03170245_ovuts.top_Google.png"]
    E2 --> E2b["03170245_xinb.fokd.store_Yahoo奇摩.png"]
    E2 --> E2c["error_log.txt"]

    F --> F1["通報TWNIC詐騙網址彙整表_(0317).xlsx"]
```

---

## 3. 其他相關資源

另外，你也可以參考下方圖片 (由 draw.io 產生) 瞭解 CyberTracker 的執行路徑：

![CyberTracker執行路徑.drawio](https://hackmd.io/_uploads/SylPfscE3Jx.png)

---
