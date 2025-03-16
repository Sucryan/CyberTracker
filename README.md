
```markdown
# CyberTracker 專案說明

這份 README 提供了 CyberTracker 程式的架構與執行路徑說明，並以 Mermaid 流程圖呈現檔案結構。

---

## 1. CyberTracker 程式本身

這裡展示的是 CyberTracker 程式的主要目錄結構，包含主要執行檔與相關資料夾。

```mermaid
flowchart TB
    A((CyberTracker))
    A --> B[csv_to_xlsx.exe <br/># 執行通報彙整檔轉換的程式]
    A --> C[CyberTracker.exe <br/># UI程式]
    A --> D[merge_csv.exe <br/># 將 all_csv 的檔案合併成 csv_stuff/total.csv]
    A --> E[web_capture.exe <br/># 執行網頁截圖與原始檔下載]
    A --> F((all_csv))
    A --> G((csv_stuff))
    
    %% all_csv 裡面的檔案
    F --> F1[2025-02-08.申報.樂信創意 (1).csv]
    F --> F2[2025-02-08.申報.樂信創意 (2).csv]
    F --> F3[2025-02-08.申報.樂信創意 (3).csv]
    F --> F4[2025-02-08.申報.蝦皮.csv]
    
    %% csv_stuff 裡面的檔案
    G --> G1[total.csv <br/># (合併結果)]
```

---

## 2. CyberTrackerOutput 結構

以下流程圖描述 CyberTracker 執行後輸出的目錄結構，包含 CSV、Line 報告以及不同平台（laptop 與 mobile）下的 HTML 與 PNG 資料夾：

```mermaid
flowchart TB
    A((CyberTrackerOutput))
    A --> B[20250317.申報.樂信+蝦皮(2筆).csv <br/># 彙整檔案 (UTF-8 CSV)]
    A --> C[LineReport.txt <br/># 用於 Line 通報]
    A --> D((laptop))
    A --> E((mobile))
    A --> F((xlsx))
    
    %% laptop 資料夾
    D --> D1((html))
    D --> D2((png))
    
    D1 --> D1a[03170245_ovuts.top_Google.html]
    D1 --> D1b[03170245_xinb.fokd.store_Yahoo奇摩.html]
    D1 --> D1c[error_log.txt <br/># 紀錄錯誤原因及 URL]
    
    D2 --> D2a[03170245_ovuts.top_Google.png]
    D2 --> D2b[03170245_xinb.fokd.store_Yahoo奇摩.png]
    D2 --> D2c[error_log.txt <br/># 同上]
    
    %% mobile 資料夾
    E --> E1((html))
    E --> E2((png))
    
    E1 --> E1a[03170245_ovuts.top_Google.html]
    E1 --> E1b[03170245_xinb.fokd.store_Yahoo奇摩.html]
    E1 --> E1c[error_log.txt]
    
    E2 --> E2a[03170245_ovuts.top_Google.png]
    E2 --> E2b[03170245_xinb.fokd.store_Yahoo奇摩.png]
    E2 --> E2c[error_log.txt]
    
    %% xlsx 資料夾
    F --> F1[通報TWNIC詐騙網址彙整表_(0317).xlsx]
```

---

## 3. 其他相關資源

另外，你也可以參考下方圖片 (由 draw.io 產生) 瞭解 CyberTracker 的執行路徑：

![CyberTracker執行路徑.drawio](https://hackmd.io/_uploads/SylPfscE3Jx.png)

---

## 使用說明

- 將這份 README.md 貼到你的專案中，並確認你使用的環境（例如 GitHub、GitLab 或 HackMD）支援 Mermaid 語法，即可直接預覽流程圖。
- 若你只想呈現 cmd 的 `tree` 結果，建議直接使用程式碼區塊，不會有高亮效果，但可以保留原始格式。

---

希望這份 README 能幫助你更清楚地展示專案結構與執行路徑！
```

---

將上述內容複製到你的 README.md 中，保存後在支援 Mermaid 的平台上就能看到完整的樹狀圖與流程圖。