## Workbook layout note

- `使用說明`: explains sheet order, source attribution, and formulas in user-facing Chinese.
- `低估清單` / `高估清單`: calculated after joining TWSE/TPEx close prices with Cnyes `targetValuation` consensus target prices.
- `全部股票`: one row per listed/OTC stock, combining exchange fields, Cnyes fields, valuation-gap calculations, and fetch status. The lists include `產業別` plus adjacent `產業名稱` for readable industry names.
- Main stock lists end with `鉅亨個股頁`, generated as `https://www.cnyes.com/twstock/{股票代號}` for manual source checks.

## Sample workbook

- `output/tw_valuation_gap_20260703_172748.xlsx` is intentionally tracked as a representative full-market sample so users can preview the workbook tabs, columns, and formatting before running the scanner.
- Day-to-day generated files under `output/` are still ignored by default; only this sample workbook is kept in Git.

# 台股估值落差 Excel 篩選器

這個專案用上市與上櫃每日收盤行情，套疊鉅亨網 Cnyes 的共識目標價，產出 Excel 報表，幫助快速查看哪些股票相對共識目標價可能明顯低估或高估。

## 執行

預設會掃描上市 + 上櫃公司，並嘗試抓取鉅亨網 `targetValuation`：

```powershell
py -3.11 src/tw_target_scan.py
```

或直接執行根目錄的批次檔：

```powershell
run_full_scan.bat
```

`run_full_scan.bat` 會用比較保守的鉅亨抓取速度：每檔間隔 1.0 秒、HTTP 失敗最多重試 2 次、每 25 檔顯示一次進度。全市場約 1980 檔，完整跑完通常需要 30 分鐘以上，實際時間會依網路狀況浮動。bat 預設會同時產完整檔與 `_lite` 輕量檔；日常閱讀建議先開 `_lite`。

若要在終端機或排程中執行且不要停在 `pause`，可以用：

```powershell
run_full_scan.bat --no-pause
```

輸出檔會放在 `output/`，預設完整全市場檔名包含最新收盤日期與產檔時間，避免同一天重跑時覆蓋舊檔，例如：

```text
output/tw_valuation_gap_20260703_160512.xlsx
output/tw_valuation_gap_20260703_160512_lite.xlsx
```

Excel 內的正式資料表會使用中文表頭；程式內部仍保留英文欄位 key 以維持計算與排序穩定。股價、目標價、成交股數(股)、成交金額(元)、成交筆數(筆)、家數等數字欄位會套用千分位格式。

Excel 的「平均目標價」來自鉅亨個股頁內嵌 `targetValuation.feMean`，不是程式自行用最高/最低目標價平均；人工回查可從 `鉅亨個股頁` 欄位直接點擊超連結開啟來源頁並查看「預估」tab，但頁面 UI 不一定逐字顯示「平均目標價」。

鉅亨目標價欄位會逐檔標示「鉅亨資料來源」。目前採用鉅亨個股頁內嵌 `targetValuation`，程式會檢查該頁是否出現 `FactSet` 或 `factSetEstimate` 註記；有才標示 `鉅亨/FactSet共識`，否則標示來源未明確。頁尾 Refinitiv 聲明較偏一般報價與市場資訊來源。此資料是共識統計，不是券商逐筆研究報告，也不列個別機構名稱。

若使用 `--universe watchlist` 或 `--skip-cnyes`，檔名會加上 `_watchlist` 或 `_no_cnyes`，避免測試檔覆蓋完整報表。

如果同名 Excel 正在被開啟，Windows 可能會鎖檔；程式會自動改用帶時間戳的檔名輸出。

小範圍測試可使用 `config/watchlist.csv`：

```powershell
py -3.11 src/tw_target_scan.py --universe watchlist
```

只產生收盤行情與公司基本資料、不抓鉅亨網：

```powershell
py -3.11 src/tw_target_scan.py --skip-cnyes
```

若要先測鉅亨是否穩定，不想一次跑全市場，可以先抓前 20 檔：

```powershell
py -3.11 src/tw_target_scan.py --cnyes-limit 20 --cnyes-delay 0.5 --cnyes-retries 1 --cnyes-progress-every 5
```

## 主要參數

- `--universe all|watchlist`：預設 `all`，可改用自選清單測試。
- `--market all|tse|otc`：預設 `all`，可只跑上市或上櫃。
- `--undervalued-threshold 20`：平均目標價高於收盤價 20% 以上視為低估。
- `--overvalued-threshold -10`：平均目標價低於收盤價 10% 以上視為高估。
- `--stale-days 90`：鉅亨評價日期超過 90 天視為過舊。
- `--min-estimates 3`：預估家數少於 3 家視為低可信度。
- `--excel-output full|lite|both`：預設 `both`，同時輸出完整檔與 `_lite` 輕量檔。
- `--skip-cnyes`：略過鉅亨網抓取。
- `--cnyes-delay 0.5`：每檔股票抓鉅亨網之間的秒數間隔；全市場建議先用預設或更慢。
- `--cnyes-retries 2`：鉅亨單檔 HTTP 失敗時最多重試 2 次。
- `--cnyes-backoff 2.0`：鉅亨重試前等待秒數，會逐次放大。
- `--cnyes-progress-every 50`：每抓 50 檔在終端機顯示一次鉅亨抓取狀態。
- `--cnyes-limit 20`：只測前 20 檔鉅亨資料，其餘股票保留但標示 `skipped_limit`。
- `--cnyes-error-stop-after 30` / `--cnyes-error-stop-rate 0.5`：前 30 檔後若 HTTP error 達 50% 以上，停止後續鉅亨請求並標示 `skipped_error_threshold`。

## `.env` 設定

可複製 `.env.example` 的格式到本機 `.env`。`.env` 不提交 Git，可放門檻與輸出模式。

常用設定：

- `TW_STOCK_UNDERVALUED_THRESHOLD=20.0`：平均目標價高於收盤價 20% 以上才算低估。
- `TW_STOCK_OVERVALUED_THRESHOLD=-10.0`：平均目標價低於收盤價 10% 以上才算高估。
- `TW_STOCK_STALE_DAYS=90`：鉅亨評價日期超過 90 天視為過舊。
- `TW_STOCK_MIN_ESTIMATES=3`：預估家數少於 3 家視為低信心。
- `TW_STOCK_EXCEL_OUTPUT=both`：產完整檔與 `_lite` 輕量檔；也可設 `full` 或 `lite`。

CLI 參數優先於 `.env`。例如命令列指定 `--stale-days 120` 時，會覆蓋 `.env` 的 `TW_STOCK_STALE_DAYS`。

## Excel 工作表

- `使用說明`：中文說明欄位分區、表頭意思，以及估值落差計算如何使用鉅亨目標價。此分頁放在第一個，方便查看說明。
- `欄位說明`：逐欄說明中文表頭、資料來源、公式、單位與括號內的內部 key；完整檔與 `_lite` 檔都會放在 `使用說明` 旁邊。
- `低估清單`：依平均目標價上漲空間由高到低排序；Excel 開啟時預設停在這個分頁。
- `高估清單`：依高估程度排序。
- `全部股票`：全上市 + 上櫃公司，一支股票一列；TWSE/TPEx 資料與鉅亨資料對齊在同一列。
- `過舊低信心`：評價過舊、預估家數太少、或缺少鉅亨目標價的股票。
- `抓取狀態`：各資料源抓取狀態、筆數與資料日期。

低估/高估清單會先排除資料品質不足的股票：缺鉅亨目標價、評價超過 `TW_STOCK_STALE_DAYS`、或預估家數少於 `TW_STOCK_MIN_ESTIMATES`，都不會放進主要低估/高估清單，而會在 Excel 顯示為「缺少目標價」、「評價過舊」或「低信心」。

## 資料源

詳見 `docs/data-sources.md`。

本專案不儲存機密資訊。`output/` 產物預設不提交 Git。
