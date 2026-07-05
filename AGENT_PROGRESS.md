# Agent Progress

## 2026-07-04

- Restored exchange trading columns in the readable artifact by merging close/trade fields back from the no-Cnyes exchange-only workbook; `交易所資料來源` and `交易所資料說明` are also auto-filled when missing before workbook output. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_complete.xlsx`; only 3 stocks remain blank for close/trade fields because the source close data has no row/value for them: 1591, 4804, 1589.
- Suppressed Excel `numberStoredAsText` warnings for `股票代號` and `產業別` columns while keeping them stored as text to preserve stock-code and industry-code semantics. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_no_warnings.xlsx`.
- Removed the extra Excel `交易所資料` and `產業別說明` worksheets from generated workbooks after user feedback. Added `industry_name` / `產業名稱` immediately next to `industry` / `產業別` in the main lists, with code-to-name mapping applied in-row. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_industryname.xlsx` with 1980 rows, 78 undervalued, 35 overvalued, active tab on `低估清單`.
- Added an Excel `產業別說明` sheet with TWSE/TPEx industry-code explanations, updated the guide/dictionary wording for `產業別`, fixed the latest readable workbook by restoring `最新收盤日期` from source status dates, and slimmed `過舊低信心` to 15 key columns to reduce workbook load. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_fixed.xlsx` with 1980 rows, 78 undervalued, 35 overvalued.
- Added detailed crawl-method notes to the Excel `使用說明` sheet: TWSE/TPEx are OpenAPI JSON downloads, Cnyes is HTML page fetching with embedded `targetValuation` JSON parsing via Python `urllib`, and the guide now documents request pacing, retries/backoff, `run_full_scan.bat` conservative settings, `--cnyes-limit`, `--skip-cnyes`, and `skipped_error_threshold`. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_crawlinfo.xlsx`.
- Clarified exchange trading units in Excel output: `成交股數(股)`, `成交金額(元)`, and `成交筆數(筆)`. The guide and data dictionary now state that TWSE `TradeVolume` and TPEx `TradingShares` are shares, not lots, and that TWSE `TradeValue` / TPEx `TransactionAmount` are NTD. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_units.xlsx`.
- Renamed the Excel display header for `close_date` from `收盤日期` to `最新收盤日期`, and updated guide/dictionary wording to match. Refreshed artifact: `output/tw_valuation_gap_20260703_readable_latestclose.xlsx`, verified no `autoFilter`, 1980 all-stock rows, 78 undervalued rows, and 35 overvalued rows.
- Changed Excel display for the `市場` column from internal codes (`tse`/`otc`) to user-facing values (`上市`/`上櫃`) while keeping internal code handling for joins. Refreshed artifact: `output/tw_valuation_gap_20260703_readable.xlsx`, verified no `autoFilter`, 1980 all-stock rows, 78 undervalued rows, and 35 overvalued rows.
- Removed default Excel `autoFilter` output after the user noted the workbook does not need to open in filter mode. A refreshed artifact was written to `output/tw_valuation_gap_20260703_sources_nofilter.xlsx`; verification found no `autoFilter` XML parts and retained 1980 all-stock rows, 78 undervalued rows, and 35 overvalued rows.
- Added a new Excel worksheet `交易所資料` to separate TWSE/TPEx-only data from the joined valuation view. New generated workbooks now open on `低估清單` as sheet 3. Because the prior full workbook was locked by Excel, a refreshed full artifact was written to `output/tw_valuation_gap_20260703_sources.xlsx` with 1980 all-stock rows, 78 undervalued rows, and 35 overvalued rows.
- Updated the Excel `使用說明` guide and `欄位說明` data dictionary to match the localized worksheet names and Chinese headers. Existing full output `output/tw_valuation_gap_20260703.xlsx` was refreshed in place without re-crawling, and the watchlist no-Cnyes smoke test was regenerated successfully.
- Implemented the planned Taiwan valuation-gap Excel workflow in `src/tw_target_scan.py`.
- Main source is now daily close data, not TWSE MIS intraday quotes: TWSE `STOCK_DAY_ALL` for listed stocks and TPEx `tpex_mainboard_daily_close_quotes` for OTC stocks.
- Company universe is joined from TWSE `t187ap03_L` and TPEx `mopsfin_t187ap03_O` so `All_Stocks` is one row per listed/OTC company.
- Cnyes target valuation is fetched from `https://www.cnyes.com/twstock/{stock_id}` embedded `targetValuation`; missing target data keeps the row and sets `cnyes_status`.
- Output is now `.xlsx` only, written under `output/` as `tw_valuation_gap_YYYYMMDD.xlsx`, with sheets `Undervalued`, `Overvalued`, `All_Stocks`, `Stale_or_Low_Confidence`, `Fetch_Status`, and `Data_Dictionary`.
- Added a `Guide` sheet explaining column groups, field names, and the user's question that valuation-gap columns compare exchange close prices against Cnyes consensus target prices.
- If the target `.xlsx` is locked by Excel, the script now writes a timestamp-suffixed filename instead of failing.
- Non-default test modes add filename suffixes (`_watchlist`, `_no_cnyes`) so validation runs do not overwrite the default full-market report.
- README and `docs/data-sources.md` were updated for the new daily-close + Cnyes consensus workflow.
- Added safer Cnyes crawling controls: browser-like headers, default `--cnyes-delay 0.5`, `--cnyes-retries`, exponential backoff, progress logging, `--cnyes-limit`, and a high HTTP-error early stop (`skipped_error_threshold`).
- Added short retry handling for TWSE/TPEx JSON downloads after a TPEx close-data `IncompleteRead` happened during testing.
- Verified `py -3.11 -m py_compile src\tw_target_scan.py`.
- Verified watchlist Cnyes run with 3/3 Cnyes `ok`.
- Verified all-market sample with `--cnyes-limit 20 --cnyes-delay 0.5 --cnyes-retries 1`: output `output/tw_valuation_gap_20260703_cnyes_limit20.xlsx`, 1980 stock rows, Cnyes 20/20 `ok`, remaining rows marked `skipped_limit`.
- Added `run_full_scan.bat` as the main Windows entrypoint. It runs the full scanner with conservative Cnyes settings: `--cnyes-delay 1.0`, `--cnyes-retries 2`, `--cnyes-backoff 2.0`, progress every 25 stocks, and high-error early stop protection.
- README now documents the `.bat` workflow, `--no-pause`, and the small `--cnyes-limit 20` stability test.
- Completed a full all-market run through `cmd /c run_full_scan.bat --no-pause`. Output: `output/tw_valuation_gap_20260703.xlsx`, 1980 rows, 78 undervalued, 35 overvalued. Cnyes fetch result: `ok=1968`, `no_target_valuation=12`, no HTTP errors. Runtime was about 55 minutes.
- Pre-push safety check found existing `origin` still points to `https://github.com/billshiou2/00_project-template.git`; user requested target repo `https://github.com/billshiou2/tw-stock-valuation-gap.git`. Remote must be changed only after explicit confirmation because it currently points to a different project.
- Updated Excel output formatting after user asked for direct Chinese headers and thousands separators. `src/tw_target_scan.py` now keeps internal English keys but writes Chinese headers to `.xlsx`; prices/targets use `#,##0.00`, integer quantities/counts use `#,##0`, and percentage fields remain `0.00%`.
- Rewrote the existing full output workbook `output/tw_valuation_gap_20260703.xlsx` in place without re-crawling data so it now has Chinese headers and thousands number formats.
- Updated Excel header styling after user asked for centered headers and no dotted-looking background fill. Header cells now use bold centered text with no fill; existing full workbook styles were refreshed in place without re-crawling.
- Localized Excel worksheet names after user asked for Chinese sheet tabs. New generated workbooks use `低估清單`, `高估清單`, `全部股票`, `過舊低信心`, `抓取狀態`, `使用說明`, and `欄位說明`; existing full workbook sheet names were updated in place without re-crawling.
- Reordered Excel tabs after user asked to put `使用說明` first while opening by default on `低估清單`. New generated workbooks place `使用說明` first, `低估清單` second, and set `activeTab=1`; existing full workbook was updated in place without re-crawling.

## 2026-06-12

- 使用者想建立台股「目前價格 vs 研究機構/券商目標價」掃描工具，偏好大間機構如元大、統一，並詢問是否有公開資源。
- 已建立第一版 Python 標準函式庫工具 `src/tw_target_scan.py`，用 TWSE 公開 MIS 行情端點抓上市/上櫃即時或最近行情。
- 已新增 `config/watchlist.csv` 作為掃描股票清單，`config/target_prices.csv` 作為可追溯目標價匯入表。
- 已新增 `docs/data-sources.md` 說明：台股行情可用公開交易所端點；券商目標價目前沒有集中免費官方 API，先以 CSV 保存來源、機構、日期與 URL，後續可再接公開新聞或付費資料源。
- README 已加入執行方式：`py -3.11 src/tw_target_scan.py`，輸出在 `output/`；本機 `py -3` 會解析到 Windows Store alias 而失敗。
- 已用 `py -3.11 src\tw_target_scan.py` 實際連 TWSE 公開行情端點驗證，成功產生 `output/tw_target_scan_summary.csv`、`output/tw_target_scan_detail.csv`、`output/tw_target_scan_report.md`。

## 目前目標

- 建立可直接使用的乾淨專案起始結構，並讓後續 agent 能接續目前進度。

## 目前進度

- 已保留 `AGENTS.md` 原本的套件版本安全規則。
- 已追加專案結構、機密管理與 agent 進度記錄規則。
- 已建立 `README.md`、`.gitignore`、`.env.example` 與基礎資料夾。
- 已新增 `.gitkeep` 以保留基礎空資料夾。
- 已新增 `docs/git-workflow.md`，記錄人工建立 GitHub repo 後的 Git 上傳流程與上傳前檢查。
- 已在 `README.md` 加入 Git 上傳流程入口。
- 使用者決定要將此專案上傳到 Git；已完成本機 Git 初始化、remote 設定與 GitHub push。
- 已補強 Git 上傳前確認規則：push 前必須核對目前資料夾、專案說明、branch / upstream、remote URL 或準備設定的 repo URL 與使用者指定目標是否能合理對應；若看起來像不同專案或與指定目標不符，必須停止並再次確認。

## 已決定事項

- 原本的套件版本安全規則不得刪除、修改或弱化。
- 不建立 `docs/DECISIONS.md`，交接重點集中記錄在 `AGENT_PROGRESS.md`。
- 不預設 Node、Python、React 或其他特定框架。
- 不使用「範本」、`template`、「請替換」等字眼描述本專案。
- `.env` 放在專案根目錄但不得提交，`.env.example` 可以提交。
- 新增、刪除或修改 `.env` 的環境變數時，必須同步更新 `.env.example`。
- `.env.example` 是環境變數名稱、格式與用途說明來源。
- `.env` 與 `.env.example` 的非敏感預設值應盡量保持一致；敏感或本機專用值可以不同。
- 任務收尾前必須檢查 `AGENT_PROGRESS.md`、環境變數同步、`README.md` 更新需求與測試執行情況。
- `config/` 放非機密、可提交的設定。
- `docs/` 放文件、需求、規格與說明。
- `output/` 放產生輸出，預設不提交。
- `src/` 放正式原始碼。
- `tmp/` 放暫存檔，預設不提交。
- 使用 `.gitkeep` 保留 `config/`、`docs/`、`src/`、`output/`、`tmp/` 的資料夾結構。
- GitHub repo 建議先由使用者在網頁人工建立，確認名稱、private/public 與權限後，再用 Git 指令上傳。
- GitHub CLI 不列入預設流程，只作為未來需要自動化時的可選方式。
- 執行 `git push` 或設定 / 修改 remote 前，必須確認本機專案名稱 / 路徑、README 或專案說明、branch、upstream 與 remote repo 名稱或準備設定的 repo URL 能合理對應到同一個專案；若看起來像不同專案或與使用者明確指定目標不符，必須先通知使用者並等待再次確認。
- 若既有 remote 看起來指向不同專案，不得自行覆蓋或 push，必須先回報差異。

## 歷程重點

- 討論後決定採用精簡結構，避免一開始加入過多管理文件。
- 討論後決定使用 `AGENT_PROGRESS.md` 記錄目前進度、重點討論、待確認事項與下一步。
- 討論後決定資料夾說明寫在 `AGENTS.md` 與 `README.md`，不在每個資料夾各放說明檔。
- 討論後決定修改 `.env` 時必須同步更新 `.env.example`。
- 參考 `poly-rewards4` 規則後，決定補強 `.env.example` 說明來源、非敏感預設值一致與任務收尾檢查規則。
- 討論後決定加入 `.gitkeep`，避免空資料夾在 Git 中消失。
- 討論後決定補上 Git 上傳流程，採用人工建立 GitHub repo 加本機 Git 指令 push 的保守流程。
- 討論後決定開始將目前專案上傳到 Git；已完成本機 Git 初始化、安全檢查與第一個 commit，並設定 remote 為 `https://github.com/billshiou2/00_project-template.git`。
- 嘗試 push 時，網路沙盒先阻擋連線；使用者批准後，已成功推送到 GitHub。
- 因 Git 偵測 repo ownership 與目前 Windows 使用者不同，已依 Git 提示將 `C:/# code/# project-template3` 加入全域 `safe.directory`。
- 使用者提醒曾發生 Git 上傳錯 repo 的風險，因此已將 push 前專案名稱 / remote / branch 核對流程寫入 `AGENTS.md`、`docs/git-workflow.md` 與 `README.md`；後續又確認不要求名稱逐字相同，只要求能合理對應。
- 使用者要求 push 本次 Git 上傳前確認規則更新；push 前已核對目前路徑、remote、branch 與變更檔案能合理對應到本專案。

## 待確認事項

- 無。

## 下一步

- 依實際專案需求新增原始碼、設定或文件。
- 若新增套件，必須遵守 `AGENTS.md` 的套件版本安全規則。
- 新專案第一次上傳 GitHub 時，參考 `docs/git-workflow.md`。
- 目前 `main` 已設定追蹤 `origin/main`，後續一般更新可使用 `git push`。
