# Data Sources

## Daily Close Prices

This project uses daily close prices as the primary price source because the goal is valuation-gap screening, not intraday trading.

### Listed Stocks

TWSE daily close endpoint:

`https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL`

Useful fields:

- `Date`: ROC date, for example `1150703`.
- `Code`: stock/security code.
- `Name`: display name.
- `TradeVolume`: traded shares.
- `TradeValue`: traded value.
- `OpeningPrice`, `HighestPrice`, `LowestPrice`, `ClosingPrice`.
- `Change`: daily price change.
- `Transaction`: transaction count.

### OTC Stocks

TPEx daily close endpoint:

`https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes`

Useful fields:

- `Date`: ROC date.
- `SecuritiesCompanyCode`: stock/security code.
- `CompanyName`: display name.
- `Close`, `Change`, `Open`, `High`, `Low`, `Average`.
- `TradingShares`, `TransactionAmount`, `TransactionNumber`.
- `LatestBidPrice`, `LatesAskPrice`.

## Company Universe

Daily close endpoints include non-common-stock products such as ETFs. The scanner builds the company universe from company profile endpoints, then joins daily close rows by market and stock code.

Listed company profiles:

`https://openapi.twse.com.tw/v1/opendata/t187ap03_L`

OTC company profiles:

`https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O`

## Exchange Fundamentals

These fields are appended to the right side of the joined stock rows. They are public exchange datasets and are not Cnyes target-price data.

### PER / PBR / Dividend Yield

Listed stocks:

`https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL`

Useful fields:

- `Date`: data date shown in Excel headers.
- `Code`: stock/security code.
- `PEratio`: PER.
- `DividendYield`: dividend yield. The source value is a percent number; the workbook stores it as an Excel percent. If the official field is blank, the adjacent dividend-yield note says so explicitly instead of labeling the whole fundamentals fetch as failed.
- `PBratio`: PBR.

OTC stocks:

`https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis`

Useful fields:

- `Date`: data date shown in Excel headers.
- `SecuritiesCompanyCode`: stock/security code.
- `PriceEarningRatio`: PER.
- `YieldRatio`: dividend yield. The source value is a percent number; the workbook stores it as an Excel percent. The adjacent note distinguishes an official blank field from a missing valuation row or parse problem.
- `PriceBookRatio`: PBR.

### Monthly Revenue

Monthly revenue is a separate official dataset from the PER/PBR/dividend-yield valuation endpoints. A stock can therefore have valuation data while its monthly-revenue row is absent or individual growth fields are blank. The workbook exposes this distinction in the adjacent monthly-revenue note column.

Listed stocks:

`https://openapi.twse.com.tw/v1/opendata/t187ap05_L`

OTC stocks:

`https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap05_O`

Useful fields:

- `資料年月`: revenue period, converted from ROC year/month to `YYYY-MM`.
- `公司代號`: stock/security code.
- `營業收入-當月營收`: monthly revenue, shown as thousand NTD.
- `營業收入-上月比較增減(%)`: month-over-month revenue change.
- `營業收入-去年同月增減(%)`: year-over-year revenue change.
- `累計營業收入-前期比較增減(%)`: cumulative revenue year-over-year change.

### EPS / Income / Margins

Listed EPS and income:

`https://openapi.twse.com.tw/v1/opendata/t187ap14_L`

Listed margin ratios:

`https://openapi.twse.com.tw/v1/opendata/t187ap17_L`

OTC EPS and income:

`https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap14_O`

Useful fields:

- `年度` / `Year` plus `季別`: financial period, converted to `YYYYQn`.
- `基本每股盈餘(元)` / `基本每股盈餘`: EPS.
- `營業收入`, `營業利益`, `稅後淨利`: financial statement values, shown as thousand NTD.
- TWSE `t187ap17_L` provides listed-stock gross margin, operating margin, and net margin directly.
- OTC operating margin and net margin are calculated from TPEx income fields when revenue is available.

Current limitations:

- OTC gross margin is left blank because a stable TPEx OpenAPI margin-ratio field has not been confirmed.
- Do not substitute OTC operating margin into gross margin; these are different profitability measures.
- ROE is not included yet because a stable public equity-data source has not been confirmed.

## Cnyes Consensus Target Valuation

Cnyes human UI URL pattern:

`https://www.cnyes.com/twstock/2330`

Use the page's `預估` tab for manual inspection. The scanner reads embedded `targetValuation` data from the same stock page. The `/forecast` path is not used.

Observed useful fields:

- `rateDate`: valuation date.
- `feHigh`: highest consensus target.
- `feLow`: lowest consensus target.
- `feMean`: average consensus target.
- `feMedian`: median consensus target.
- `feUp`, `feDown`: estimate direction counts when available.
- `feStdDev`: standard deviation.
- `numEst`: estimate count.
- `currency`: currency.
- `last`: Cnyes page's paired latest price.

Cnyes target valuation is consensus data, not broker-by-broker target prices.

The generated workbook labels `鉅亨資料來源` per stock. For target valuation, the scanner reads Cnyes embedded `targetValuation` and then checks that stock page for `FactSet` or `factSetEstimate`; only those rows are labeled `鉅亨/FactSet共識`. Rows without an explicit FactSet marker are labeled as Cnyes consensus with unclear source attribution. Cnyes page footers also mention Refinitiv for market quote information, but that is broader quote/market-data attribution rather than broker-by-broker target-price detail.

## Optional Intraday / Recent Quote Fallback

TWSE MIS can query listed and OTC recent quote data:

`https://mis.twse.com.tw/stock/api/getStockInfo.jsp`

This endpoint is no longer the default because the report is based on daily closing prices.
### 基本面欄位的來源、備註與整體狀態

- 「資料來源」表示該組欄位使用的官方端點；即使該股票沒有資料列，來源欄仍會顯示程式查詢的端點。
- 「資料備註」說明官方端點沒有該股票，或官方資料列中有哪些欄位缺值。
- 「基本面整體狀態」只彙總估值、月營收、財報三組資料列是否齊全；個別欄位缺值仍以各組備註為準。
