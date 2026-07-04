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

## Optional Intraday / Recent Quote Fallback

TWSE MIS can query listed and OTC recent quote data:

`https://mis.twse.com.tw/stock/api/getStockInfo.jsp`

This endpoint is no longer the default because the report is based on daily closing prices.
