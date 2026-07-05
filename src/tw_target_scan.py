#!/usr/bin/env python3
"""Build a Taiwan stock valuation-gap Excel report.

The report joins TWSE/TPEx daily close data, company profile data, and Cnyes
consensus target-valuation data. It intentionally uses only Python's standard
library so the project has no package supply-chain surface.
"""

from __future__ import annotations

import argparse
import csv
import json
import posixpath
import re
import sys
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
from urllib.request import Request, urlopen
from xml.sax.saxutils import escape


TWSE_CLOSE_URL = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
TWSE_PROFILE_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
TPEX_CLOSE_URL = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"
TPEX_PROFILE_URL = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"
CNYES_STOCK_URL = "https://www.cnyes.com/twstock/{stock_id}"
CNYES_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
)
MARKETS = ("tse", "otc")


@dataclass(frozen=True)
class StockInfo:
    market: str
    stock_id: str
    name: str = ""
    company_name: str = ""
    industry: str = ""
    listing_date: str = ""


@dataclass
class SourceStatus:
    source: str
    url: str
    status: str
    rows: int = 0
    data_date: str = ""
    message: str = ""


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        return [{k: (v or "").strip() for k, v in row.items()} for row in csv.DictReader(fh)]


def request_text(url: str, timeout: int = 30, referer: str = "") -> str:
    headers = {
        "User-Agent": CNYES_USER_AGENT,
        "Accept": "text/html,application/json,*/*",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    }
    if referer:
        headers["Referer"] = referer
    req = Request(
        url,
        headers=headers,
    )
    with urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
        charset = resp.headers.get_content_charset() or "utf-8"
    return raw.decode(charset, errors="replace").lstrip("\ufeff")


def fetch_json(url: str, source: str, statuses: list[SourceStatus]) -> Any:
    attempts = 3
    last_error = ""
    text = ""
    for attempt in range(1, attempts + 1):
        try:
            text = request_text(url)
            break
        except Exception as exc:
            last_error = str(exc)
            if attempt < attempts:
                time.sleep(1.5 * attempt)
    if not text:
        statuses.append(SourceStatus(source=source, url=url, status="error", message=last_error))
        raise RuntimeError(last_error)
    try:
        data = json.loads(text)
    except Exception as exc:
        statuses.append(SourceStatus(source=source, url=url, status="error", message=str(exc)))
        raise
    rows = len(data) if isinstance(data, list) else 1
    data_date = ""
    if isinstance(data, list) and data:
        first = data[0]
        if isinstance(first, dict):
            data_date = normalize_date(str(first.get("Date") or first.get("出表日期") or ""))
    statuses.append(SourceStatus(source=source, url=url, status="ok", rows=rows, data_date=data_date))
    return data


def normalize_date(raw: str) -> str:
    raw = (raw or "").strip()
    if not raw:
        return ""
    if re.fullmatch(r"\d{7}", raw):
        year = int(raw[:3]) + 1911
        return f"{year:04d}-{raw[3:5]}-{raw[5:7]}"
    if re.fullmatch(r"\d{8}", raw):
        return f"{raw[:4]}-{raw[4:6]}-{raw[6:8]}"
    match = re.fullmatch(r"(\d{3})/(\d{2})/(\d{2})", raw)
    if match:
        year = int(match.group(1)) + 1911
        return f"{year:04d}-{match.group(2)}-{match.group(3)}"
    match = re.fullmatch(r"(\d{4})/(\d{2})/(\d{2})", raw)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    return raw


def parse_date(value: str) -> date | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return None


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    text = str(value).replace(",", "").replace("+", "").strip()
    if not text or text in {"-", "--", "除權息"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def clean_code(value: Any) -> str:
    return str(value or "").strip()


def load_watchlist(path: Path) -> list[StockInfo]:
    stocks: list[StockInfo] = []
    for row in read_csv(path):
        stock_id = clean_code(row.get("stock_id"))
        if not stock_id:
            continue
        market = clean_code(row.get("market")).lower()
        if market and market not in MARKETS:
            raise ValueError(f"Unsupported market for {stock_id}: {market}. Use tse, otc, or blank.")
        stocks.append(StockInfo(market=market, stock_id=stock_id, name=row.get("name", "")))
    return stocks


def fetch_twse_profiles(statuses: list[SourceStatus]) -> dict[tuple[str, str], StockInfo]:
    rows = fetch_json(TWSE_PROFILE_URL, "twse_profiles", statuses)
    profiles: dict[tuple[str, str], StockInfo] = {}
    for row in rows:
        code = clean_code(row.get("公司代號"))
        if not code:
            continue
        profiles[("tse", code)] = StockInfo(
            market="tse",
            stock_id=code,
            name=clean_code(row.get("公司簡稱")),
            company_name=clean_code(row.get("公司名稱")),
            industry=clean_code(row.get("產業別")),
            listing_date=normalize_date(clean_code(row.get("上市日期"))),
        )
    return profiles


def fetch_tpex_profiles(statuses: list[SourceStatus]) -> dict[tuple[str, str], StockInfo]:
    rows = fetch_json(TPEX_PROFILE_URL, "tpex_profiles", statuses)
    profiles: dict[tuple[str, str], StockInfo] = {}
    for row in rows:
        code = clean_code(row.get("SecuritiesCompanyCode"))
        if not code:
            continue
        profiles[("otc", code)] = StockInfo(
            market="otc",
            stock_id=code,
            name=clean_code(row.get("CompanyAbbreviation")),
            company_name=clean_code(row.get("CompanyName")),
            industry=clean_code(row.get("SecuritiesIndustryCode")),
            listing_date=normalize_date(clean_code(row.get("DateOfListing"))),
        )
    return profiles


def fetch_twse_closes(statuses: list[SourceStatus]) -> dict[tuple[str, str], dict[str, Any]]:
    rows = fetch_json(TWSE_CLOSE_URL, "twse_daily_close", statuses)
    closes: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        code = clean_code(row.get("Code"))
        if not code:
            continue
        closes[("tse", code)] = {
            "close_date": normalize_date(clean_code(row.get("Date"))),
            "close": to_float(row.get("ClosingPrice")),
            "change": to_float(row.get("Change")),
            "open": to_float(row.get("OpeningPrice")),
            "high": to_float(row.get("HighestPrice")),
            "low": to_float(row.get("LowestPrice")),
            "trade_volume": to_float(row.get("TradeVolume")),
            "trade_value": to_float(row.get("TradeValue")),
            "transactions": to_float(row.get("Transaction")),
        }
    return closes


def fetch_tpex_closes(statuses: list[SourceStatus]) -> dict[tuple[str, str], dict[str, Any]]:
    rows = fetch_json(TPEX_CLOSE_URL, "tpex_daily_close", statuses)
    closes: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        code = clean_code(row.get("SecuritiesCompanyCode"))
        if not code:
            continue
        closes[("otc", code)] = {
            "close_date": normalize_date(clean_code(row.get("Date"))),
            "close": to_float(row.get("Close")),
            "change": to_float(row.get("Change")),
            "open": to_float(row.get("Open")),
            "high": to_float(row.get("High")),
            "low": to_float(row.get("Low")),
            "trade_volume": to_float(row.get("TradingShares")),
            "trade_value": to_float(row.get("TransactionAmount")),
            "transactions": to_float(row.get("TransactionNumber")),
        }
    return closes


def select_universe(
    profiles: dict[tuple[str, str], StockInfo],
    watchlist: list[StockInfo],
    universe: str,
    market: str,
) -> list[StockInfo]:
    allowed_markets = set(MARKETS if market == "all" else [market])
    if universe == "all":
        stocks = [stock for key, stock in profiles.items() if key[0] in allowed_markets]
        return sorted(stocks, key=lambda stock: (stock.market, stock.stock_id))

    selected: list[StockInfo] = []
    seen: set[tuple[str, str]] = set()
    for watch in watchlist:
        markets = [watch.market] if watch.market else list(allowed_markets)
        for candidate_market in markets:
            if candidate_market not in allowed_markets:
                continue
            key = (candidate_market, watch.stock_id)
            base = profiles.get(key, watch)
            stock = StockInfo(
                market=candidate_market,
                stock_id=watch.stock_id,
                name=base.name or watch.name,
                company_name=base.company_name,
                industry=base.industry,
                listing_date=base.listing_date,
            )
            if key not in seen:
                selected.append(stock)
                seen.add(key)
    return selected


def extract_cnyes_target(stock_id: str, timeout: int) -> dict[str, Any]:
    url = CNYES_STOCK_URL.format(stock_id=stock_id)
    text = request_text(url, timeout=timeout, referer="https://www.cnyes.com/")
    match = re.search(r'"targetValuation":(\{.*?\}|null),"factSetEstimate"', text)
    if not match:
        return {"cnyes_status": "no_target_valuation", "cnyes_url": url}
    if match.group(1) == "null":
        return {"cnyes_status": "no_target_valuation", "cnyes_url": url}
    try:
        data = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        return {"cnyes_status": "parse_error", "cnyes_url": url, "cnyes_error": str(exc)}
    return {
        "cnyes_url": url,
        "cnyes_status": "ok",
        "target_date": clean_code(data.get("rateDate")),
        "target_high": to_float(data.get("feHigh")),
        "target_low": to_float(data.get("feLow")),
        "target_mean": to_float(data.get("feMean")),
        "target_median": to_float(data.get("feMedian")),
        "fe_up": to_float(data.get("feUp")),
        "fe_down": to_float(data.get("feDown")),
        "fe_stddev": to_float(data.get("feStdDev")),
        "num_est": to_float(data.get("numEst")),
        "currency": clean_code(data.get("currency")),
        "cnyes_last": to_float(data.get("last")),
    }


def extract_cnyes_target_with_retries(stock_id: str, timeout: int, retries: int, backoff: float) -> dict[str, Any]:
    attempts = max(1, retries + 1)
    last_error = ""
    for attempt in range(1, attempts + 1):
        try:
            result = extract_cnyes_target(stock_id, timeout)
            result["cnyes_attempts"] = attempt
            return result
        except Exception as exc:
            last_error = str(exc)
            if attempt < attempts and backoff > 0:
                time.sleep(backoff * (2 ** (attempt - 1)))
    return {
        "cnyes_url": CNYES_STOCK_URL.format(stock_id=stock_id),
        "cnyes_status": "http_error",
        "cnyes_error": last_error,
        "cnyes_attempts": attempts,
    }


def fetch_cnyes_targets(
    stocks: list[StockInfo],
    skip: bool,
    delay: float,
    timeout: int,
    retries: int,
    backoff: float,
    progress_every: int,
    limit: int,
    error_stop_after: int,
    error_stop_rate: float,
    statuses: list[SourceStatus],
) -> dict[str, dict[str, Any]]:
    if skip:
        statuses.append(SourceStatus(source="cnyes_target_valuation", url=CNYES_STOCK_URL, status="skipped"))
        return {}

    targets: dict[str, dict[str, Any]] = {}
    counts: dict[str, int] = {}
    stopped_message = ""
    max_fetch = len(stocks) if limit <= 0 else min(limit, len(stocks))
    fetch_stocks = stocks[:max_fetch]
    skipped_stocks = stocks[max_fetch:]

    for index, stock in enumerate(fetch_stocks, start=1):
        result = extract_cnyes_target_with_retries(stock.stock_id, timeout, retries, backoff)
        targets[stock.stock_id] = result
        status = str(result.get("cnyes_status") or "unknown")
        counts[status] = counts.get(status, 0) + 1

        if progress_every > 0 and (index == 1 or index % progress_every == 0 or index == len(fetch_stocks)):
            message = ", ".join(f"{key}={value}" for key, value in sorted(counts.items()))
            print(f"Cnyes progress {index}/{len(fetch_stocks)}: {message}", file=sys.stderr)

        processed = sum(counts.values())
        http_errors = counts.get("http_error", 0)
        if error_stop_after > 0 and processed >= error_stop_after and processed > 0:
            http_error_rate = http_errors / processed
            if http_error_rate >= error_stop_rate:
                stopped_message = (
                    f"stopped early after {processed} Cnyes requests because "
                    f"http_error_rate={http_error_rate:.0%} >= {error_stop_rate:.0%}"
                )
                counts["skipped_error_threshold"] = counts.get("skipped_error_threshold", 0) + len(fetch_stocks) - index
                for remaining in fetch_stocks[index:]:
                    targets[remaining.stock_id] = {
                        "cnyes_url": CNYES_STOCK_URL.format(stock_id=remaining.stock_id),
                        "cnyes_status": "skipped_error_threshold",
                        "cnyes_error": stopped_message,
                    }
                break

        if index < len(fetch_stocks) and delay > 0:
            time.sleep(delay)

    if skipped_stocks:
        counts["skipped_limit"] = counts.get("skipped_limit", 0) + len(skipped_stocks)
        for stock in skipped_stocks:
            targets[stock.stock_id] = {
                "cnyes_url": CNYES_STOCK_URL.format(stock_id=stock.stock_id),
                "cnyes_status": "skipped_limit",
                "cnyes_error": f"skipped by --cnyes-limit {limit}",
            }

    message = ", ".join(f"{key}={value}" for key, value in sorted(counts.items()))
    if stopped_message:
        message = f"{message}; {stopped_message}"
    statuses.append(
        SourceStatus(
            source="cnyes_target_valuation",
            url=CNYES_STOCK_URL,
            status="partial" if stopped_message or skipped_stocks else "ok",
            rows=len(targets),
            message=message,
        )
    )
    return targets


def pct(target: float | None, close: float | None) -> float | None:
    if target is None or close is None or close == 0:
        return None
    return target / close - 1.0


def days_between(later: str, earlier: str) -> int | None:
    later_date = parse_date(later)
    earlier_date = parse_date(earlier)
    if not later_date or not earlier_date:
        return None
    return (later_date - earlier_date).days


def classify(row: dict[str, Any], undervalued: float, overvalued: float, stale_days: int, min_estimates: int) -> None:
    status = row.get("cnyes_status")
    upside = row.get("upside_to_mean_pct")
    age = row.get("valuation_age_days")
    num_est = row.get("num_est")

    if status != "ok" or upside is None:
        row["valuation_signal"] = "missing_target"
        row["confidence_note"] = "missing_cnyes_target"
        return
    if isinstance(age, int) and age > stale_days:
        row["valuation_signal"] = "stale"
        row["confidence_note"] = f"stale_{age}_days"
        return
    if num_est is not None and num_est < min_estimates:
        row["valuation_signal"] = "low_confidence"
        row["confidence_note"] = f"only_{int(num_est)}_estimates"
        return
    if upside >= undervalued:
        row["valuation_signal"] = "undervalued"
    elif upside <= overvalued:
        row["valuation_signal"] = "overvalued"
    else:
        row["valuation_signal"] = "neutral"
    row["confidence_note"] = f"fresh_{int(num_est or 0)}_estimates"


def build_rows(
    stocks: list[StockInfo],
    closes: dict[tuple[str, str], dict[str, Any]],
    targets: dict[str, dict[str, Any]],
    args: argparse.Namespace,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for stock in stocks:
        close_row = closes.get((stock.market, stock.stock_id), {})
        target_row = targets.get(stock.stock_id, {})
        cnyes_status = target_row.get("cnyes_status", "skipped" if args.skip_cnyes else "missing")
        row: dict[str, Any] = {
            "market": stock.market,
            "stock_id": stock.stock_id,
            "name": stock.name,
            "company_name": stock.company_name,
            "industry": stock.industry,
            "listing_date": stock.listing_date,
            "close_date": close_row.get("close_date", ""),
            "close": close_row.get("close"),
            "change": close_row.get("change"),
            "open": close_row.get("open"),
            "high": close_row.get("high"),
            "low": close_row.get("low"),
            "trade_volume": close_row.get("trade_volume"),
            "trade_value": close_row.get("trade_value"),
            "transactions": close_row.get("transactions"),
            "exchange_source": exchange_source(stock.market),
            "exchange_note": exchange_note(stock.market),
            "upside_to_mean_pct": None,
            "upside_to_median_pct": None,
            "upside_to_high_pct": None,
            "downside_to_low_pct": None,
            "valuation_age_days": None,
            "valuation_signal": "",
            "confidence_note": "",
            "cnyes_url": target_row.get("cnyes_url", CNYES_STOCK_URL.format(stock_id=stock.stock_id)),
            "target_date": target_row.get("target_date", ""),
            "target_high": target_row.get("target_high"),
            "target_low": target_row.get("target_low"),
            "target_mean": target_row.get("target_mean"),
            "target_median": target_row.get("target_median"),
            "num_est": target_row.get("num_est"),
            "fe_up": target_row.get("fe_up"),
            "fe_down": target_row.get("fe_down"),
            "fe_stddev": target_row.get("fe_stddev"),
            "currency": target_row.get("currency", ""),
            "cnyes_last": target_row.get("cnyes_last"),
            "cnyes_status": cnyes_status,
            "cnyes_error": target_row.get("cnyes_error", ""),
            "cnyes_attempts": target_row.get("cnyes_attempts", ""),
        }
        row["upside_to_mean_pct"] = pct(row["target_mean"], row["close"])
        row["upside_to_median_pct"] = pct(row["target_median"], row["close"])
        row["upside_to_high_pct"] = pct(row["target_high"], row["close"])
        row["downside_to_low_pct"] = pct(row["target_low"], row["close"])
        row["valuation_age_days"] = days_between(str(row["close_date"]), str(row["target_date"]))
        classify(
            row,
            args.undervalued_threshold / 100.0,
            args.overvalued_threshold / 100.0,
            args.stale_days,
            args.min_estimates,
        )
        rows.append(row)
    return rows


def column_order() -> list[str]:
    return [
        "market",
        "stock_id",
        "name",
        "company_name",
        "industry",
        "listing_date",
        "close_date",
        "close",
        "change",
        "open",
        "high",
        "low",
        "trade_volume",
        "trade_value",
        "transactions",
        "exchange_source",
        "exchange_note",
        "upside_to_mean_pct",
        "upside_to_median_pct",
        "upside_to_high_pct",
        "downside_to_low_pct",
        "valuation_age_days",
        "valuation_signal",
        "confidence_note",
        "cnyes_url",
        "target_date",
        "target_high",
        "target_low",
        "target_mean",
        "target_median",
        "num_est",
        "fe_up",
        "fe_down",
        "fe_stddev",
        "currency",
        "cnyes_last",
        "cnyes_status",
        "cnyes_error",
        "cnyes_attempts",
    ]


def exchange_column_order() -> list[str]:
    return [
        "market",
        "stock_id",
        "name",
        "company_name",
        "industry",
        "listing_date",
        "close_date",
        "close",
        "change",
        "open",
        "high",
        "low",
        "trade_volume",
        "trade_value",
        "transactions",
        "exchange_source",
        "exchange_note",
    ]


def market_name(market: str) -> str:
    if market == "tse":
        return "上市"
    if market == "otc":
        return "上櫃"
    return market


def exchange_source(market: str) -> str:
    if market == "tse":
        return "TWSE: STOCK_DAY_ALL + t187ap03_L"
    if market == "otc":
        return "TPEx: tpex_mainboard_daily_close_quotes + mopsfin_t187ap03_O"
    return ""


def exchange_note(market: str) -> str:
    if market == "tse":
        return "上市股票；收盤行情來自 TWSE STOCK_DAY_ALL，基本資料來自 TWSE t187ap03_L。"
    if market == "otc":
        return "上櫃股票；收盤行情來自 TPEx tpex_mainboard_daily_close_quotes，基本資料來自 TPEx mopsfin_t187ap03_O。"
    return ""


def exchange_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for row in rows:
        market = str(row.get("market") or "")
        exchange_row = {column: row.get(column, "") for column in exchange_column_order()}
        exchange_row["exchange_source"] = exchange_source(market)
        exchange_row["exchange_note"] = exchange_note(market)
        result.append(exchange_row)
    return result


def column_label(column: str) -> str:
    labels = {
        "market": "市場",
        "stock_id": "股票代號",
        "name": "股票名稱",
        "company_name": "公司全名",
        "industry": "產業別",
        "listing_date": "上市櫃日期",
        "close_date": "最新收盤日期",
        "close": "收盤價",
        "change": "漲跌",
        "open": "開盤價",
        "high": "最高價",
        "low": "最低價",
        "trade_volume": "成交股數",
        "trade_value": "成交金額",
        "transactions": "成交筆數",
        "exchange_source": "交易所資料來源",
        "exchange_note": "交易所資料說明",
        "upside_to_mean_pct": "平均目標價潛在漲跌幅",
        "upside_to_median_pct": "中位目標價潛在漲跌幅",
        "upside_to_high_pct": "最高目標價潛在漲跌幅",
        "downside_to_low_pct": "最低目標價潛在漲跌幅",
        "valuation_age_days": "評價距今天數",
        "valuation_signal": "估值訊號",
        "confidence_note": "可信度註記",
        "cnyes_url": "鉅亨個股頁",
        "target_date": "目標價日期",
        "target_high": "最高目標價",
        "target_low": "最低目標價",
        "target_mean": "平均目標價",
        "target_median": "中位目標價",
        "num_est": "預估家數",
        "fe_up": "上修家數",
        "fe_down": "下修家數",
        "fe_stddev": "目標價標準差",
        "currency": "幣別",
        "cnyes_last": "鉅亨頁面價格",
        "cnyes_status": "鉅亨抓取狀態",
        "cnyes_error": "鉅亨錯誤訊息",
        "cnyes_attempts": "鉅亨請求次數",
        "generated_at": "產出時間",
        "source": "資料源",
        "status": "狀態",
        "rows": "筆數",
        "data_date": "資料日期",
        "url": "網址",
        "message": "訊息",
        "section": "區塊",
        "item": "項目",
        "description": "說明",
        "field": "欄位",
    }
    return labels.get(column, column)


def status_rows(statuses: list[SourceStatus], all_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        {
            "generated_at": generated_at,
            "source": status.source,
            "status": status.status,
            "rows": status.rows,
            "data_date": status.data_date,
            "url": status.url,
            "message": status.message,
        }
        for status in statuses
    ]
    rows.append(
        {
            "generated_at": generated_at,
            "source": "report",
            "status": "ok",
            "rows": len(all_rows),
            "data_date": report_date(all_rows),
            "url": "",
            "message": "one row per listed or OTC company in the selected universe",
        }
    )
    return rows


def dictionary_rows() -> list[dict[str, str]]:
    definitions = {
        "market": "市場別；Excel 直接顯示上市或上櫃，內部資料處理時 tse=上市、otc=上櫃。",
        "stock_id": "股票代號。",
        "name": "股票名稱。",
        "company_name": "公司全名，來自 TWSE/TPEx 基本資料。",
        "industry": "產業別，來自 TWSE/TPEx 基本資料。",
        "listing_date": "上市或上櫃日期。",
        "close_date": "最新收盤行情日期，也是輸出檔名使用的主要日期。",
        "close": "TWSE 或 TPEx 每日收盤價，不是即時價。",
        "trade_volume": "成交股數，Excel 以千分位顯示。",
        "trade_value": "成交金額，Excel 以千分位顯示。",
        "transactions": "成交筆數，Excel 以千分位顯示。",
        "exchange_source": "交易所資料來源；上市為 TWSE，上櫃為 TPEx，並列出使用的資料端點名稱。",
        "exchange_note": "交易所資料說明；說明該列基本資料與收盤行情分別來自哪個來源。",
        "target_mean": "鉅亨 targetValuation 的平均共識目標價（feMean）。",
        "target_median": "鉅亨 targetValuation 的中位共識目標價（feMedian）。",
        "target_high": "鉅亨 targetValuation 的最高共識目標價（feHigh）。",
        "target_low": "鉅亨 targetValuation 的最低共識目標價（feLow）。",
        "target_date": "鉅亨 targetValuation 的評價日期。",
        "num_est": "鉅亨 targetValuation 的預估家數（numEst）。",
        "upside_to_mean_pct": "鉅亨平均目標價 target_mean / TWSE或TPEx 收盤價 close - 1；主要低估/高估判斷欄位。",
        "upside_to_median_pct": "鉅亨中位目標價 target_median / TWSE或TPEx 收盤價 close - 1。",
        "upside_to_high_pct": "鉅亨最高目標價 target_high / TWSE或TPEx 收盤價 close - 1。",
        "downside_to_low_pct": "鉅亨最低目標價 target_low / TWSE或TPEx 收盤價 close - 1。",
        "valuation_age_days": "TWSE或TPEx 最新收盤日期 close_date - 鉅亨目標價日期 target_date；天數越大代表資料越舊。",
        "valuation_signal": "估值訊號：undervalued、overvalued、neutral、stale、low_confidence、missing_target。",
        "confidence_note": "信心註記，例如 fresh_33_estimates、stale_120_days、only_2_estimates、missing_cnyes_target。",
        "cnyes_status": "鉅亨抓取狀態：ok、no_target_valuation、parse_error、http_error、skipped_limit、skipped_error_threshold 等。",
        "cnyes_attempts": "鉅亨個股頁嘗試抓取次數。",
    }
    return [{"field": key, "description": value} for key, value in definitions.items()]


def guide_rows() -> list[dict[str, str]]:
    return [
        {
            "section": "分頁",
            "item": "使用說明",
            "description": "本分頁放在第一個，說明中文分頁、中文表頭、估值公式與抓取狀態；打開檔案會預設停在「低估清單」。",
        },
        {
            "section": "分頁",
            "item": "交易所資料",
            "description": "只放 TWSE/TPEx 來源欄位，市場欄直接顯示上市/上櫃，並包含基本資料、每日收盤行情與資料來源說明；不包含鉅亨目標價與估值計算。",
        },
        {
            "section": "分頁",
            "item": "低估清單",
            "description": "用鉅亨平均共識目標價 join TWSE/TPEx 收盤價後計算；平均目標價潛在漲跌幅達低估門檻、且未被判定過舊或低信心。",
        },
        {
            "section": "分頁",
            "item": "高估清單",
            "description": "用鉅亨平均共識目標價 join TWSE/TPEx 收盤價後計算；平均目標價潛在漲跌幅低於高估門檻，依高估程度排序。",
        },
        {
            "section": "分頁",
            "item": "全部股票",
            "description": "上市與上櫃股票一支一列，交易所基本資料、收盤行情與鉅亨目標價對齊在同一列；鉅亨沒有資料也會保留。",
        },
        {
            "section": "分頁",
            "item": "過舊低信心",
            "description": "集中列出評價超過門檻天數、預估家數不足、缺少鉅亨目標價或抓取異常的股票。",
        },
        {
            "section": "分頁",
            "item": "抓取狀態",
            "description": "記錄 TWSE、TPEx、鉅亨等資料源的抓取狀態、資料日期、筆數與錯誤訊息。",
        },
        {
            "section": "分頁",
            "item": "欄位說明",
            "description": "補充主要欄位、門檻與資料來源；若要查內部欄位名稱，可看括號內英文 key。",
        },
        {
            "section": "閱讀順序",
            "item": "建議先看",
            "description": "先看「低估清單」與「高估清單」，再用「交易所資料」確認 TWSE/TPEx 原始欄位，最後到「全部股票」看 join 後完整資料。",
        },
        {
            "section": "資料來源",
            "item": "收盤價",
            "description": "使用 TWSE/TPEx 每日收盤行情，不是即時報價；檔名日期代表本次報表使用的最新收盤日期。",
        },
        {
            "section": "資料來源",
            "item": "共識目標價",
            "description": "使用鉅亨個股頁內嵌 targetValuation 共識目標價，屬於彙總共識，不是券商逐筆研究報告。",
        },
        {
            "section": "公式",
            "item": "平均目標價潛在漲跌幅 (upside_to_mean_pct)",
            "description": "鉅亨 targetValuation 平均目標價 target_mean / TWSE或TPEx 收盤價 close - 1；Excel 以百分比顯示，是主要低估/高估判斷欄位。",
        },
        {
            "section": "公式",
            "item": "中位目標價潛在漲跌幅 (upside_to_median_pct)",
            "description": "鉅亨 targetValuation 中位目標價 target_median / TWSE或TPEx 收盤價 close - 1；可降低極端目標價對判讀的影響。",
        },
        {
            "section": "公式",
            "item": "最高目標價潛在漲跌幅 (upside_to_high_pct)",
            "description": "鉅亨 targetValuation 最高目標價 target_high / TWSE或TPEx 收盤價 close - 1；偏樂觀情境參考。",
        },
        {
            "section": "公式",
            "item": "最低目標價潛在漲跌幅 (downside_to_low_pct)",
            "description": "鉅亨 targetValuation 最低目標價 target_low / TWSE或TPEx 收盤價 close - 1；用來看保守情境下是否仍有下修空間。",
        },
        {
            "section": "公式",
            "item": "評價距今天數 (valuation_age_days)",
            "description": "TWSE或TPEx 最新收盤日期 close_date - 鉅亨目標價日期 target_date；天數越大代表共識目標價越舊。",
        },
        {
            "section": "判斷",
            "item": "估值訊號 (valuation_signal)",
            "description": "undervalued=低估、overvalued=高估、neutral=中性、stale=過舊、low_confidence=低信心、missing_target=缺目標價。",
        },
        {
            "section": "判斷",
            "item": "信心註記 (confidence_note)",
            "description": "例如 fresh_33_estimates、stale_120_days、only_2_estimates、missing_cnyes_target，用來快速看資料新鮮度與預估家數。",
        },
        {
            "section": "中文表頭",
            "item": "市場 (market)",
            "description": "tse=上市，otc=上櫃。",
        },
        {
            "section": "中文表頭",
            "item": "股票代號 / 股票名稱 / 公司全名",
            "description": "股票識別與公司基本資料，來自 TWSE 或 TPEx 基本資料來源。",
        },
        {
            "section": "中文表頭",
            "item": "產業別 / 上市櫃日期",
            "description": "公司所屬產業與上市或上櫃日期。",
        },
        {
            "section": "中文表頭",
            "item": "最新收盤日期 / 收盤價 / 漲跌 / 開盤價 / 最高價 / 最低價",
            "description": "交易所每日收盤行情欄位；數字欄已套用千分位或小數格式。",
        },
        {
            "section": "中文表頭",
            "item": "成交股數 / 成交金額 / 成交筆數",
            "description": "交易所每日成交量、成交金額與成交筆數；Excel 顯示千分位。",
        },
        {
            "section": "中文表頭",
            "item": "平均目標價 / 中位目標價 / 最高目標價 / 最低目標價",
            "description": "鉅亨 targetValuation 對應欄位；若沒有鉅亨資料，欄位會留空。",
        },
        {
            "section": "中文表頭",
            "item": "預估家數 (num_est)",
            "description": "形成共識目標價的預估家數；低於門檻會被標示為低信心。",
        },
        {
            "section": "中文表頭",
            "item": "鉅亨抓取狀態 (cnyes_status)",
            "description": "ok=成功取得目標價；no_target_valuation=頁面沒有目標價；skipped_limit=測試限制未抓；http_error/parse_error=抓取或解析異常。",
        },
        {
            "section": "格式",
            "item": "中文化與閱讀性",
            "description": "分頁名稱與表頭以中文顯示；表頭置中且不使用填滿網底；金額、股數與筆數欄位顯示千分位。",
        },
    ]


def report_date(rows: list[dict[str, Any]]) -> str:
    dates = sorted({str(row.get("close_date") or "") for row in rows if row.get("close_date")})
    return dates[-1] if dates else datetime.now().strftime("%Y-%m-%d")


def cell_ref(row_idx: int, col_idx: int) -> str:
    letters = ""
    col = col_idx
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row_idx}"


def xml_text(value: Any) -> str:
    return escape(str(value), {'"': "&quot;"})


def display_cell_value(column: str, value: Any) -> Any:
    if column == "market":
        return market_name(str(value or ""))
    return value


def xlsx_sheet_xml(name: str, rows: list[dict[str, Any]], columns: list[str], tab_selected: bool = False) -> str:
    table = [[column_label(column) for column in columns]] + [
        [display_cell_value(col, row.get(col, "")) for col in columns] for row in rows
    ]
    max_row = max(1, len(table))
    max_col = max(1, len(columns))
    selected_attr = ' tabSelected="1"' if tab_selected else ""
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        f'<dimension ref="A1:{cell_ref(max_row, max_col)}"/>',
        f'<sheetViews><sheetView workbookViewId="0"{selected_attr}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        "<cols>",
    ]
    for idx, column in enumerate(columns, start=1):
        label = column_label(column)
        width = min(42, max(10, len(label) * 2 + 2))
        if column in {"cnyes_url", "company_name", "cnyes_error", "confidence_note"}:
            width = 28
        parts.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')
    parts.append("</cols><sheetData>")

    percent_cols = {col for col in columns if col.endswith("_pct")}
    decimal_cols = {
        "close",
        "change",
        "open",
        "high",
        "low",
        "target_high",
        "target_low",
        "target_mean",
        "target_median",
        "fe_stddev",
        "cnyes_last",
    }
    integer_cols = {
        "trade_volume",
        "trade_value",
        "transactions",
        "num_est",
        "fe_up",
        "fe_down",
        "valuation_age_days",
        "cnyes_attempts",
        "rows",
    }
    number_cols = decimal_cols.union(integer_cols)
    for row_idx, row_values in enumerate(table, start=1):
        parts.append(f'<row r="{row_idx}">')
        for col_idx, value in enumerate(row_values, start=1):
            ref = cell_ref(row_idx, col_idx)
            column = columns[col_idx - 1]
            style = 1 if row_idx == 1 else 0
            if row_idx > 1 and column in percent_cols:
                style = 3
            elif row_idx > 1 and column in integer_cols:
                style = 4
            elif row_idx > 1 and column in decimal_cols:
                style = 2

            if value is None or value == "":
                parts.append(f'<c r="{ref}" s="{style}"/>')
            elif row_idx > 1 and isinstance(value, (int, float)) and column in percent_cols.union(number_cols):
                parts.append(f'<c r="{ref}" s="{style}"><v>{value}</v></c>')
            else:
                parts.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    parts.append("</worksheet>")
    return "".join(parts)


def workbook_xml(sheet_names: list[str], active_tab: int = 0) -> str:
    sheets = "".join(
        f'<sheet name="{xml_text(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<bookViews><workbookView activeTab="{active_tab}"/></bookViews>'
        f"<sheets>{sheets}</sheets></workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for idx in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    rels.append("</Relationships>")
    return "".join(rels)


def content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for idx in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="0.00%"/><numFmt numFmtId="166" formatCode="#,##0"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def package_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def write_xlsx(path: Path, sheets: list[tuple[str, list[dict[str, Any]], list[str]]], active_sheet_index: int = 3) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        zf.writestr("_rels/.rels", package_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml([name for name, _, _ in sheets], active_tab=max(0, active_sheet_index - 1)))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        zf.writestr("xl/styles.xml", styles_xml())
        for idx, (name, rows, columns) in enumerate(sheets, start=1):
            zf.writestr(posixpath.join("xl", "worksheets", f"sheet{idx}.xml"), xlsx_sheet_xml(name, rows, columns, tab_selected=idx == active_sheet_index))


def build_workbook_rows(rows: list[dict[str, Any]], statuses: list[SourceStatus], args: argparse.Namespace) -> list[tuple[str, list[dict[str, Any]], list[str]]]:
    columns = column_order()
    undervalued = sorted(
        [row for row in rows if row.get("valuation_signal") == "undervalued"],
        key=lambda row: row.get("upside_to_mean_pct") or -999,
        reverse=True,
    )
    overvalued = sorted(
        [row for row in rows if row.get("valuation_signal") == "overvalued"],
        key=lambda row: row.get("upside_to_mean_pct") or 999,
    )
    stale_low = [
        row
        for row in rows
        if row.get("valuation_signal") in {"stale", "low_confidence", "missing_target"}
        or row.get("cnyes_status") != "ok"
    ]
    return [
        ("\u4f7f\u7528\u8aaa\u660e", guide_rows(), ["section", "item", "description"]),
        ("\u4ea4\u6613\u6240\u8cc7\u6599", exchange_rows(rows), exchange_column_order()),
        ("\u4f4e\u4f30\u6e05\u55ae", undervalued, columns),
        ("\u9ad8\u4f30\u6e05\u55ae", overvalued, columns),
        ("\u5168\u90e8\u80a1\u7968", rows, columns),
        ("\u904e\u820a\u4f4e\u4fe1\u5fc3", stale_low, columns),
        ("\u6293\u53d6\u72c0\u614b", status_rows(statuses, rows), ["generated_at", "source", "status", "rows", "data_date", "url", "message"]),
        ("\u6b04\u4f4d\u8aaa\u660e", dictionary_rows(), ["field", "description"]),
    ]


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a Taiwan stock valuation-gap Excel report.")
    parser.add_argument("--universe", choices=["all", "watchlist"], default="all")
    parser.add_argument("--market", choices=["all", "tse", "otc"], default="all")
    parser.add_argument("--watchlist", default="config/watchlist.csv")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--undervalued-threshold", type=float, default=20.0)
    parser.add_argument("--overvalued-threshold", type=float, default=-10.0)
    parser.add_argument("--stale-days", type=int, default=90)
    parser.add_argument("--min-estimates", type=int, default=3)
    parser.add_argument("--skip-cnyes", action="store_true")
    parser.add_argument("--cnyes-delay", type=float, default=0.5)
    parser.add_argument("--cnyes-retries", type=int, default=2)
    parser.add_argument("--cnyes-backoff", type=float, default=2.0)
    parser.add_argument("--cnyes-progress-every", type=int, default=50)
    parser.add_argument("--cnyes-limit", type=int, default=0)
    parser.add_argument("--cnyes-error-stop-after", type=int, default=30)
    parser.add_argument("--cnyes-error-stop-rate", type=float, default=0.5)
    parser.add_argument("--timeout", type=int, default=30)
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    statuses: list[SourceStatus] = []
    profiles: dict[tuple[str, str], StockInfo] = {}
    closes: dict[tuple[str, str], dict[str, Any]] = {}

    if args.market in {"all", "tse"}:
        profiles.update(fetch_twse_profiles(statuses))
        closes.update(fetch_twse_closes(statuses))
    if args.market in {"all", "otc"}:
        profiles.update(fetch_tpex_profiles(statuses))
        closes.update(fetch_tpex_closes(statuses))

    watchlist = load_watchlist(Path(args.watchlist))
    stocks = select_universe(profiles, watchlist, args.universe, args.market)
    if not stocks:
        print("No stocks found for the selected universe.", file=sys.stderr)
        return 2

    targets = fetch_cnyes_targets(
        stocks,
        args.skip_cnyes,
        args.cnyes_delay,
        args.timeout,
        args.cnyes_retries,
        args.cnyes_backoff,
        args.cnyes_progress_every,
        args.cnyes_limit,
        args.cnyes_error_stop_after,
        args.cnyes_error_stop_rate,
        statuses,
    )
    rows = build_rows(stocks, closes, targets, args)
    date_part = report_date(rows).replace("-", "")
    suffix = ""
    if args.universe == "watchlist":
        suffix += "_watchlist"
    if args.skip_cnyes:
        suffix += "_no_cnyes"
    if args.cnyes_limit > 0 and not args.skip_cnyes:
        suffix += f"_cnyes_limit{args.cnyes_limit}"
    output_path = Path(args.output_dir) / f"tw_valuation_gap_{date_part}{suffix}.xlsx"
    sheets = build_workbook_rows(rows, statuses, args)
    try:
        write_xlsx(output_path, sheets)
    except PermissionError:
        timestamp = datetime.now().strftime("%H%M%S")
        output_path = output_path.with_name(f"{output_path.stem}_{timestamp}{output_path.suffix}")
        write_xlsx(output_path, sheets)

    print(f"Wrote {output_path}")
    print(f"Rows: {len(rows)}")
    print(f"Undervalued: {sum(1 for row in rows if row.get('valuation_signal') == 'undervalued')}")
    print(f"Overvalued: {sum(1 for row in rows if row.get('valuation_signal') == 'overvalued')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
