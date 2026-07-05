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


def column_label(column: str) -> str:
    labels = {
        "market": "市場",
        "stock_id": "股票代號",
        "name": "股票名稱",
        "company_name": "公司全名",
        "industry": "產業別",
        "listing_date": "上市櫃日期",
        "close_date": "收盤日期",
        "close": "收盤價",
        "change": "漲跌",
        "open": "開盤價",
        "high": "最高價",
        "low": "最低價",
        "trade_volume": "成交股數",
        "trade_value": "成交金額",
        "transactions": "成交筆數",
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
        "market": "tse=listed, otc=TPEx mainboard",
        "close": "Daily closing price from TWSE or TPEx",
        "target_mean": "Cnyes targetValuation feMean consensus target price",
        "target_median": "Cnyes targetValuation feMedian consensus target price",
        "num_est": "Cnyes targetValuation numEst estimate count",
        "upside_to_mean_pct": "target_mean / close - 1",
        "downside_to_low_pct": "target_low / close - 1",
        "valuation_age_days": "close_date - target_date",
        "valuation_signal": "undervalued, overvalued, neutral, stale, low_confidence, missing_target",
        "cnyes_status": "ok, no_target_valuation, parse_error, http_error, skipped, skipped_limit, skipped_error_threshold, or missing",
        "cnyes_attempts": "Number of HTTP attempts used for this Cnyes stock page",
    }
    return [{"field": key, "description": value} for key, value in definitions.items()]


def guide_rows() -> list[dict[str, str]]:
    return [
        {
            "section": "How to read",
            "item": "重點先看哪些欄位",
            "description": "先看 stock_id, name, close, target_mean, upside_to_mean_pct, target_date, valuation_age_days, num_est, valuation_signal.",
        },
        {
            "section": "User question",
            "item": "估值落差計算是不是與鉅亨網的數據差",
            "description": "是。估值落差計算是用交易所收盤價 close，對比鉅亨 Cnyes targetValuation 的共識目標價 target_* 算出來的差距。",
        },
        {
            "section": "Formula",
            "item": "upside_to_mean_pct",
            "description": "target_mean / close - 1。例：收盤價 100、平均目標價 130，結果是 30%，代表相對平均共識目標價有 30% 上漲空間。",
        },
        {
            "section": "Formula",
            "item": "downside_to_low_pct",
            "description": "target_low / close - 1。這是用最低共識目標價看悲觀情境；如果負很多，代表悲觀目標價比現在收盤價低很多。",
        },
        {"section": "Basic", "item": "market", "description": "市場別。tse = 上市，otc = 上櫃。"},
        {"section": "Basic", "item": "stock_id", "description": "股票代號。"},
        {"section": "Basic", "item": "name", "description": "股票簡稱。"},
        {"section": "Basic", "item": "company_name", "description": "公司完整名稱，來自 TWSE/TPEx 公司基本資料。"},
        {"section": "Basic", "item": "industry", "description": "產業別或產業代碼，來自 TWSE/TPEx 公司基本資料。"},
        {"section": "Basic", "item": "listing_date", "description": "上市或上櫃日期。"},
        {"section": "Close", "item": "close_date", "description": "收盤行情日期。"},
        {"section": "Close", "item": "close", "description": "收盤價，來自 TWSE/TPEx 每日收盤行情。"},
        {"section": "Close", "item": "change", "description": "當日漲跌。"},
        {"section": "Close", "item": "open", "description": "開盤價。"},
        {"section": "Close", "item": "high", "description": "最高價。"},
        {"section": "Close", "item": "low", "description": "最低價。"},
        {"section": "Close", "item": "trade_volume", "description": "成交股數。"},
        {"section": "Close", "item": "trade_value", "description": "成交金額。"},
        {"section": "Close", "item": "transactions", "description": "成交筆數。"},
        {"section": "Valuation gap", "item": "upside_to_mean_pct", "description": "平均目標價相對收盤價的空間；越高越可能低估。"},
        {"section": "Valuation gap", "item": "upside_to_median_pct", "description": "中位數目標價相對收盤價的空間。"},
        {"section": "Valuation gap", "item": "upside_to_high_pct", "description": "最高目標價相對收盤價的空間。"},
        {"section": "Valuation gap", "item": "downside_to_low_pct", "description": "最低目標價相對收盤價的空間；負值越大代表悲觀情境越低。"},
        {"section": "Valuation gap", "item": "valuation_age_days", "description": "收盤日距離鉅亨評價日期幾天；越大代表目標價越舊。"},
        {"section": "Valuation gap", "item": "valuation_signal", "description": "系統分類：undervalued, overvalued, neutral, stale, low_confidence, missing_target。"},
        {"section": "Valuation gap", "item": "confidence_note", "description": "分類補充，例如 fresh_33_estimates, stale_120_days, only_2_estimates, missing_cnyes_target。"},
        {"section": "Cnyes", "item": "cnyes_url", "description": "鉅亨個股頁網址。"},
        {"section": "Cnyes", "item": "target_date", "description": "鉅亨 targetValuation 的評價日期。"},
        {"section": "Cnyes", "item": "target_high", "description": "最高共識目標價，來自鉅亨 feHigh。"},
        {"section": "Cnyes", "item": "target_low", "description": "最低共識目標價，來自鉅亨 feLow。"},
        {"section": "Cnyes", "item": "target_mean", "description": "平均共識目標價，來自鉅亨 feMean；主要用來判斷低估/高估。"},
        {"section": "Cnyes", "item": "target_median", "description": "中位數共識目標價，來自鉅亨 feMedian。"},
        {"section": "Cnyes", "item": "num_est", "description": "預估家數，來自鉅亨 numEst；越多通常參考性越好。"},
        {"section": "Cnyes", "item": "fe_up", "description": "鉅亨原始欄位，正向/上調估計數。"},
        {"section": "Cnyes", "item": "fe_down", "description": "鉅亨原始欄位，負向/下調估計數。"},
        {"section": "Cnyes", "item": "fe_stddev", "description": "目標價標準差；越高代表市場目標價分歧越大。"},
        {"section": "Cnyes", "item": "currency", "description": "目標價幣別。"},
        {"section": "Cnyes", "item": "cnyes_last", "description": "鉅亨頁面搭配的最近價格，僅供對照；主計算使用交易所 close。"},
        {"section": "Status", "item": "cnyes_status", "description": "鉅亨抓取狀態。ok = 有目標價；no_target_valuation = 頁面沒有共識目標價；skipped_limit = 因 --cnyes-limit 未抓；skipped_error_threshold = HTTP error 比例過高所以停止後續請求。"},
        {"section": "Status", "item": "cnyes_error", "description": "鉅亨抓取或解析失敗時的錯誤訊息。"},
        {"section": "Status", "item": "cnyes_attempts", "description": "該股票鉅亨頁面實際嘗試請求次數；1 代表第一次就成功，2 以上代表有重試。"},
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


def xlsx_sheet_xml(name: str, rows: list[dict[str, Any]], columns: list[str]) -> str:
    table = [[column_label(column) for column in columns]] + [[row.get(col, "") for col in columns] for row in rows]
    max_row = max(1, len(table))
    max_col = max(1, len(columns))
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        f'<dimension ref="A1:{cell_ref(max_row, max_col)}"/>',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
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
    parts.append(f'<autoFilter ref="A1:{cell_ref(max_row, max_col)}"/>')
    parts.append("</worksheet>")
    return "".join(parts)


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = "".join(
        f'<sheet name="{xml_text(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        for idx, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
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


def write_xlsx(path: Path, sheets: list[tuple[str, list[dict[str, Any]], list[str]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        zf.writestr("_rels/.rels", package_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml([name for name, _, _ in sheets]))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        zf.writestr("xl/styles.xml", styles_xml())
        for idx, (name, rows, columns) in enumerate(sheets, start=1):
            zf.writestr(posixpath.join("xl", "worksheets", f"sheet{idx}.xml"), xlsx_sheet_xml(name, rows, columns))


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
        ("Undervalued", undervalued, columns),
        ("Overvalued", overvalued, columns),
        ("All_Stocks", rows, columns),
        ("Stale_or_Low_Confidence", stale_low, columns),
        ("Fetch_Status", status_rows(statuses, rows), ["generated_at", "source", "status", "rows", "data_date", "url", "message"]),
        ("Guide", guide_rows(), ["section", "item", "description"]),
        ("Data_Dictionary", dictionary_rows(), ["field", "description"]),
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
