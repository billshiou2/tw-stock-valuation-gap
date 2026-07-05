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
import os
import posixpath
import re
import sys
import time
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timezone
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

ENV_DEFAULTS = {
    "TW_STOCK_UNIVERSE": "all",
    "TW_STOCK_MARKET": "all",
    "TW_STOCK_UNDERVALUED_THRESHOLD": "20.0",
    "TW_STOCK_OVERVALUED_THRESHOLD": "-10.0",
    "TW_STOCK_STALE_DAYS": "90",
    "TW_STOCK_MIN_ESTIMATES": "3",
    "TW_STOCK_EXCEL_OUTPUT": "both",
    "TW_STOCK_CNYES_DELAY": "0.5",
    "TW_STOCK_CNYES_RETRIES": "2",
    "TW_STOCK_CNYES_BACKOFF": "2.0",
    "TW_STOCK_CNYES_PROGRESS_EVERY": "50",
    "TW_STOCK_CNYES_LIMIT": "0",
    "TW_STOCK_CNYES_ERROR_STOP_AFTER": "30",
    "TW_STOCK_CNYES_ERROR_STOP_RATE": "0.5",
}

INDUSTRY_CODES = {
    "01": "水泥工業",
    "02": "食品工業",
    "03": "塑膠工業",
    "04": "紡織纖維",
    "05": "電機機械",
    "06": "電器電纜",
    "08": "玻璃陶瓷",
    "09": "造紙工業",
    "10": "鋼鐵工業",
    "11": "橡膠工業",
    "12": "汽車工業",
    "14": "建材營造",
    "15": "航運業",
    "16": "觀光餐旅",
    "17": "金融保險",
    "18": "貿易百貨",
    "20": "其他",
    "21": "化學工業",
    "22": "生技醫療業",
    "23": "油電燃氣業",
    "24": "半導體業",
    "25": "電腦及週邊設備業",
    "26": "光電業",
    "27": "通信網路業",
    "28": "電子零組件業",
    "29": "電子通路業",
    "30": "資訊服務業",
    "31": "其他電子業",
    "32": "文化創意業",
    "33": "農業科技業",
    "34": "電子商務業",
    "35": "綠能環保",
    "36": "數位雲端",
    "37": "運動休閒",
    "38": "居家生活",
}


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


def load_env_file(path: Path = Path(".env")) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            values[key] = value
    return values


def env_value(env: dict[str, str], key: str) -> str:
    return os.environ.get(key) or env.get(key) or ENV_DEFAULTS[key]


def env_float(env: dict[str, str], key: str) -> float:
    return float(env_value(env, key))


def env_int(env: dict[str, str], key: str) -> int:
    return int(float(env_value(env, key)))


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
            "industry_name": industry_name_for(stock.industry),
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
        "industry_name",
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
        "upside_to_mean_pct",
        "upside_to_median_pct",
        "upside_to_high_pct",
        "downside_to_low_pct",
        "valuation_age_days",
        "valuation_signal",
        "confidence_note",
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
        "cnyes_attempts",
    ]


def exchange_column_order() -> list[str]:
    return [
        "market",
        "stock_id",
        "name",
        "company_name",
        "industry",
        "industry_name",
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


def stale_column_order() -> list[str]:
    return [
        "market",
        "stock_id",
        "name",
        "industry",
        "industry_name",
        "close_date",
        "close",
        "upside_to_mean_pct",
        "target_date",
        "target_mean",
        "valuation_age_days",
        "num_est",
        "valuation_signal",
        "confidence_note",
        "cnyes_status",
    ]


def lite_column_order() -> list[str]:
    return [
        "market",
        "stock_id",
        "name",
        "company_name",
        "industry_name",
        "close_date",
        "close",
        "change",
        "trade_volume",
        "trade_value",
        "target_date",
        "target_mean",
        "target_median",
        "target_high",
        "target_low",
        "upside_to_mean_pct",
        "upside_to_median_pct",
        "downside_to_low_pct",
        "valuation_age_days",
        "num_est",
        "valuation_signal",
        "confidence_note",
        "cnyes_status",
    ]


def industry_name_for(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return INDUSTRY_CODES.get(text, text)


def industry_rows(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    seen_codes = sorted(
        {
            str(row.get("industry") or "").strip()
            for row in rows
            if str(row.get("industry") or "").strip().isdigit()
        }
    )
    result = []
    for code in seen_codes:
        result.append(
            {
                "industry_code": code,
                "industry_name": INDUSTRY_CODES.get(code, ""),
                "industry_note": "TWSE/TPEx 產業別代碼；空白代表目前對照表尚未列入，仍保留原始代碼。",
            }
        )
    text_industries = sorted(
        {
            str(row.get("industry") or "").strip()
            for row in rows
            if str(row.get("industry") or "").strip()
            and not str(row.get("industry") or "").strip().isdigit()
        }
    )
    for value in text_industries:
        result.append(
            {
                "industry_code": value,
                "industry_name": value,
                "industry_note": "TWSE 已提供文字產業別，非數字代碼。",
            }
        )
    return result


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
        "industry_code": "產業別代碼",
        "industry_name": "產業名稱",
        "industry_note": "產業別說明",
        "listing_date": "上市櫃日期",
        "close_date": "最新收盤日期",
        "close": "收盤價",
        "change": "漲跌",
        "open": "開盤價",
        "high": "最高價",
        "low": "最低價",
        "trade_volume": "成交股數(股)",
        "trade_value": "成交金額(元)",
        "transactions": "成交筆數(筆)",
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
        "industry": "產業別；保留 TWSE/TPEx 基本資料原始值，可能是文字名稱或產業代碼。",
        "industry_name": "產業名稱；由產業別代碼轉成中文名稱，若原本就是文字則沿用原值。",
        "listing_date": "上市或上櫃日期。",
        "close_date": "最新收盤行情日期，也是輸出檔名使用的主要日期。",
        "close": "TWSE 或 TPEx 每日收盤價，不是即時價。",
        "trade_volume": "成交股數，單位是股，不是張；TWSE TradeVolume 與 TPEx TradingShares 皆為股數，Excel 以千分位顯示。",
        "trade_value": "成交金額，單位是新台幣元；TWSE TradeValue 與 TPEx TransactionAmount 皆為元，Excel 以千分位顯示。",
        "transactions": "成交筆數，單位是筆；TWSE Transaction 與 TPEx TransactionNumber 皆為筆，Excel 以千分位顯示。",
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
            "description": "先看「低估清單」與「高估清單」，最後到「全部股票」看 join 後完整資料。",
        },
        {
            "section": "資料來源",
            "item": "收盤價",
            "description": "使用 TWSE/TPEx 每日收盤行情，不是即時報價；檔名日期代表本次報表使用的最新收盤日期。",
        },
        {
            "section": "資料來源",
            "item": "TWSE 是什麼",
            "description": "TWSE 是台灣證券交易所（Taiwan Stock Exchange）；本報表中市場顯示為「上市」的股票，基本資料與每日收盤行情主要來自 TWSE。",
        },
        {
            "section": "資料來源",
            "item": "TPEx 是什麼",
            "description": "TPEx 是證券櫃檯買賣中心（Taipei Exchange）；本報表中市場顯示為「上櫃」的股票，基本資料與每日收盤行情主要來自 TPEx。",
        },
        {
            "section": "資料來源",
            "item": "共識目標價",
            "description": "使用鉅亨個股頁內嵌 targetValuation 共識目標價，屬於彙總共識，不是券商逐筆研究報告。",
        },
        {
            "section": "爬取方式",
            "item": "TWSE 上市收盤行情",
            "description": "使用 TWSE OpenAPI JSON 端點 STOCK_DAY_ALL；一次下載上市每日收盤行情，不逐檔爬個股頁。",
        },
        {
            "section": "爬取方式",
            "item": "TWSE 上市基本資料",
            "description": "使用 TWSE OpenAPI JSON 端點 t187ap03_L；一次下載上市公司基本資料，用來補股票名稱、公司全名、產業別與上市日期。",
        },
        {
            "section": "爬取方式",
            "item": "TPEx 上櫃收盤行情",
            "description": "使用 TPEx OpenAPI JSON 端點 tpex_mainboard_daily_close_quotes；一次下載上櫃每日收盤行情，不逐檔爬個股頁。",
        },
        {
            "section": "爬取方式",
            "item": "TPEx 上櫃基本資料",
            "description": "使用 TPEx OpenAPI JSON 端點 mopsfin_t187ap03_O；一次下載上櫃公司基本資料，用來補股票名稱、公司全名、產業別與上櫃日期。",
        },
        {
            "section": "爬取方式",
            "item": "鉅亨目標價",
            "description": "不是正式公開 API；程式用 Python urllib 加瀏覽器樣式 headers 下載 https://www.cnyes.com/twstock/{stock_id} 個股頁 HTML，再解析頁面內嵌的 targetValuation JSON。",
        },
        {
            "section": "爬取方式",
            "item": "鉅亨解析欄位",
            "description": "從 targetValuation 取 rateDate、feMean、feMedian、feHigh、feLow、numEst、feStdDev、currency、last 等共識目標價欄位。",
        },
        {
            "section": "爬取頻率",
            "item": "TWSE/TPEx",
            "description": "每個交易所來源各發一次 JSON 請求；若下載或解析失敗最多嘗試 3 次，重試等待約 1.5 秒、3.0 秒。",
        },
        {
            "section": "爬取頻率",
            "item": "鉅亨 CLI 預設",
            "description": "逐檔 sequential 請求；預設每檔間隔 0.5 秒，失敗重試 2 次（最多 3 次請求），backoff 以 2 秒倍增，約等待 2 秒、4 秒。",
        },
        {
            "section": "爬取頻率",
            "item": "run_full_scan.bat",
            "description": "完整全市場 bat 使用較保守設定：每檔間隔 1.0 秒、重試 2 次、backoff 2.0 秒、每 25 檔輸出進度。",
        },
        {
            "section": "保護機制",
            "item": "鉅亨錯誤停止",
            "description": "處理至少 30 檔後，若 HTTP error 比率 >= 50%，停止後續鉅亨請求，尚未抓取的股票標示 skipped_error_threshold，避免持續產生大量錯誤。",
        },
        {
            "section": "保護機制",
            "item": "測試限制",
            "description": "--cnyes-limit 可限制鉅亨抓取檔數，未抓的股票標示 skipped_limit；--skip-cnyes 則只產交易所資料並保留股票列。",
        },
        {
            "section": "Join 邏輯",
            "item": "股票代號",
            "description": "交易所資料以市場加股票代號建立 universe；鉅亨資料以股票代號 join 到同一列，缺鉅亨資料不刪股票列，只在狀態欄標示原因。",
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
            "item": "為何沒進低估/高估",
            "description": "程式會先檢查資料品質：缺目標價、評價日期超過 stale-days、或預估家數少於 min-estimates 時，會先歸到 missing_target/stale/low_confidence，不放進主要低估或高估清單。",
        },
        {
            "section": "判斷",
            "item": "判斷順序",
            "description": "順序是：鉅亨狀態與目標價是否存在 -> 評價是否過舊 -> 預估家數是否足夠 -> 才用平均目標價潛在漲跌幅判斷 undervalued / overvalued / neutral。",
        },
        {
            "section": "設定",
            "item": ".env 門檻設定",
            "description": "可在 .env 設定 TW_STOCK_UNDERVALUED_THRESHOLD、TW_STOCK_OVERVALUED_THRESHOLD、TW_STOCK_STALE_DAYS、TW_STOCK_MIN_ESTIMATES；CLI 參數若有指定會覆蓋 .env。",
        },
        {
            "section": "設定",
            "item": "Excel 輸出設定",
            "description": "TW_STOCK_EXCEL_OUTPUT 可設 full、lite、both；預設 both 會同時產完整檔與 _lite 輕量檔。",
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
            "item": "產業別 / 產業名稱 / 上市櫃日期",
            "description": "產業別保留交易所原始值，可能是代碼或文字；產業名稱是程式依代碼補出的中文名稱，放在產業別旁邊方便閱讀。",
        },
        {
            "section": "中文表頭",
            "item": "最新收盤日期 / 收盤價 / 漲跌 / 開盤價 / 最高價 / 最低價",
            "description": "交易所每日收盤行情欄位；數字欄已套用千分位或小數格式。",
        },
        {
            "section": "中文表頭",
            "item": "成交股數(股) / 成交金額(元) / 成交筆數(筆)",
            "description": "交易所每日成交欄位；成交股數單位是股不是張，成交金額單位是新台幣元，成交筆數單位是筆；Excel 顯示千分位。",
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


def column_ref(col_idx: int) -> str:
    return "".join(ch for ch in cell_ref(1, col_idx) if ch.isalpha())


def xml_text(value: Any) -> str:
    return escape(str(value), {'"': "&quot;"})


def display_cell_value(column: str, value: Any) -> Any:
    if column == "market":
        return market_name(str(value or ""))
    return value


def collect_shared_strings(sheets: list[tuple[str, list[dict[str, Any]], list[str]]]) -> tuple[dict[str, int], int]:
    strings: dict[str, int] = {}
    total_count = 0

    def add(value: Any) -> None:
        nonlocal total_count
        text = str(value)
        total_count += 1
        if text not in strings:
            strings[text] = len(strings)

    for _, rows, columns in sheets:
        for column in columns:
            add(column_label(column))
        for row in rows:
            for column in columns:
                value = display_cell_value(column, row.get(column, ""))
                if value is not None and value != "" and not isinstance(value, (int, float)):
                    add(value)
    return strings, total_count


def shared_strings_xml(shared_strings: dict[str, int], total_count: int) -> str:
    items = [""] * len(shared_strings)
    for text, idx in shared_strings.items():
        items[idx] = f"<si><t>{xml_text(text)}</t></si>"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        f' count="{total_count}" uniqueCount="{len(shared_strings)}">'
        + "".join(items)
        + "</sst>"
    )


def xlsx_sheet_xml(
    name: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    shared_strings: dict[str, int],
    tab_selected: bool = False,
) -> str:
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
                parts.append(f'<c r="{ref}" s="{style}" t="s"><v>{shared_strings[str(value)]}</v></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    ignored_ranges = []
    if max_row > 1:
        for idx, column in enumerate(columns, start=1):
            if column in {"stock_id", "industry"}:
                letter = column_ref(idx)
                ignored_ranges.append(f"{letter}2:{letter}{max_row}")
    if ignored_ranges:
        parts.append(
            f'<ignoredErrors><ignoredError sqref="{" ".join(ignored_ranges)}" numberStoredAsText="1"/></ignoredErrors>'
        )
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


def workbook_rels_xml(sheet_count: int, has_shared_strings: bool = False) -> str:
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
    if has_shared_strings:
        rels.append(
            f'<Relationship Id="rId{sheet_count + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        )
    rels.append("</Relationships>")
    return "".join(rels)


def content_types_xml(sheet_count: int, has_shared_strings: bool = False) -> str:
    overrides = [
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    if has_shared_strings:
        overrides.append(
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        )
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
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def core_properties_xml() -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        ' xmlns:dcterms="http://purl.org/dc/terms/"'
        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>tw_target_scan</dc:creator>"
        "<cp:lastModifiedBy>tw_target_scan</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_properties_xml(sheet_names: list[str]) -> str:
    titles = "".join(f"<vt:lpstr>{xml_text(name)}</vt:lpstr>" for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Microsoft Excel</Application>"
        "<DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop>"
        "<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\">"
        "<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>"
        f"<vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>"
        "</vt:vector></HeadingPairs>"
        f"<TitlesOfParts><vt:vector size=\"{len(sheet_names)}\" baseType=\"lpstr\">{titles}</vt:vector></TitlesOfParts>"
        "<Company></Company>"
        "<LinksUpToDate>false</LinksUpToDate>"
        "<SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged>"
        "<AppVersion>16.0300</AppVersion>"
        "</Properties>"
    )


def package_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def write_xlsx(path: Path, sheets: list[tuple[str, list[dict[str, Any]], list[str]]], active_sheet_index: int = 2) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    shared_strings, shared_string_count = collect_shared_strings(sheets)
    sheet_names = [name for name, _, _ in sheets]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets), has_shared_strings=bool(shared_strings)))
        zf.writestr("_rels/.rels", package_rels_xml())
        zf.writestr("docProps/core.xml", core_properties_xml())
        zf.writestr("docProps/app.xml", app_properties_xml(sheet_names))
        zf.writestr("xl/workbook.xml", workbook_xml(sheet_names, active_tab=max(0, active_sheet_index - 1)))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets), has_shared_strings=bool(shared_strings)))
        zf.writestr("xl/styles.xml", styles_xml())
        if shared_strings:
            zf.writestr("xl/sharedStrings.xml", shared_strings_xml(shared_strings, shared_string_count))
        for idx, (name, rows, columns) in enumerate(sheets, start=1):
            zf.writestr(
                posixpath.join("xl", "worksheets", f"sheet{idx}.xml"),
                xlsx_sheet_xml(name, rows, columns, shared_strings, tab_selected=idx == active_sheet_index),
            )


def build_workbook_rows(rows: list[dict[str, Any]], statuses: list[SourceStatus], args: argparse.Namespace) -> list[tuple[str, list[dict[str, Any]], list[str]]]:
    for row in rows:
        if not row.get("industry_name"):
            row["industry_name"] = industry_name_for(row.get("industry"))
        market = str(row.get("market") or "")
        if not row.get("exchange_source"):
            row["exchange_source"] = exchange_source(market)
        if not row.get("exchange_note"):
            row["exchange_note"] = exchange_note(market)
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
        ("\u4f4e\u4f30\u6e05\u55ae", undervalued, columns),
        ("\u9ad8\u4f30\u6e05\u55ae", overvalued, columns),
        ("\u5168\u90e8\u80a1\u7968", rows, columns),
        ("\u904e\u820a\u4f4e\u4fe1\u5fc3", stale_low, stale_column_order()),
        ("\u6293\u53d6\u72c0\u614b", status_rows(statuses, rows), ["generated_at", "source", "status", "rows", "data_date", "url", "message"]),
        ("\u6b04\u4f4d\u8aaa\u660e", dictionary_rows(), ["field", "description"]),
    ]


def build_lite_workbook_rows(rows: list[dict[str, Any]], statuses: list[SourceStatus], args: argparse.Namespace) -> list[tuple[str, list[dict[str, Any]], list[str]]]:
    for row in rows:
        if not row.get("industry_name"):
            row["industry_name"] = industry_name_for(row.get("industry"))
    columns = lite_column_order()
    undervalued = sorted(
        [row for row in rows if row.get("valuation_signal") == "undervalued"],
        key=lambda row: row.get("upside_to_mean_pct") or -999,
        reverse=True,
    )
    overvalued = sorted(
        [row for row in rows if row.get("valuation_signal") == "overvalued"],
        key=lambda row: row.get("upside_to_mean_pct") or 999,
    )
    return [
        ("\u4f7f\u7528\u8aaa\u660e", guide_rows(), ["section", "item", "description"]),
        ("\u4f4e\u4f30\u6e05\u55ae", undervalued, columns),
        ("\u9ad8\u4f30\u6e05\u55ae", overvalued, columns),
        ("\u5168\u90e8\u80a1\u7968", rows, columns),
        ("\u6293\u53d6\u72c0\u614b", status_rows(statuses, rows), ["generated_at", "source", "status", "rows", "data_date", "url", "message"]),
    ]


def write_xlsx_with_fallback(path: Path, sheets: list[tuple[str, list[dict[str, Any]], list[str]]]) -> Path:
    try:
        write_xlsx(path, sheets)
        return path
    except PermissionError:
        timestamp = datetime.now().strftime("%H%M%S")
        fallback = path.with_name(f"{path.stem}_{timestamp}{path.suffix}")
        write_xlsx(fallback, sheets)
        return fallback


def parse_args(argv: list[str]) -> argparse.Namespace:
    env = load_env_file()
    parser = argparse.ArgumentParser(description="Build a Taiwan stock valuation-gap Excel report.")
    parser.add_argument("--universe", choices=["all", "watchlist"], default=env_value(env, "TW_STOCK_UNIVERSE"))
    parser.add_argument("--market", choices=["all", "tse", "otc"], default=env_value(env, "TW_STOCK_MARKET"))
    parser.add_argument("--watchlist", default="config/watchlist.csv")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--undervalued-threshold", type=float, default=env_float(env, "TW_STOCK_UNDERVALUED_THRESHOLD"))
    parser.add_argument("--overvalued-threshold", type=float, default=env_float(env, "TW_STOCK_OVERVALUED_THRESHOLD"))
    parser.add_argument("--stale-days", type=int, default=env_int(env, "TW_STOCK_STALE_DAYS"))
    parser.add_argument("--min-estimates", type=int, default=env_int(env, "TW_STOCK_MIN_ESTIMATES"))
    parser.add_argument("--excel-output", choices=["full", "lite", "both"], default=env_value(env, "TW_STOCK_EXCEL_OUTPUT"))
    parser.add_argument("--skip-cnyes", action="store_true")
    parser.add_argument("--cnyes-delay", type=float, default=env_float(env, "TW_STOCK_CNYES_DELAY"))
    parser.add_argument("--cnyes-retries", type=int, default=env_int(env, "TW_STOCK_CNYES_RETRIES"))
    parser.add_argument("--cnyes-backoff", type=float, default=env_float(env, "TW_STOCK_CNYES_BACKOFF"))
    parser.add_argument("--cnyes-progress-every", type=int, default=env_int(env, "TW_STOCK_CNYES_PROGRESS_EVERY"))
    parser.add_argument("--cnyes-limit", type=int, default=env_int(env, "TW_STOCK_CNYES_LIMIT"))
    parser.add_argument("--cnyes-error-stop-after", type=int, default=env_int(env, "TW_STOCK_CNYES_ERROR_STOP_AFTER"))
    parser.add_argument("--cnyes-error-stop-rate", type=float, default=env_float(env, "TW_STOCK_CNYES_ERROR_STOP_RATE"))
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
    output_base = Path(args.output_dir) / f"tw_valuation_gap_{date_part}{suffix}.xlsx"
    written_paths: list[Path] = []
    if args.excel_output in {"full", "both"}:
        written_paths.append(write_xlsx_with_fallback(output_base, build_workbook_rows(rows, statuses, args)))
    if args.excel_output in {"lite", "both"}:
        lite_path = output_base.with_name(f"{output_base.stem}_lite{output_base.suffix}")
        written_paths.append(write_xlsx_with_fallback(lite_path, build_lite_workbook_rows(rows, statuses, args)))

    for output_path in written_paths:
        print(f"Wrote {output_path}")
    print(f"Rows: {len(rows)}")
    print(f"Undervalued: {sum(1 for row in rows if row.get('valuation_signal') == 'undervalued')}")
    print(f"Overvalued: {sum(1 for row in rows if row.get('valuation_signal') == 'overvalued')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
