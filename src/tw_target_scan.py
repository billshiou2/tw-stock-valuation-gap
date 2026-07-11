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
TWSE_VALUATION_URL = "https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL"
TPEX_VALUATION_URL = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis"
TWSE_REVENUE_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap05_L"
TPEX_REVENUE_URL = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap05_O"
TWSE_EPS_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap14_L"
TPEX_EPS_URL = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap14_O"
TWSE_MARGIN_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap17_L"
CNYES_STOCK_URL = "https://www.cnyes.com/twstock/{stock_id}"
CNYES_FACTSET_SOURCE = "鉅亨/FactSet共識"
CNYES_UNSPECIFIED_SOURCE = "鉅亨共識(來源未明確標示)"
CNYES_NO_TARGET_SOURCE = "鉅亨未提供目標價"
CNYES_NOT_FETCHED_SOURCE = "鉅亨未抓取"
CNYES_DATA_SOURCE_NOTE = (
    "鉅亨網個股頁內嵌 targetValuation 共識目標價；"
    "程式會逐檔檢查頁面是否出現 FactSet 或 factSetEstimate 註記，"
    "有才標示為鉅亨/FactSet共識，否則標示為來源未明確標示。"
    "頁尾 Refinitiv 聲明較偏一般報價與市場資訊來源。"
    "此資料不是券商逐筆研究報告，也不列個別機構名稱。"
)
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


def normalize_month(raw: str) -> str:
    raw = (raw or "").strip()
    if not raw:
        return ""
    if re.fullmatch(r"\d{5}", raw):
        year = int(raw[:3]) + 1911
        return f"{year:04d}-{raw[3:5]}"
    if re.fullmatch(r"\d{6}", raw):
        return f"{raw[:4]}-{raw[4:6]}"
    return raw


def normalize_year(raw: Any) -> str:
    text = str(raw or "").strip()
    if not text:
        return ""
    if re.fullmatch(r"\d{3}", text):
        return str(int(text) + 1911)
    return text


def normalize_period(year: Any, quarter: Any) -> str:
    normalized_year = normalize_year(year)
    quarter_text = str(quarter or "").strip()
    match = re.search(r"(\d+)", quarter_text)
    if not normalized_year or not match:
        return ""
    return f"{normalized_year}Q{match.group(1)}"


def percent_value(value: Any) -> float | None:
    number = to_float(value)
    if number is None:
        return None
    return number / 100.0


def valuation_note(
    valuation_row: dict[str, Any],
    per: float | None,
    pbr: float | None,
    raw_dividend_yield: Any,
    dividend_yield: float | None,
) -> str:
    if not valuation_row:
        return "官方估值端點未提供該股票資料"
    missing: list[str] = []
    if per is None:
        missing.append("本益比")
    if pbr is None:
        missing.append("股價淨值比")
    raw_text = str(raw_dividend_yield or "").strip()
    if not raw_text or raw_text in {"-", "--", "N/A", "null"}:
        missing.append("殖利率")
    elif dividend_yield is None:
        return "官方殖利率格式無法解析"
    if missing:
        return "官方欄位缺值：" + "、".join(missing)
    return ""


def monthly_revenue_note(
    revenue_row: dict[str, Any],
    revenue_period: str,
    monthly_revenue: float | None,
    monthly_revenue_mom: float | None,
    monthly_revenue_yoy: float | None,
    cumulative_revenue_yoy: float | None,
) -> str:
    if not revenue_row:
        return "官方月營收端點未提供該股票資料"
    missing: list[str] = []
    if not revenue_period:
        missing.append("月營收期別")
    if monthly_revenue is None:
        missing.append("月營收")
    if monthly_revenue_mom is None:
        missing.append("月營收月增率")
    if monthly_revenue_yoy is None:
        missing.append("月營收年增率")
    if cumulative_revenue_yoy is None:
        missing.append("累計營收年增率")
    if missing:
        return "官方欄位缺值：" + "、".join(missing)
    return ""


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


def latest_by_code(rows: Iterable[dict[str, Any]], code_keys: tuple[str, ...], period_keys: tuple[str, ...]) -> dict[str, dict[str, Any]]:
    indexed: dict[str, dict[str, Any]] = {}
    for row in rows:
        code = ""
        for key in code_keys:
            code = clean_code(row.get(key))
            if code:
                break
        if not code:
            continue
        current = indexed.get(code)
        if current is None:
            indexed[code] = row
            continue
        row_period = "".join(clean_code(row.get(key)) for key in period_keys)
        current_period = "".join(clean_code(current.get(key)) for key in period_keys)
        if row_period >= current_period:
            indexed[code] = row
    return indexed


def fundamentals_source(market: str) -> str:
    if market == "tse":
        return "TWSE: BWIBBU_ALL + t187ap05_L + t187ap14_L + t187ap17_L"
    if market == "otc":
        return "TPEx: tpex_mainboard_peratio_analysis + mopsfin_t187ap05_O + mopsfin_t187ap14_O"
    return ""


def valuation_source(market: str) -> str:
    return "TWSE BWIBBU_ALL" if market == "tse" else "TPEx tpex_mainboard_peratio_analysis"


def monthly_revenue_source(market: str) -> str:
    return "TWSE t187ap05_L" if market == "tse" else "TPEx mopsfin_t187ap05_O"


def financial_source(market: str) -> str:
    return "TWSE t187ap14_L / t187ap17_L" if market == "tse" else "TPEx mopsfin_t187ap14_O"


def fetch_fundamentals(
    stocks: list[StockInfo],
    skip: bool,
    statuses: list[SourceStatus],
) -> dict[tuple[str, str], dict[str, Any]]:
    if skip:
        statuses.append(SourceStatus(source="fundamentals", url="", status="skipped"))
        return {}

    markets = {stock.market for stock in stocks}
    result: dict[tuple[str, str], dict[str, Any]] = {}
    twse_valuation: dict[str, dict[str, Any]] = {}
    tpex_valuation: dict[str, dict[str, Any]] = {}
    twse_revenue: dict[str, dict[str, Any]] = {}
    tpex_revenue: dict[str, dict[str, Any]] = {}
    twse_eps: dict[str, dict[str, Any]] = {}
    tpex_eps: dict[str, dict[str, Any]] = {}
    twse_margin: dict[str, dict[str, Any]] = {}

    if "tse" in markets:
        twse_valuation = latest_by_code(fetch_json(TWSE_VALUATION_URL, "twse_valuation_metrics", statuses), ("Code",), ("Date",))
        twse_revenue = latest_by_code(fetch_json(TWSE_REVENUE_URL, "twse_monthly_revenue", statuses), ("公司代號",), ("資料年月", "出表日期"))
        twse_eps = latest_by_code(fetch_json(TWSE_EPS_URL, "twse_eps_income", statuses), ("公司代號",), ("年度", "季別", "出表日期"))
        twse_margin = latest_by_code(fetch_json(TWSE_MARGIN_URL, "twse_margin_ratios", statuses), ("公司代號",), ("年度", "季別", "出表日期"))
    if "otc" in markets:
        tpex_valuation = latest_by_code(
            fetch_json(TPEX_VALUATION_URL, "tpex_valuation_metrics", statuses),
            ("SecuritiesCompanyCode",),
            ("Date",),
        )
        tpex_revenue = latest_by_code(
            fetch_json(TPEX_REVENUE_URL, "tpex_monthly_revenue", statuses),
            ("公司代號", "SecuritiesCompanyCode"),
            ("資料年月", "出表日期", "Date"),
        )
        tpex_eps = latest_by_code(
            fetch_json(TPEX_EPS_URL, "tpex_eps_income", statuses),
            ("SecuritiesCompanyCode", "公司代號"),
            ("Year", "年度", "季別", "Date", "出表日期"),
        )

    for stock in stocks:
        if stock.market == "tse":
            valuation_row = twse_valuation.get(stock.stock_id, {})
            revenue_row = twse_revenue.get(stock.stock_id, {})
            eps_row = twse_eps.get(stock.stock_id, {})
            margin_row = twse_margin.get(stock.stock_id, {})
            valuation_metric_date = normalize_date(clean_code(valuation_row.get("Date")))
            per = to_float(valuation_row.get("PEratio"))
            pbr = to_float(valuation_row.get("PBratio"))
            raw_dividend_yield = valuation_row.get("DividendYield")
            dividend_yield = percent_value(raw_dividend_yield)
            revenue_period = normalize_month(clean_code(revenue_row.get("資料年月")))
            financial_period = normalize_period(eps_row.get("年度"), eps_row.get("季別"))
            financial_revenue = to_float(eps_row.get("營業收入"))
            operating_income = to_float(eps_row.get("營業利益"))
            net_income = to_float(eps_row.get("稅後淨利"))
            gross_margin = percent_value(margin_row.get("毛利率(%)(營業毛利)/(營業收入)"))
            operating_margin = percent_value(margin_row.get("營業利益率(%)(營業利益)/(營業收入)"))
            net_margin = percent_value(margin_row.get("稅後純益率(%)(稅後純益)/(營業收入)"))
        else:
            valuation_row = tpex_valuation.get(stock.stock_id, {})
            revenue_row = tpex_revenue.get(stock.stock_id, {})
            eps_row = tpex_eps.get(stock.stock_id, {})
            valuation_metric_date = normalize_date(clean_code(valuation_row.get("Date")))
            per = to_float(valuation_row.get("PriceEarningRatio"))
            pbr = to_float(valuation_row.get("PriceBookRatio"))
            raw_dividend_yield = valuation_row.get("YieldRatio")
            dividend_yield = percent_value(raw_dividend_yield)
            revenue_period = normalize_month(clean_code(revenue_row.get("資料年月")))
            financial_period = normalize_period(eps_row.get("Year") or eps_row.get("年度"), eps_row.get("季別"))
            financial_revenue = to_float(eps_row.get("營業收入"))
            operating_income = to_float(eps_row.get("營業利益"))
            net_income = to_float(eps_row.get("稅後淨利"))
            gross_margin = None
            operating_margin = operating_income / financial_revenue if operating_income is not None and financial_revenue not in {None, 0} else None
            net_margin = net_income / financial_revenue if net_income is not None and financial_revenue not in {None, 0} else None

        notes: list[str] = []
        monthly_revenue = to_float(revenue_row.get("營業收入-當月營收"))
        monthly_revenue_mom = percent_value(revenue_row.get("營業收入-上月比較增減(%)"))
        monthly_revenue_yoy = percent_value(revenue_row.get("營業收入-去年同月增減(%)"))
        cumulative_revenue_yoy = percent_value(revenue_row.get("累計營業收入-前期比較增減(%)"))
        if stock.market == "tse":
            notes.append("上市毛利率來自 TWSE t187ap17_L")
        if stock.market == "otc":
            notes.append("上櫃毛利率目前未找到穩定 TPEx OpenAPI 欄位，暫留空；請看營業利益率與稅後淨利率")
        available_groups = sum(bool(row) for row in (valuation_row, revenue_row, eps_row))
        overall_status = "complete" if available_groups == 3 else "partial" if available_groups else "missing"
        result[(stock.market, stock.stock_id)] = {
            "valuation_metric_date": valuation_metric_date,
            "valuation_source": valuation_source(stock.market),
            "per": per,
            "pbr": pbr,
            "dividend_yield_pct": dividend_yield,
            "valuation_note": valuation_note(valuation_row, per, pbr, raw_dividend_yield, dividend_yield),
            "revenue_period": revenue_period,
            "monthly_revenue_source": monthly_revenue_source(stock.market),
            "monthly_revenue_note": monthly_revenue_note(
                revenue_row,
                revenue_period,
                monthly_revenue,
                monthly_revenue_mom,
                monthly_revenue_yoy,
                cumulative_revenue_yoy,
            ),
            "monthly_revenue": monthly_revenue,
            "monthly_revenue_mom_pct": monthly_revenue_mom,
            "monthly_revenue_yoy_pct": monthly_revenue_yoy,
            "cumulative_revenue_yoy_pct": cumulative_revenue_yoy,
            "financial_period": financial_period,
            "financial_source": financial_source(stock.market),
            "eps": to_float(eps_row.get("基本每股盈餘(元)") or eps_row.get("基本每股盈餘")),
            "financial_revenue": financial_revenue,
            "operating_income": operating_income,
            "net_income": net_income,
            "gross_margin_pct": gross_margin,
            "operating_margin_pct": operating_margin,
            "net_margin_pct": net_margin,
            "fundamentals_source": fundamentals_source(stock.market),
            "fundamentals_status": overall_status,
            "fundamentals_note": "；".join(notes),
        }
    return result


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


def limit_stocks(stocks: list[StockInfo], limit: int) -> list[StockInfo]:
    if limit <= 0 or len(stocks) <= limit:
        return stocks
    markets = {stock.market for stock in stocks}
    if markets == {"tse", "otc"}:
        tse_limit = limit // 2
        otc_limit = limit - tse_limit
        tse = [stock for stock in stocks if stock.market == "tse"][:tse_limit]
        otc = [stock for stock in stocks if stock.market == "otc"][:otc_limit]
        return sorted(tse + otc, key=lambda stock: (stock.market, stock.stock_id))
    return stocks[:limit]


def detect_cnyes_source(page_text: str) -> str:
    if "factSetEstimate" in page_text or re.search(r"FactSet", page_text, flags=re.IGNORECASE):
        return CNYES_FACTSET_SOURCE
    return CNYES_UNSPECIFIED_SOURCE


def extract_cnyes_target(stock_id: str, timeout: int) -> dict[str, Any]:
    url = CNYES_STOCK_URL.format(stock_id=stock_id)
    text = request_text(url, timeout=timeout, referer="https://www.cnyes.com/")
    source = detect_cnyes_source(text)
    match = re.search(r'"targetValuation"\s*:\s*(\{.*?\}|null)\s*,', text)
    if not match:
        return {"cnyes_status": "no_target_valuation", "cnyes_url": url, "cnyes_source": CNYES_NO_TARGET_SOURCE}
    if match.group(1) == "null":
        return {"cnyes_status": "no_target_valuation", "cnyes_url": url, "cnyes_source": CNYES_NO_TARGET_SOURCE}
    try:
        data = json.loads(match.group(1))
    except json.JSONDecodeError as exc:
        return {"cnyes_status": "parse_error", "cnyes_url": url, "cnyes_source": source, "cnyes_error": str(exc)}
    return {
        "cnyes_url": url,
        "cnyes_status": "ok",
        "cnyes_source": source,
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
        "cnyes_source": CNYES_NOT_FETCHED_SOURCE,
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
                        "cnyes_source": CNYES_NOT_FETCHED_SOURCE,
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
                "cnyes_source": CNYES_NOT_FETCHED_SOURCE,
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
    fundamentals: dict[tuple[str, str], dict[str, Any]],
    args: argparse.Namespace,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for stock in stocks:
        close_row = closes.get((stock.market, stock.stock_id), {})
        target_row = targets.get(stock.stock_id, {})
        fundamental_row = fundamentals.get((stock.market, stock.stock_id), {})
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
            "cnyes_source": target_row.get(
                "cnyes_source",
                CNYES_NOT_FETCHED_SOURCE if args.skip_cnyes else CNYES_NO_TARGET_SOURCE,
            ),
            "valuation_metric_date": fundamental_row.get("valuation_metric_date", ""),
            "valuation_source": fundamental_row.get("valuation_source", ""),
            "per": fundamental_row.get("per"),
            "pbr": fundamental_row.get("pbr"),
            "dividend_yield_pct": fundamental_row.get("dividend_yield_pct"),
            "valuation_note": fundamental_row.get("valuation_note", ""),
            "revenue_period": fundamental_row.get("revenue_period", ""),
            "monthly_revenue_source": fundamental_row.get("monthly_revenue_source", ""),
            "monthly_revenue_note": fundamental_row.get("monthly_revenue_note", ""),
            "monthly_revenue": fundamental_row.get("monthly_revenue"),
            "monthly_revenue_mom_pct": fundamental_row.get("monthly_revenue_mom_pct"),
            "monthly_revenue_yoy_pct": fundamental_row.get("monthly_revenue_yoy_pct"),
            "cumulative_revenue_yoy_pct": fundamental_row.get("cumulative_revenue_yoy_pct"),
            "financial_period": fundamental_row.get("financial_period", ""),
            "financial_source": fundamental_row.get("financial_source", ""),
            "eps": fundamental_row.get("eps"),
            "financial_revenue": fundamental_row.get("financial_revenue"),
            "operating_income": fundamental_row.get("operating_income"),
            "net_income": fundamental_row.get("net_income"),
            "gross_margin_pct": fundamental_row.get("gross_margin_pct"),
            "operating_margin_pct": fundamental_row.get("operating_margin_pct"),
            "net_margin_pct": fundamental_row.get("net_margin_pct"),
            "fundamentals_source": fundamental_row.get("fundamentals_source", ""),
            "fundamentals_status": fundamental_row.get("fundamentals_status", "skipped" if args.skip_fundamentals else "missing"),
            "fundamentals_note": fundamental_row.get("fundamentals_note", ""),
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
        "cnyes_source",
        "cnyes_url",
        "valuation_metric_date",
        "valuation_source",
        "valuation_note",
        "per",
        "pbr",
        "dividend_yield_pct",
        "revenue_period",
        "monthly_revenue_source",
        "monthly_revenue_note",
        "monthly_revenue",
        "monthly_revenue_mom_pct",
        "monthly_revenue_yoy_pct",
        "cumulative_revenue_yoy_pct",
        "financial_period",
        "financial_source",
        "eps",
        "financial_revenue",
        "operating_income",
        "net_income",
        "gross_margin_pct",
        "operating_margin_pct",
        "net_margin_pct",
        "fundamentals_status",
        "fundamentals_note",
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
        "cnyes_source",
        "cnyes_url",
        "valuation_metric_date",
        "valuation_source",
        "valuation_note",
        "per",
        "pbr",
        "dividend_yield_pct",
        "revenue_period",
        "monthly_revenue_source",
        "monthly_revenue_note",
        "monthly_revenue_yoy_pct",
        "cumulative_revenue_yoy_pct",
        "financial_period",
        "financial_source",
        "eps",
        "operating_margin_pct",
        "net_margin_pct",
        "fundamentals_status",
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
        "cnyes_source",
        "cnyes_url",
        "valuation_metric_date",
        "valuation_source",
        "valuation_note",
        "per",
        "pbr",
        "dividend_yield_pct",
        "revenue_period",
        "monthly_revenue_source",
        "monthly_revenue_note",
        "monthly_revenue_yoy_pct",
        "cumulative_revenue_yoy_pct",
        "financial_period",
        "financial_source",
        "eps",
        "operating_margin_pct",
        "net_margin_pct",
        "fundamentals_status",
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
        return "TWSE"
    if market == "otc":
        return "TPEx"
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


COLUMN_LABEL_PERIODS: dict[str, str] = {}


def most_common_value(rows: list[dict[str, Any]], column: str) -> str:
    counts: dict[str, int] = {}
    for row in rows:
        value = str(row.get(column) or "").strip()
        if value:
            counts[value] = counts.get(value, 0) + 1
    if not counts:
        return ""
    return sorted(counts.items(), key=lambda item: (-item[1], item[0]))[0][0]


def set_column_label_periods(rows: list[dict[str, Any]]) -> None:
    COLUMN_LABEL_PERIODS.clear()
    valuation_date = most_common_value(rows, "valuation_metric_date")
    revenue_period = most_common_value(rows, "revenue_period")
    financial_period = most_common_value(rows, "financial_period")
    for column in ("per", "pbr", "dividend_yield_pct"):
        if valuation_date:
            COLUMN_LABEL_PERIODS[column] = valuation_date
    for column in (
        "monthly_revenue",
        "monthly_revenue_mom_pct",
        "monthly_revenue_yoy_pct",
        "cumulative_revenue_yoy_pct",
    ):
        if revenue_period:
            COLUMN_LABEL_PERIODS[column] = revenue_period
    for column in (
        "eps",
        "financial_revenue",
        "operating_income",
        "net_income",
        "gross_margin_pct",
        "operating_margin_pct",
        "net_margin_pct",
    ):
        if financial_period:
            COLUMN_LABEL_PERIODS[column] = financial_period


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
        "cnyes_source": "鉅亨資料來源",
        "valuation_metric_date": "估值資料日",
        "valuation_source": "估值資料來源",
        "valuation_note": "估值資料備註",
        "per": "本益比",
        "pbr": "股價淨值比",
        "dividend_yield_pct": "殖利率",
        "revenue_period": "月營收期別",
        "monthly_revenue_source": "月營收資料來源",
        "monthly_revenue_note": "月營收資料備註",
        "monthly_revenue": "月營收(仟元)",
        "monthly_revenue_mom_pct": "月營收月增率",
        "monthly_revenue_yoy_pct": "月營收年增率",
        "cumulative_revenue_yoy_pct": "累計營收年增率",
        "financial_period": "財報期別",
        "financial_source": "財報資料來源",
        "eps": "EPS",
        "financial_revenue": "財報營業收入(仟元)",
        "operating_income": "營業利益(仟元)",
        "net_income": "稅後淨利(仟元)",
        "gross_margin_pct": "毛利率",
        "operating_margin_pct": "營業利益率",
        "net_margin_pct": "稅後淨利率",
        "fundamentals_status": "基本面整體狀態",
        "fundamentals_source": "基本面整體資料來源",
        "fundamentals_note": "基本面整體備註",
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
    label = labels.get(column, column)
    period = COLUMN_LABEL_PERIODS.get(column)
    if period:
        return f"{label}\n({period})"
    return label


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
        "exchange_source": "交易所資料來源；股票表內精簡顯示上市 TWSE 或上櫃 TPEx，完整端點請看相鄰的交易所資料說明。",
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
        "valuation_signal": "估值訊號；Excel 顯示中文，例如低估、高估、中性、評價過舊、低信心、缺少目標價。程式內部仍用英文代碼。",
        "confidence_note": "可信度註記；Excel 顯示中文，例如資料新鮮，預估 33 家、評價過舊 120 天、預估家數僅 2 家、缺少鉅亨目標價。",
        "cnyes_status": "鉅亨抓取狀態；Excel 顯示中文，例如成功讀取鉅亨資料、鉅亨無目標價、抓取失敗、解析失敗、錯誤率過高已停止。",
        "cnyes_source": CNYES_DATA_SOURCE_NOTE,
        "cnyes_url": "鉅亨個股頁網址，格式為 https://www.cnyes.com/twstock/{股票代號}；方便人工開啟來源頁，回查該股票頁內嵌的 targetValuation 資料。",
        "cnyes_attempts": "鉅亨個股頁嘗試抓取次數。",
        "valuation_metric_date": "本益比、股價淨值比、殖利率使用的資料日期；上市來自 TWSE BWIBBU_ALL.Date，上櫃來自 TPEx tpex_mainboard_peratio_analysis.Date。",
        "valuation_source": "估值資料來源；上市為 TWSE BWIBBU_ALL，上櫃為 TPEx tpex_mainboard_peratio_analysis。",
        "valuation_note": "估值資料備註；區分官方估值端點未提供該股票，以及官方列存在但本益比、股價淨值比或殖利率欄位缺值。",
        "per": "本益比；上市來自 TWSE BWIBBU_ALL.PEratio，上櫃來自 TPEx tpex_mainboard_peratio_analysis.PriceEarningRatio。這不是財報期別欄位，而是交易所估值資料日的指標。",
        "pbr": "股價淨值比；上市來自 TWSE BWIBBU_ALL.PBratio，上櫃來自 TPEx tpex_mainboard_peratio_analysis.PriceBookRatio。",
        "dividend_yield_pct": "殖利率；上市來自 TWSE BWIBBU_ALL.DividendYield，上櫃來自 TPEx tpex_mainboard_peratio_analysis.YieldRatio。來源原值是百分比數字，Excel 轉成百分比格式顯示。",
        "revenue_period": "月營收期別；來自 TWSE t187ap05_L 或 TPEx mopsfin_t187ap05_O 的資料年月，民國年月會轉成西元 YYYY-MM。",
        "monthly_revenue_source": "月營收資料來源；上市為 TWSE t187ap05_L，上櫃為 TPEx mopsfin_t187ap05_O。",
        "monthly_revenue_note": "月營收逐欄狀態；區分官方月營收端點未提供該股票資料，以及官方列存在但月營收期別、月營收、月增率、年增率或累計年增率缺值。估值資料與月營收資料來自不同端點，不能用估值有值推定月營收也應有值。",
        "monthly_revenue": "月營收，來源為 t187ap05_L/O 的營業收入-當月營收；單位依交易所公開資料欄位，以仟元呈現。",
        "monthly_revenue_mom_pct": "月營收月增率，來源為 t187ap05_L/O 的營業收入-上月比較增減(%)。",
        "monthly_revenue_yoy_pct": "月營收年增率，來源為 t187ap05_L/O 的營業收入-去年同月增減(%)。",
        "cumulative_revenue_yoy_pct": "累計營收年增率，來源為 t187ap05_L/O 的累計營業收入-前期比較增減(%)。",
        "financial_period": "EPS、財報營業收入、營業利益、稅後淨利與利潤率使用的財報期別；由年度/Year 加季別轉成 YYYYQn。",
        "financial_source": "財報資料來源；上市為 TWSE t187ap14_L/t187ap17_L，上櫃為 TPEx mopsfin_t187ap14_O。",
        "eps": "每股盈餘；上市來自 TWSE t187ap14_L.基本每股盈餘(元)，上櫃來自 TPEx mopsfin_t187ap14_O.基本每股盈餘。",
        "financial_revenue": "財報營業收入；上市來自 TWSE t187ap14_L，上櫃來自 TPEx mopsfin_t187ap14_O，單位以公開資料欄位常用仟元呈現。",
        "operating_income": "營業利益；上市來自 TWSE t187ap14_L，上櫃來自 TPEx mopsfin_t187ap14_O，單位以仟元呈現。",
        "net_income": "稅後淨利；上市來自 TWSE t187ap14_L，上櫃來自 TPEx mopsfin_t187ap14_O，單位以仟元呈現。",
        "gross_margin_pct": "毛利率；上市來自 TWSE t187ap17_L.毛利率(%)(營業毛利)/(營業收入)。上櫃目前未找到穩定 TPEx OpenAPI 對應欄位，暫留空。",
        "operating_margin_pct": "營業利益率；上市優先使用 TWSE t187ap17_L.營業利益率(%)，上櫃以 mopsfin_t187ap14_O 的營業利益 / 營業收入計算。",
        "net_margin_pct": "稅後淨利率；上市優先使用 TWSE t187ap17_L.稅後純益率(%)，上櫃以 mopsfin_t187ap14_O 的稅後淨利 / 營業收入計算。",
        "fundamentals_status": "基本面整體狀態；綜合估值、月營收、財報三組官方資料，顯示三組齊全、部分缺少、三組皆缺少或略過。個別欄位缺值請看各組資料備註。",
        "fundamentals_source": "基本面資料來源；列出本列基本面欄位使用的 TWSE/TPEx 端點。",
        "fundamentals_note": "基本面整體備註；股票列只保留會直接影響該列判讀的資料限制，例如上櫃毛利率暫缺公開穩定端點。ROE 尚未納入的共同限制集中放在使用說明，不在每支股票重複顯示。",
    }
    rows = []
    for key, value in definitions.items():
        label = column_label(key).replace("\n", "")
        field = f"{label} ({key})" if label != key else key
        rows.append({"field": field, "description": value})
    return rows


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
            "section": "資料來源",
            "item": "鉅亨資料來源",
            "description": CNYES_DATA_SOURCE_NOTE,
        },
        {
            "section": "資料來源",
            "item": "本益比/股價淨值比/殖利率",
            "description": "這三欄不是鉅亨資料，也不是財報期別資料；上市來自 TWSE BWIBBU_ALL，上櫃來自 TPEx tpex_mainboard_peratio_analysis。估值資料日後會列出估值資料來源與備註；個別欄位空白不等於整批爬取失敗。",
        },
        {
            "section": "資料來源",
            "item": "估值與月營收資料不同",
            "description": "估值資料日、本益比、股價淨值比與殖利率使用 BWIBBU_ALL / tpex_mainboard_peratio_analysis；月營收期別、月營收及成長率使用 t187ap05_L / mopsfin_t187ap05_O。兩者是不同官方端點，更新頻率與涵蓋股票可能不同。",
        },
        {
            "section": "資料來源",
            "item": "月營收資料",
            "description": "月營收、月增率、年增率與累計年增率來自 TWSE t187ap05_L 或 TPEx mopsfin_t187ap05_O，表頭括號顯示月營收期別，例如 2026-05。",
        },
        {
            "section": "資料來源",
            "item": "財報資料",
            "description": "EPS、財報營業收入、營業利益、稅後淨利與利潤率依財報期別呈現；上市使用 TWSE t187ap14_L/t187ap17_L，上櫃使用 TPEx mopsfin_t187ap14_O，表頭括號顯示例如 2026Q1。",
        },
        {
            "section": "資料來源",
            "item": "基本面限制",
            "description": "上櫃毛利率目前沒有確認到穩定 TPEx OpenAPI 對應欄位，先留空；ROE 也暫未納入，因公開端點尚未確認穩定權益資料來源。",
        },
        {
            "section": "資料品質",
            "item": "基本面整體狀態",
            "description": "只彙總估值、月營收、財報三組官方資料是否有該股票資料列；三組都有為齊全，只有部分有為部分缺少。個別欄位為何空白仍請看估值或月營收資料備註。",
        },
        {
            "section": "資料來源",
            "item": "毛利率註記",
            "description": "毛利率欄位上市有資料，來源是 TWSE t187ap17_L；上櫃目前空白，不會用營業利益率代替，以免混淆兩種不同指標。",
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
            "item": "交易所估值指標",
            "description": "上市一次下載 TWSE BWIBBU_ALL；上櫃一次下載 TPEx tpex_mainboard_peratio_analysis。這批資料提供本益比、股價淨值比與殖利率，依股票代號 join 到同一列。",
        },
        {
            "section": "爬取方式",
            "item": "交易所月營收",
            "description": "上市一次下載 TWSE t187ap05_L；上櫃一次下載 TPEx mopsfin_t187ap05_O。這批資料提供月營收、月增率、年增率與累計年增率。",
        },
        {
            "section": "爬取方式",
            "item": "交易所財報資料",
            "description": "上市一次下載 TWSE t187ap14_L 與 t187ap17_L；上櫃一次下載 TPEx mopsfin_t187ap14_O。這批資料提供 EPS、損益欄位與部分利潤率。",
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
            "section": "爬取方式",
            "item": "人工回查平均目標價",
            "description": "Excel 的「平均目標價」不是程式自行用最高/最低目標價平均，而是鉅亨個股頁內嵌 targetValuation.feMean。人工開啟鉅亨個股頁時可先切到「預估」tab；頁面 UI 不一定逐字顯示「平均目標價」，但程式抓的是同頁內嵌 JSON 的 feMean。",
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
            "description": "Excel 顯示中文：低估、高估、中性、評價過舊、低信心、缺少目標價；程式內部仍保留英文代碼以維持排序與設定穩定。",
        },
        {
            "section": "判斷",
            "item": "為何沒進低估/高估",
            "description": "程式會先檢查資料品質：缺目標價、評價日期超過過舊天數門檻、或預估家數少於最低家數門檻時，會顯示為缺少目標價、評價過舊或低信心，不放進主要低估或高估清單。",
        },
        {
            "section": "判斷",
            "item": "判斷順序",
            "description": "順序是：鉅亨狀態與目標價是否存在 -> 評價是否過舊 -> 預估家數是否足夠 -> 才用平均目標價潛在漲跌幅判斷低估、高估或中性。",
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
            "description": "Excel 顯示中文，例如資料新鮮，預估 33 家、評價過舊 120 天、預估家數僅 2 家、缺少鉅亨目標價，用來快速看資料新鮮度與預估家數。",
        },
        {
            "section": "欄位說明",
            "item": "逐欄查詢",
            "description": "每個中文表頭的來源、公式、單位與內部 key 集中放在「欄位說明」分頁；「使用說明」只保留閱讀順序、資料來源、爬取方式與判斷邏輯。",
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


def worksheet_hyperlinks(rows: list[dict[str, Any]], columns: list[str]) -> list[tuple[str, str, str]]:
    if "cnyes_url" not in columns:
        return []
    col_idx = columns.index("cnyes_url") + 1
    links = []
    for data_idx, row in enumerate(rows, start=2):
        url = str(row.get("cnyes_url") or "").strip()
        if not url:
            continue
        links.append((cell_ref(data_idx, col_idx), f"rId{len(links) + 1}", url))
    return links


def worksheet_rels_xml(hyperlinks: list[tuple[str, str, str]]) -> str:
    rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for _, rel_id, url in hyperlinks:
        rels.append(
            f'<Relationship Id="{xml_text(rel_id)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{xml_text(url)}" TargetMode="External"/>'
        )
    rels.append("</Relationships>")
    return "".join(rels)


VALUATION_SIGNAL_LABELS = {
    "undervalued": "低估",
    "overvalued": "高估",
    "neutral": "中性",
    "stale": "評價過舊",
    "low_confidence": "低信心",
    "missing_target": "缺少目標價",
}

CNYES_STATUS_LABELS = {
    "ok": "成功讀取鉅亨資料",
    "no_target_valuation": "鉅亨無目標價",
    "parse_error": "解析失敗",
    "http_error": "抓取失敗",
    "skipped_limit": "測試限制未抓",
    "skipped_error_threshold": "錯誤率過高已停止",
    "skipped": "略過鉅亨",
    "missing": "缺少鉅亨資料",
}

FUNDAMENTALS_STATUS_LABELS = {
    "ok": "基本面至少有部分資料",
    "complete": "估值、月營收、財報三組齊全",
    "partial": "基本面部分資料缺少",
    "missing": "估值、月營收、財報三組皆缺少",
    "skipped": "略過基本面抓取",
}


def confidence_note_label(value: Any) -> str:
    text = str(value or "")
    if text == "missing_cnyes_target":
        return "缺少鉅亨目標價"
    match = re.fullmatch(r"stale_(\d+)_days", text)
    if match:
        return f"評價過舊 {match.group(1)} 天"
    match = re.fullmatch(r"only_(\d+)_estimates", text)
    if match:
        return f"預估家數僅 {match.group(1)} 家"
    match = re.fullmatch(r"fresh_(\d+)_estimates", text)
    if match:
        return f"資料新鮮，預估 {match.group(1)} 家"
    return text


def display_cell_value(column: str, value: Any) -> Any:
    if column == "market":
        return market_name(str(value or ""))
    if column == "valuation_signal":
        return VALUATION_SIGNAL_LABELS.get(str(value or ""), value)
    if column == "confidence_note":
        return confidence_note_label(value)
    if column == "cnyes_status":
        return CNYES_STATUS_LABELS.get(str(value or ""), value)
    if column == "fundamentals_status":
        return FUNDAMENTALS_STATUS_LABELS.get(str(value or ""), value)
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
    hyperlinks = worksheet_hyperlinks(rows, columns)
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
        width = min(26, max(10, len(label) + 4))
        if column in {"cnyes_url", "company_name", "cnyes_error", "confidence_note"}:
            width = 24
        if column == "cnyes_source":
            width = 18
        if column == "exchange_source":
            width = 12
        if column == "section":
            width = 14
        if column == "item":
            width = 28
        if column == "description":
            width = 72
        if column == "field":
            width = 36
        if column == "generated_at":
            width = 20
        if column == "source":
            width = 28
        if column == "status":
            width = 12
        if column == "rows":
            width = 12
        if column == "data_date":
            width = 14
        if column == "url":
            width = 60
        if column == "message":
            width = 36
        if column in {"listing_date", "close_date"}:
            width = 14
        if column == "valuation_metric_date":
            width = 14
        if column == "revenue_period":
            width = 14
        if column == "target_date":
            width = 14
        if column in {"target_high", "target_low", "target_mean", "target_median"}:
            width = 14
        if column in {"num_est", "fe_up", "fe_down", "currency"}:
            width = 12
        if column == "fe_stddev":
            width = 16
        if column == "cnyes_last":
            width = 14
        if column in {
            "valuation_metric_date",
            "per",
            "pbr",
            "dividend_yield_pct",
            "revenue_period",
            "monthly_revenue_mom_pct",
            "monthly_revenue_yoy_pct",
            "cumulative_revenue_yoy_pct",
            "financial_period",
            "eps",
            "gross_margin_pct",
            "operating_margin_pct",
            "net_margin_pct",
            "fundamentals_status",
        }:
            width = min(width, 18)
        if column == "fundamentals_status":
            width = 20
        if column in {"valuation_note", "monthly_revenue_note"}:
            width = 28
        if column in {"valuation_source", "monthly_revenue_source", "financial_source", "fundamentals_source", "fundamentals_note", "exchange_note"}:
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
        "per",
        "pbr",
        "eps",
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
        "monthly_revenue",
        "financial_revenue",
        "operating_income",
        "net_income",
    }
    number_cols = decimal_cols.union(integer_cols)
    for row_idx, row_values in enumerate(table, start=1):
        row_attrs = f' r="{row_idx}"'
        if row_idx == 1:
            row_attrs += ' ht="40" customHeight="1"'
        parts.append(f'<row{row_attrs}>')
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
            elif row_idx > 1 and column == "cnyes_url" and value:
                style = 5

            if value is None or value == "":
                continue
            elif row_idx > 1 and isinstance(value, (int, float)) and column in percent_cols.union(number_cols):
                parts.append(f'<c r="{ref}" s="{style}"><v>{value}</v></c>')
            else:
                parts.append(f'<c r="{ref}" s="{style}" t="s"><v>{shared_strings[str(value)]}</v></c>')
        parts.append("</row>")
    parts.append("</sheetData>")
    if hyperlinks:
        parts.append("<hyperlinks>")
        for ref, rel_id, _ in hyperlinks:
            parts.append(f'<hyperlink ref="{xml_text(ref)}" r:id="{xml_text(rel_id)}"/>')
        parts.append("</hyperlinks>")
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
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><u/><color rgb="FF0563C1"/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
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


def write_xlsx(path: Path, sheets: list[tuple[str, list[dict[str, Any]], list[str]]], active_sheet_index: int | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    shared_strings, shared_string_count = collect_shared_strings(sheets)
    sheet_names = [name for name, _, _ in sheets]
    if active_sheet_index is None:
        active_sheet_index = sheet_names.index("低估清單") + 1 if "低估清單" in sheet_names else 1
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
            hyperlinks = worksheet_hyperlinks(rows, columns)
            zf.writestr(
                posixpath.join("xl", "worksheets", f"sheet{idx}.xml"),
                xlsx_sheet_xml(name, rows, columns, shared_strings, tab_selected=idx == active_sheet_index),
            )
            if hyperlinks:
                zf.writestr(
                    posixpath.join("xl", "worksheets", "_rels", f"sheet{idx}.xml.rels"),
                    worksheet_rels_xml(hyperlinks),
                )


def build_workbook_rows(rows: list[dict[str, Any]], statuses: list[SourceStatus], args: argparse.Namespace) -> list[tuple[str, list[dict[str, Any]], list[str]]]:
    set_column_label_periods(rows)
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
        ("\u6b04\u4f4d\u8aaa\u660e", dictionary_rows(), ["field", "description"]),
        ("\u4f4e\u4f30\u6e05\u55ae", undervalued, columns),
        ("\u9ad8\u4f30\u6e05\u55ae", overvalued, columns),
        ("\u5168\u90e8\u80a1\u7968", rows, columns),
        ("\u904e\u820a\u4f4e\u4fe1\u5fc3", stale_low, stale_column_order()),
        ("\u6293\u53d6\u72c0\u614b", status_rows(statuses, rows), ["generated_at", "source", "status", "rows", "data_date", "url", "message"]),
    ]


def build_lite_workbook_rows(rows: list[dict[str, Any]], statuses: list[SourceStatus], args: argparse.Namespace) -> list[tuple[str, list[dict[str, Any]], list[str]]]:
    set_column_label_periods(rows)
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
        ("\u6b04\u4f4d\u8aaa\u660e", dictionary_rows(), ["field", "description"]),
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
    parser.add_argument("--stock-limit", type=int, default=0, help="Limit output rows for a small balanced test run.")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--undervalued-threshold", type=float, default=env_float(env, "TW_STOCK_UNDERVALUED_THRESHOLD"))
    parser.add_argument("--overvalued-threshold", type=float, default=env_float(env, "TW_STOCK_OVERVALUED_THRESHOLD"))
    parser.add_argument("--stale-days", type=int, default=env_int(env, "TW_STOCK_STALE_DAYS"))
    parser.add_argument("--min-estimates", type=int, default=env_int(env, "TW_STOCK_MIN_ESTIMATES"))
    parser.add_argument("--excel-output", choices=["full", "lite", "both"], default=env_value(env, "TW_STOCK_EXCEL_OUTPUT"))
    parser.add_argument("--skip-cnyes", action="store_true")
    parser.add_argument("--skip-fundamentals", action="store_true")
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
    stocks = limit_stocks(stocks, args.stock_limit)
    if not stocks:
        print("No stocks found for the selected universe.", file=sys.stderr)
        return 2

    fundamentals = fetch_fundamentals(stocks, args.skip_fundamentals, statuses)
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
    rows = build_rows(stocks, closes, targets, fundamentals, args)
    date_part = report_date(rows).replace("-", "")
    generated_time_part = datetime.now().strftime("%H%M%S")
    suffix = ""
    if args.universe == "watchlist":
        suffix += "_watchlist"
    if args.skip_cnyes:
        suffix += "_no_cnyes"
    if args.cnyes_limit > 0 and not args.skip_cnyes:
        suffix += f"_cnyes_limit{args.cnyes_limit}"
    output_base = Path(args.output_dir) / f"tw_valuation_gap_{date_part}_{generated_time_part}{suffix}.xlsx"
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
