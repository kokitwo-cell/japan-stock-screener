"""
日本株スクリーナー - データ更新スクリプト
GitHub Actions から週次で実行され、data/*.json を更新する。

ローカル実行:
  pip install yfinance requests beautifulsoup4 pandas openpyxl xlrd
  python scripts/update_data.py

環境変数:
  JQUANTS_REFRESH_TOKEN  - J-Quants API キー（オプション、業種名取得用）
  FETCH_LIMIT            - 取得する銘柄数の上限（オプション、デバッグ用）
  ENRICH_IRBANK          - "1" で ir-bank からの長期業績補完を実行
  UPDATE_PRICES_ONLY     - "1" で株価のみ更新
"""

import os
import sys
import json
import time
import threading
import csv
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone

# 東証の営業時間判定用。日本にサマータイムは無いので固定オフセットで十分。
JST = timezone(timedelta(hours=9))

import yfinance as yf
import requests

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
os.makedirs(DATA_DIR, exist_ok=True)

CACHE_FILE         = os.path.join(DATA_DIR, "stock_cache.json")
TSE_CODES_FILE     = os.path.join(DATA_DIR, "tse_codes.json")
JQUANTS_INFO_FILE  = os.path.join(DATA_DIR, "jquants_info.json")
PORTFOLIO_FILE     = os.path.join(DATA_DIR, "portfolio.json")

CACHE_EXPIRE_HOURS = 168
MAX_WORKERS        = 3
FETCH_DELAY        = 0.5
RETRY_WAIT_SEC     = 65

JPX_XLS_URL = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"

# ============================================================
#  TSE コード読み込み
# ============================================================
def fetch_tse_codes_from_jpx():
    print(f"JPX XLSダウンロード: {JPX_XLS_URL}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }
    try:
        r = requests.get(JPX_XLS_URL, headers=headers, timeout=30)
        r.raise_for_status()
        tmp = os.path.join(DATA_DIR, "_data_j.xls")
        with open(tmp, "wb") as f:
            f.write(r.content)
        codes = parse_jpx_xlsx(tmp)
        try: os.remove(tmp)
        except: pass
        return codes
    except Exception as e:
        print(f"JPX取得失敗: {e}")
        return None


def _parse_xls(filepath):
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)

    header_idx = None
    for i in range(min(10, ws.nrows)):
        row = [str(ws.cell_value(i, j)).strip() for j in range(ws.ncols)]
        if "コード" in row or "銘柄コード" in row:
            header_idx = i
            break

    if header_idx is None:
        return None

    headers = [str(ws.cell_value(header_idx, j)).strip().lstrip("﻿")
               for j in range(ws.ncols)]
    print(f"XLS列名: {headers[:6]}")

    def col(name, *aliases):
        for n in (name,) + aliases:
            if n in headers:
                return headers.index(n)
        return None

    code_col   = col("コード", "銘柄コード")
    name_col   = col("銘柄名")
    market_col = col("市場・商品区分", "市場区分")

    if code_col is None:
        return None

    codes = []
    for i in range(header_idx + 1, ws.nrows):
        try:
            raw_code = ws.cell_value(i, code_col)
            if isinstance(raw_code, float):
                raw_code = int(raw_code)
            code = str(raw_code).strip().zfill(4)
            if not (len(code) == 4 and code.isdigit() and code != "0000"):
                continue
            name   = str(ws.cell_value(i, name_col) or "").strip()   if name_col   is not None else ""
            market = str(ws.cell_value(i, market_col) or "").strip() if market_col is not None else ""
            if market and not any(x in market for x in ["プライム", "スタンダード", "グロース"]):
                continue
            codes.append({"code": code, "name": name, "market": market})
        except Exception:
            continue

    print(f"XLSから {len(codes)} 銘柄取得")
    return codes if codes else None


def parse_jpx_xlsx(filepath):
    if HAS_XLRD:
        try:
            result = _parse_xls(filepath)
            if result:
                return result
        except Exception as e:
            print(f"XLSパース失敗: {e}")

    if not HAS_OPENPYXL:
        return None
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None

        header_idx = None
        for i, row in enumerate(rows[:10]):
            row_str = [str(c or "").strip() for c in row]
            if "コード" in row_str or "銘柄コード" in row_str:
                header_idx = i
                break

        if header_idx is None:
            return None

        headers = [str(c or "").strip().lstrip("﻿") for c in rows[header_idx]]

        def col(name, *aliases):
            for n in (name,) + aliases:
                if n in headers:
                    return headers.index(n)
            return None

        code_col   = col("コード", "銘柄コード")
        name_col   = col("銘柄名")
        market_col = col("市場・商品区分", "市場区分")

        if code_col is None:
            return None

        codes = []
        for row in rows[header_idx + 1:]:
            try:
                code = str(row[code_col] or "").strip().zfill(4)
                if not (len(code) == 4 and code.isdigit() and code != "0000"):
                    continue
                name   = str(row[name_col] or "").strip()   if name_col   is not None else ""
                market = str(row[market_col] or "").strip() if market_col is not None else ""
                if market and not any(x in market for x in ["プライム", "スタンダード", "グロース"]):
                    continue
                codes.append({"code": code, "name": name, "market": market})
            except Exception:
                continue
        return codes if codes else None
    except Exception as e:
        print(f"XLSX読み込みエラー: {e}")
        return None


def load_tse_codes():
    """data/tse_codes.json があれば優先、なければ JPX から取得"""
    if os.path.exists(TSE_CODES_FILE):
        try:
            with open(TSE_CODES_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if data and len(data) > 100:
                print(f"既存tse_codes.json読み込み: {len(data)}社")
                # 週次更新時はリフレッシュ
                age = time.time() - os.path.getmtime(TSE_CODES_FILE)
                if age < 7 * 24 * 3600:
                    return data
                print("古いので JPX から再取得")
        except Exception:
            pass

    codes = fetch_tse_codes_from_jpx()
    if codes:
        with open(TSE_CODES_FILE, "w", encoding="utf-8") as f:
            json.dump(codes, f, ensure_ascii=False)
        print(f"tse_codes.json 保存: {len(codes)}社")
    return codes


# ============================================================
#  キャッシュ
# ============================================================
_cache = {}
_cache_lock = threading.Lock()
_jquants_info = {}

def get_cache():
    with _cache_lock:
        return dict(_cache)

def set_cache(data):
    with _cache_lock:
        _cache.clear()
        _cache.update(data)

def load_cache():
    if not os.path.exists(CACHE_FILE):
        return {}
    try:
        with open(CACHE_FILE, encoding="utf-8") as f:
            cache = json.load(f)
        stocks = cache.get("stocks", {})
        print(f"キャッシュ読み込み: {len(stocks)}銘柄")
        return stocks
    except Exception as e:
        print(f"キャッシュ読み込みエラー: {e}")
        return {}

def _sanitize_for_json(obj):
    """Infinity/NaN を null に置換（標準JSONで扱えないため）"""
    import math
    if isinstance(obj, float):
        if math.isinf(obj) or math.isnan(obj):
            return None
        return obj
    if isinstance(obj, dict):
        return {k: _sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_for_json(v) for v in obj]
    return obj


def save_cache(stocks_dict):
    try:
        payload = _sanitize_for_json({"saved_at": datetime.now().isoformat(), "stocks": stocks_dict})
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        print(f"キャッシュ保存: {len(stocks_dict)}銘柄")
    except Exception as e:
        print(f"キャッシュ保存エラー: {e}")


# ============================================================
#  J-Quants API
# ============================================================
def fetch_jquants_info():
    api_key = os.environ.get("JQUANTS_REFRESH_TOKEN") or os.environ.get("JQUANTS_API_KEY")

    if os.path.exists(JQUANTS_INFO_FILE):
        try:
            with open(JQUANTS_INFO_FILE, encoding="utf-8") as f:
                cached = json.load(f)
            saved_at = datetime.fromisoformat(cached.get("saved_at", "2000-01-01"))
            if datetime.now() - saved_at < timedelta(days=30):
                info = cached.get("info", {})
                print(f"既存J-Quantsキャッシュ使用: {len(info)}銘柄")
                return info
        except Exception:
            pass

    if not api_key:
        print("⚠️  J-Quants APIキーなし（業種名は更新されません）")
        return None

    headers = {"x-api-key": api_key}
    for days_ago in [90, 180, 365]:
        target_date = (datetime.now() - timedelta(days=days_ago)).strftime("%Y%m%d")
        for attempt in range(3):
            try:
                r = requests.get("https://api.jquants.com/v2/equities/master",
                    headers=headers, params={"date": target_date}, timeout=15)
                if r.status_code == 429:
                    print(f"レートリミット、15秒待機...")
                    time.sleep(15)
                    continue
                if r.status_code == 200:
                    items = r.json().get("data", r.json().get("items", []))
                    if items:
                        info = {}
                        for item in items:
                            code = str(item.get("Code", "")).zfill(4)[:4]
                            info[code] = {
                                "jaName":  item.get("CoName", ""),
                                "s33Code": item.get("S33", ""),
                                "s33Name": item.get("S33Nm", ""),
                                "s17Code": item.get("S17", ""),
                                "s17Name": item.get("S17Nm", ""),
                                "market":  item.get("MktNm", ""),
                            }
                        with open(JQUANTS_INFO_FILE, "w", encoding="utf-8") as f:
                            json.dump({"saved_at": datetime.now().isoformat(), "info": info},
                                      f, ensure_ascii=False)
                        print(f"J-Quants銘柄情報取得: {len(info)}銘柄 (日付: {target_date})")
                        return info
                break
            except Exception as e:
                print(f"J-Quants取得エラー (日付:{target_date}): {e}")
                break
    return None


def apply_jquants_info(cache, jquants_info):
    if not jquants_info:
        return 0
    updated = 0
    for code, data in cache.items():
        if data.get("isETF"):
            # ETF/REITは専用ルートで業種(ETF・他)を設定済み。J-Quantsの「その他」で上書きしない
            continue
        info = jquants_info.get(str(code).zfill(4))
        if info:
            data["jaName"]  = info["jaName"] or data.get("jaName")
            data["s33Name"] = info["s33Name"]
            data["s33Code"] = info["s33Code"]
            data["s17Name"] = info["s17Name"]
            data["market"]  = info.get("market", data.get("market"))
            updated += 1
    return updated


# ============================================================
def calc_trend(values):
    n = len(values)
    if n < 2:
        return {"slope": 0, "r2": 0, "growthRate": 0}
    xm = (n-1)/2
    ym = sum(values)/n
    num = sum((i-xm)*(v-ym) for i,v in enumerate(values))
    den = sum((i-xm)**2 for i in range(n))
    if den == 0:
        return {"slope": 0, "r2": 0, "growthRate": 0}
    slope = num/den
    yp = [ym+slope*(i-xm) for i in range(n)]
    ss_tot = sum((v-ym)**2 for v in values)
    ss_res = sum((v-yp[i])**2 for i,v in enumerate(values))
    r2 = 1 - ss_res/ss_tot if ss_tot != 0 else 1.0

    base = abs(values[0]) if values[0] != 0 else (abs(ym) if ym != 0 else 1)
    growth_rate = round(slope / base * 100, 2)

    return {"slope": round(slope,4), "r2": round(r2,4), "growthRate": growth_rate}


def remove_dividend_outliers(dividends):
    if len(dividends) < 3:
        return dividends[:]
    result = dividends[:]
    for i in range(1, len(result) - 1):
        prev, cur, nxt = result[i-1], result[i], result[i+1]
        median_neighbors = (prev + nxt) / 2
        if median_neighbors > 0:
            ratio = cur / median_neighbors
            if ratio > 1.8 or ratio < 0.5:
                result[i] = round(median_neighbors)
    return result


def has_no_dividend_cut(dividends):
    if len(dividends) < 2:
        return True
    i = 1
    consecutive_cuts = 0
    while i < len(dividends):
        prev = dividends[i - 1]
        cur  = dividends[i]
        if i >= 2:
            prev_prev = dividends[i - 2]
            prev_is_spike = (prev_prev > 0 and prev / prev_prev >= 1.3)
        else:
            prev_is_spike = False
        if prev_is_spike:
            baseline = dividends[i - 2]
            if cur < baseline * 0.95:
                return False
            consecutive_cuts = 0
        else:
            if cur < prev * 0.95:
                consecutive_cuts += 1
                if consecutive_cuts >= 2:
                    return False
            else:
                consecutive_cuts = 0
        i += 1
    return True


def consecutive_dividend_growth(dividends):
    cleaned = remove_dividend_outliers(dividends)
    if not cleaned or len(cleaned) < 2:
        return 0
    count = 0
    for i in range(len(cleaned)-1, 0, -1):
        if cleaned[i] >= cleaned[i-1] * 0.95:
            count += 1
        else:
            break
    return count


def fetch_stock_data(code, name_hint=""):
    try:
        ticker = yf.Ticker(f"{code}.T")
        info = ticker.info

        if not info or len(info) < 3:
            return None

        qt = info.get("quoteType", "")
        if qt and qt.upper() not in ("EQUITY", ""):
            return None

        if info.get("currency") and info.get("currency") != "JPY":
            return None

        stmt = None
        for attr in ("income_stmt", "financials"):
            try:
                s = getattr(ticker, attr)
                if s is not None and not s.empty:
                    stmt = s
                    break
            except Exception:
                continue

        if stmt is None:
            return None

        years, revs, profs = [], [], []
        for col in sorted(stmt.columns):
            try:
                year = col.year if hasattr(col, "year") else int(str(col)[:4])
                rev = None
                for rev_key in ("Total Revenue", "Revenue", "Net Revenue"):
                    if rev_key in stmt.index:
                        rev = stmt.loc[rev_key, col]
                        break
                op = None
                for op_key in ("Operating Income", "Operating Revenue", "Ebit"):
                    if op_key in stmt.index:
                        op = stmt.loc[op_key, col]
                        break
                if rev is not None and op is not None:
                    try:
                        rev_val = float(rev)
                        op_val  = float(op)
                        if rev_val != 0:
                            years.append(year)
                            revs.append(round(rev_val / 1e8))
                            profs.append(round(op_val / 1e8))
                    except (ValueError, TypeError):
                        continue
            except Exception:
                continue

        if len(years) < 3:
            return None

        divs = ticker.dividends
        div_annual = {}
        if divs is not None and not divs.empty:
            for ts, val in divs.items():
                y = ts.year
                div_annual[y] = div_annual.get(y, 0) + float(val)

        div_years  = sorted(div_annual.keys())[-10:]
        div_values = [round(div_annual[y]) for y in div_years]

        hist = ticker.history(period="10y")
        yearly_price_dict = {}
        if not hist.empty:
            hist['Year'] = hist.index.year
            yearly_mean = hist.groupby('Year')['Close'].mean()
            for y, val in yearly_mean.items():
                yearly_price_dict[y] = val

        yearly_prices = []
        for y in div_years:
            yearly_prices.append(round(float(yearly_price_dict.get(y, 0)), 1))

        name = info.get("longName") or info.get("shortName") or name_hint or code
        div_streak = consecutive_dividend_growth(div_values)

        current_price = float(
            info.get("currentPrice") or
            info.get("regularMarketPrice") or
            info.get("previousClose") or
            0
        )

        import pandas as pd
        latest_annual_div = 0
        if divs is not None and not divs.empty:
            try:
                tz = divs.index.tz
                now = pd.Timestamp.now(tz=tz)
                cutoff = now - pd.DateOffset(months=12)
                recent = divs[(divs.index >= cutoff) & (divs.index <= now)]
                latest_annual_div = round(float(recent.sum())) if not recent.empty else 0
            except Exception:
                latest_annual_div = div_values[-1] if div_values else 0

        # 直近配当が前年比3倍超 → yfinance 重複集計バグ、前年値で補正
        if len(div_values) >= 2 and div_values[-2] > 0 and div_values[-1] > div_values[-2] * 3:
            div_values[-1] = div_values[-2]
            latest_annual_div = div_values[-1]

        if current_price > 0 and latest_annual_div > 0:
            div_yield = round(latest_annual_div / current_price * 100, 2)
            if div_yield > 100:  # 総支払額バグ
                div_yield = 0
        else:
            div_yield = 0

        return {
            "code": code,
            "name": name,
            "jaName": name_hint,
            "sector": info.get("sector","不明"),
            "industry": info.get("industry",""),
            "businessSummaryEn": info.get("longBusinessSummary",""),
            "market": info.get("exchange",""),
            "per": round(float(info.get("trailingPE") or 0),1),
            "pbr": round(float(info.get("priceToBook") or 0),1),
            "marketCap": info.get("marketCap"),
            "years": years,
            "revenue": revs,
            "profit": profs,
            "dividend": div_values,
            "dividendYears": div_years,
            "yearlyPrices": yearly_prices,
            "revenueTrend": calc_trend(revs),
            "profitTrend":  calc_trend(profs),
            "noDividendCut": has_no_dividend_cut(div_values) if len(div_values)>=3 else None,
            "dividendStreak": div_streak,
            "dividendYield": div_yield,
            "currentPrice": current_price,
            "cachedAt": datetime.now().isoformat(),
        }
    except Exception:
        return None


# ============================================================
#  ETF / REIT 専用ルート
# ============================================================
def fetch_etf_data(code, name_hint=""):
    """ETF・REIT等、財務諸表を持たない上場商品の価格・分配金を取得する。
    fetch_stock_data と違い quoteType や損益計算書の有無で弾かない。"""
    try:
        import pandas as pd
        ticker = yf.Ticker(f"{code}.T")
        try:
            info = ticker.info or {}
        except Exception:
            info = {}

        # 通常株式は個別株ルート(fetch_stock_data)の担当。誤ってETF扱いで登録しない
        if str(info.get("quoteType", "")).upper() == "EQUITY":
            return None

        price = 0.0
        try:
            hist = ticker.history(period="5d")
            if hist is not None and not hist.empty:
                price = float(hist["Close"].dropna().iloc[-1])
        except Exception:
            pass
        if price <= 0:
            price = float(info.get("regularMarketPrice") or info.get("previousClose") or 0)
        if price <= 0:
            return None

        div_annual = {}
        latest_annual_div = 0
        try:
            divs = ticker.dividends
        except Exception:
            divs = None
        if divs is not None and not divs.empty:
            for ts, val in divs.items():
                div_annual[ts.year] = div_annual.get(ts.year, 0) + float(val)
            try:
                tz = divs.index.tz
                now = pd.Timestamp.now(tz=tz)
                cutoff = now - pd.DateOffset(months=12)
                recent = divs[(divs.index >= cutoff) & (divs.index <= now)]
                latest_annual_div = round(float(recent.sum()), 1) if not recent.empty else 0
            except Exception:
                latest_annual_div = 0

        # 暦年別の分配金履歴（当年は集計途中で不完全なため直近12ヶ月合計に置き換える）
        current_year = datetime.now().year
        div_years  = sorted(y for y in div_annual if y < current_year)[-9:]
        div_values = [round(div_annual[y], 1) for y in div_years]
        if latest_annual_div > 0:
            div_years.append(current_year)
            div_values.append(latest_annual_div)
        elif div_values:
            latest_annual_div = div_values[-1]

        name = info.get("longName") or info.get("shortName") or name_hint or code
        div_yield = round(latest_annual_div / price * 100, 2) if latest_annual_div > 0 and price > 0 else 0

        return {
            "code": code,
            "name": name,
            "jaName": name_hint or name,
            "sector": "ETF・他",
            "s33Name": "ETF・他",
            "industry": "ETF",
            "market": info.get("exchange", ""),
            "isETF": True,
            "per": 0,
            "pbr": 0,
            "marketCap": info.get("totalAssets") or info.get("marketCap"),
            "years": [],
            "revenue": [],
            "profit": [],
            "dividend": div_values,
            "dividendYears": div_years,
            "dividendYield": div_yield,
            "currentPrice": round(price, 1),
            "cachedAt": datetime.now().isoformat(),
        }
    except Exception:
        return None


def load_portfolio_holdings():
    """data/portfolio.json の保有銘柄 [(code, name_hint), ...] を返す"""
    try:
        with open(PORTFOLIO_FILE, encoding="utf-8") as f:
            pf = json.load(f)
        result = []
        for h in pf.get("holdings", []):
            code = str(h.get("code") or "").strip()
            if code:
                result.append((code, (h.get("manual") or {}).get("name", "")))
        return result
    except Exception as e:
        print(f"portfolio.json 読み込みエラー: {e}")
        return []


def update_portfolio_etfs(cache):
    """ポートフォリオ保有銘柄のうち、個別株ルート（財務諸表ベース）で取得できない
    ETF・REIT等を専用ルートで取得・更新する。毎回の実行で価格・分配金を最新化する。"""
    targets = []
    for code, name_hint in load_portfolio_holdings():
        entry = cache.get(code)
        if entry is None or entry.get("isETF"):
            targets.append((code, name_hint))
    if not targets:
        return

    print(f"ETF/REITルート更新: {len(targets)}銘柄 ({', '.join(c for c, _ in targets)})")
    updated = 0
    for code, name_hint in targets:
        time.sleep(FETCH_DELAY)
        result = None
        # 未取得銘柄はまず個別株ルートを試す（単に前回失敗しただけの個別株かもしれない）
        if cache.get(code) is None:
            result = fetch_stock_data(code, name_hint)
        if result is None:
            result = fetch_etf_data(code, name_hint)
        if result is None:
            print(f"  ⚠️ {code} 取得失敗（既存データを維持）")
            continue
        cache[code] = result
        updated += 1
        yld = result.get("dividendYield", 0)
        print(f"  ✔ {code} {result.get('name','')}: 価格 {result.get('currentPrice')}円 / 利回り {yld}%")

    if updated:
        set_cache(cache)
        save_cache(cache)
    print(f"✅ ETF/REIT更新: {updated}/{len(targets)}")


# ============================================================
#  ir-bank 補完
# ============================================================
_IRBANK_DIAG_PRINTED = 0
_IRBANK_DIAG_MAX = 3

def _fetch_irbank_dividend_page(code, ua_headers):
    """ir-bank の /dividend ページから 年度 → 1株配当 の辞書を返す"""
    try:
        from bs4 import BeautifulSoup
        import re
        resp = requests.get(f"https://irbank.net/{code}/dividend",
                            headers=ua_headers, timeout=10)
        if resp.status_code != 200:
            return {}
        soup = BeautifulSoup(resp.text, "html.parser")
        result = {}
        for table in soup.find_all("table"):
            rows = table.find_all("tr")
            if not rows: continue
            header = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]
            if not header: continue
            # 年度列の特定
            year_idx = next((i for i, h in enumerate(header)
                             if h in ("年度", "決算期", "期")), None)
            if year_idx is None: continue
            # 合計/年間/一株配当 を優先、無ければ「配当」を含む列
            def _div_col(h):
                if "配" not in h: return False
                if any(x in h for x in ("性向", "利回", "総額", "落ち", "落日", "権利", "回数")): return False
                return True
            div_idx = next((i for i, h in enumerate(header)
                            if h in ("合計", "年配", "年間配当") or "1株配" in h or "一株配" in h), None)
            if div_idx is None:
                div_idx = next((i for i, h in enumerate(header) if _div_col(h)), None)
            if div_idx is None: continue

            for row in rows[1:]:
                cells = row.find_all(["th", "td"])
                if len(cells) <= max(year_idx, div_idx): continue
                m = re.match(r"(\d{4})", cells[year_idx].get_text(strip=True))
                if not m: continue
                yr = int(m.group(1))
                s = cells[div_idx].get_text(strip=True).replace(",", "")
                if not s or s in ("-", "―", "－", "—"): continue
                try:
                    v = float(re.sub(r"[^\d.\-]", "", s))
                    if v >= 0: result[yr] = round(v)
                except Exception:
                    continue
            if result: break  # 最初に当たった有効テーブルを採用
        return result
    except Exception:
        return {}


def fetch_irbank_financials(code):
    global _IRBANK_DIAG_PRINTED
    try:
        from bs4 import BeautifulSoup
        import re
        ua_headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "ja,en;q=0.9",
        }
        url = f"https://irbank.net/{code}/results"
        resp = requests.get(url, headers=ua_headers, timeout=10)
        if resp.status_code != 200:
            return None
        soup = BeautifulSoup(resp.text, "html.parser")
        tables = soup.find_all("table")
        if not tables:
            return None
        REV_KEYWORDS = ["売上", "完成工事高", "営業収益", "経常収益", "売収", "収益"]

        def parse_jpy(s):
            s = s.strip().replace(",", "").replace(" ", "")
            if not s or s in ("-", "―", "－", "—", ""): return None
            try:
                if "兆" in s: return round(float(s.replace("兆", "")) * 10000)
                elif "億" in s: return round(float(s.replace("億", "")))
                elif "百万" in s or "M" in s: return round(float(re.sub(r"[^\d.\-]", "", s)) / 100)
                else:
                    v = float(re.sub(r"[^\d.\-]", "", s))
                    return round(v / 100) if abs(v) > 1000 else round(v)
            except: return None

        def parse_float(s):
            s = s.strip().replace(",", "")
            if not s or s in ("-", "―", "－", "—", ""): return None
            try: return float(re.sub(r"[^\d.\-]", "", s))
            except: return None

        # 業績テーブル(売上+営利を持つ)と配当テーブル(一株配当を持つ)を別々に特定
        fin_table = fin_header = fin_rows = None
        div_table = div_header = div_rows = div_col_idx = None
        for t in tables:
            rs = t.find_all("tr")
            if not rs: continue
            h = [c.get_text(strip=True) for c in rs[0].find_all(["th", "td"])]
            if "年度" not in h or len(h) < 3:
                continue
            if fin_table is None:
                has_rev = any(any(k in c for k in REV_KEYWORDS) for c in h)
                has_prof = any("営利" in c for c in h)
                if has_rev and has_prof:
                    fin_table, fin_header, fin_rows = t, h, rs
                    continue
            if div_table is None:
                # 「一株配当」「1株配当」を最優先で採用
                idx = next((i for i, c in enumerate(h) if c in ("一株配当", "1株配当", "1株配", "年間配当")), None)
                if idx is not None:
                    div_table, div_header, div_rows, div_col_idx = t, h, rs, idx

        if fin_table is None:
            return None

        try:
            rev_idx  = next(i for i, h in enumerate(fin_header) if any(k in h for k in REV_KEYWORDS))
            prof_idx = next(i for i, h in enumerate(fin_header) if "営利" in h)
        except StopIteration:
            return None
        eps_idx = next((i for i, h in enumerate(fin_header) if h == "EPS"), None)

        # 業績/EPSを抽出
        years, revs, profs, epss = [], [], [], []
        for row in fin_rows[1:]:
            cells = row.find_all(["th", "td"])
            if len(cells) <= max(rev_idx, prof_idx): continue
            yr_str = cells[0].get_text(strip=True)
            m = re.match(r"(\d{4})", yr_str)
            if not m: continue
            yr = int(m.group(1))
            rv = parse_jpy(cells[rev_idx].get_text(strip=True))
            pf = parse_jpy(cells[prof_idx].get_text(strip=True))
            ep = parse_float(cells[eps_idx].get_text(strip=True)) if eps_idx and len(cells) > eps_idx else None
            if rv is not None and pf is not None and rv != 0:
                years.append(yr)
                revs.append(rv)
                profs.append(pf)
                epss.append(ep)

        if len(years) < 5: return None

        # 配当を別テーブルから抽出（年度→1株配当の辞書）
        div_by_year = {}
        if div_table is not None:
            for row in div_rows[1:]:
                cells = row.find_all(["th", "td"])
                if len(cells) <= div_col_idx: continue
                m = re.match(r"(\d{4})", cells[0].get_text(strip=True))
                if not m: continue
                yr = int(m.group(1))
                v = parse_float(cells[div_col_idx].get_text(strip=True))
                if v is not None and v >= 0:
                    div_by_year[yr] = round(v)

        combined = sorted(zip(years, revs, profs, epss))
        years = [x[0] for x in combined]
        revs  = [x[1] for x in combined]
        profs = [x[2] for x in combined]
        epss  = [x[3] for x in combined]
        divs  = [div_by_year.get(y) for y in years]

        eps_clean_years  = [years[i] for i, v in enumerate(epss) if v is not None]
        eps_clean_values = [v for v in epss if v is not None]

        div_clean_years  = [years[i] for i, v in enumerate(divs) if v is not None]
        div_clean_values = [v for v in divs if v is not None]

        # /results の配当列が空/未検出のときは /dividend ページから補完
        if not div_clean_values:
            div_map = _fetch_irbank_dividend_page(code, ua_headers)
            if div_map:
                div_clean_years  = sorted(div_map.keys())
                div_clean_values = [div_map[y] for y in div_clean_years]
            elif _IRBANK_DIAG_PRINTED < _IRBANK_DIAG_MAX:
                _IRBANK_DIAG_PRINTED += 1
                print(f"  [diag] {code} dividend extraction failed. div_table={'yes' if div_table is not None else 'no'} div_header={div_header}")

        return {
            "years": years,
            "revenue": revs,
            "profit": profs,
            "eps": epss,
            "epsYears": eps_clean_years,
            "epsValues": eps_clean_values,
            "dividendYears": div_clean_years,
            "dividend": div_clean_values,
        }
    except Exception:
        return None


# ============================================================
#  バッチ実行
# ============================================================
_consecutive_401 = 0
_rate_limit_lock = threading.Lock()
_rate_limited_until = 0


def fetch_all(stock_list):
    global _consecutive_401, _rate_limited_until
    total = len(stock_list)
    cache_data = get_cache()
    _consecutive_401 = 0
    _rate_limited_until = 0

    todo = [item for item in stock_list if item["code"] not in cache_data]
    skipped = total - len(todo)
    print(f"全{total}社 / 未取得{len(todo)}社（既存キャッシュ {skipped}社をスキップ）")

    def fetch_one(item):
        global _consecutive_401, _rate_limited_until
        code = item["code"]
        name = item.get("name", "")

        wait_until = _rate_limited_until
        if wait_until > time.time():
            wait_sec = wait_until - time.time()
            print(f"  ⏳ レート制限待機 {int(wait_sec)}秒")
            time.sleep(wait_sec + 1)

        time.sleep(FETCH_DELAY)
        result = fetch_stock_data(code, name)

        with _rate_limit_lock:
            if result is None:
                _consecutive_401 += 1
                if _consecutive_401 >= 5:
                    print(f"  ⚠️  連続失敗{_consecutive_401}回 → {RETRY_WAIT_SEC}秒待機")
                    _rate_limited_until = time.time() + RETRY_WAIT_SEC
                    _consecutive_401 = 0
            else:
                _consecutive_401 = 0
        return code, result

    done = skipped
    found = skipped
    errors = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(fetch_one, item): item for item in todo}
        for f in as_completed(futures):
            code, result = f.result()
            done += 1
            if result:
                cache_data[code] = result
                found += 1
            else:
                errors += 1
            if done % 50 == 0:
                pct = int(done/total*100)
                print(f"  {done}/{total} ({pct}%) found={found} err={errors}")
                save_cache(cache_data)

    set_cache(cache_data)
    save_cache(cache_data)
    print(f"\n✅ 完了: {found}/{total}銘柄 取得 (エラー{errors}件)")
    return cache_data


def latest_closed_session_date():
    """東証で直近に大引けを迎えた営業日を、実データ(トヨタ)から判定して date で返す。

    yfinance は場中だとその日の未確定バーも返すため、「まだ大引けしていない日」を
    除外しないと場中の値を終値として保存してしまう。祝日カレンダーを持たなくて済むよう
    流動性の高い銘柄の実際の取引日から判定する。

    判定は必ず dropna() 後の行から行う。Yahoo は大引け直後、終値が NaN のままの行を
    先に作ることがあり、index だけを見ると「その日の終値がある」と誤認する。すると
    取得側(dropna 済み)とズレて、実際には前営業日の終値しか保存できていないのに
    当日分を取得済みと記録してしまい、終値が配信された後に取りに行かなくなる。
    """
    try:
        hist = yf.Ticker("7203.T").history(period="10d")
        if hist.empty:
            return None
        closes = hist["Close"].dropna()
        if closes.empty:
            return None
        now_jst = datetime.now(JST)
        for d in sorted({ts.date() for ts in closes.index}, reverse=True):
            # 大引け15:00 + 確定待ちの余裕10分
            if now_jst >= datetime(d.year, d.month, d.day, 15, 10, tzinfo=JST):
                return d
    except Exception:
        pass
    return None


def update_prices_only(cache):
    """株価と配当利回りのみ高速に再取得"""
    # 遅延対策で同じ日に複数回cronを仕掛けているため、目的の終値が既に入っていれば
    # フル再取得をスキップする。判定は「実行した日」ではなく「取得済みの終値の営業日」で
    # 行う: 実行日で判定すると、場中や早朝に走った回で当日分が済んだ扱いになり、
    # 大引け後の本来の更新が丸ごとスキップされてしまう。
    #
    # 比較は「一致」ではなく「対象営業日以降を持っているか」で行う。Yahooの配信は
    # 遅れることがあり、既に新しい終値を持っている状態で古い対象営業日が算出される
    # ことがある。一致で判定すると、その回がスキップされずフル取得に入り、
    # 持っていた新しい終値を古い値で上書きしてしまう(実際に 2026-08-31 の終値が
    # 8/28 に巻き戻る事故が起きた)。
    target_date = latest_closed_session_date()
    if target_date:
        target_key = target_date.isoformat()
        already_done = sum(
            1 for d in cache.values()
            if isinstance(d, dict) and str(d.get("priceDate") or "") >= target_key
        )
        if cache and already_done >= len(cache) * 0.9:
            print(f"⏭ {target_key} 以降の終値は取得済み({already_done}/{len(cache)}銘柄)のためスキップ")
            return
        print(f"対象営業日(大引け済み): {target_key}")

    codes = list(cache.keys())
    total = len(codes)
    print(f"株価更新: {total}銘柄")
    updated = 0
    prices_updated_at = datetime.now().isoformat()
    import pandas as pd

    for i, code in enumerate(codes):
        try:
            ticker = yf.Ticker(f"{code}.T")
            hist = ticker.history(period="10d")
            if hist.empty:
                continue
            closes = hist["Close"].dropna()
            if target_date is not None:
                # 大引け前の未確定バーを終値として拾わないよう対象営業日以前に限定する
                closes = closes[[ts.date() <= target_date for ts in closes.index]]
            if closes.empty:
                continue
            price = round(float(closes.iloc[-1]))
            if price <= 0:
                continue
            price_date = closes.index[-1].date().isoformat()

            # 既により新しい営業日の終値を持っているなら古い値で塗り替えない。
            # Yahooは同じ銘柄でも呼ぶタイミングによって直近の終値をまだ返さないことが
            # あり、そのまま書き戻すと表示が前の営業日に巻き戻ってしまう。
            if str(cache[code].get("priceDate") or "") > price_date:
                continue

            latest_div = 0
            try:
                divs = ticker.dividends
                if divs is not None and not divs.empty:
                    tz = divs.index.tz
                    now = pd.Timestamp.now(tz=tz)
                    cutoff = now - pd.DateOffset(months=12)
                    recent = divs[(divs.index >= cutoff) & (divs.index <= now)]
                    latest_div = round(float(recent.sum())) if not recent.empty else 0
            except Exception:
                pass

            if latest_div == 0:
                div_vals = cache[code].get("dividend", [])
                latest_div = div_vals[-1] if div_vals else 0

            cache[code]["currentPrice"] = price
            cache[code]["priceDate"] = price_date
            cache[code]["pricesUpdatedAt"] = prices_updated_at
            if latest_div > 0:
                cache[code]["dividendYield"] = round(latest_div / price * 100, 2)
            updated += 1
        except Exception:
            pass

        if (i + 1) % 100 == 0:
            print(f"  {i+1}/{total} updated={updated}")
            save_cache(cache)
        time.sleep(0.4)

    save_cache(cache)
    print(f"✅ 株価更新完了: {updated}/{total}")


def enrich_irbank(cache):
    """ir-bank から長期業績データを補完"""
    force = os.environ.get("FORCE_IRBANK") == "1"
    if force:
        targets = list(cache.keys())
    else:
        targets = [code for code, d in cache.items() if not d.get("irbank_enriched")]
    total = len(targets)
    print(f"ir-bank補完: {total}銘柄対象")
    enriched = 0

    for i, code in enumerate(targets):
        time.sleep(1.0)
        result = fetch_irbank_financials(code)
        if result and len(result["years"]) >= 5:
            cache[code]["years"]        = result["years"]
            cache[code]["revenue"]      = result["revenue"]
            cache[code]["profit"]       = result["profit"]
            cache[code]["eps"]          = result.get("eps", [])
            cache[code]["epsYears"]     = result.get("epsYears", [])
            cache[code]["epsValues"]    = result.get("epsValues", [])
            cache[code]["revenueTrend"] = calc_trend(result["revenue"])
            cache[code]["profitTrend"]  = calc_trend(result["profit"])
            cache[code]["epsTrend"]     = calc_trend(result["epsValues"]) if len(result.get("epsValues", [])) >= 3 else {"slope": 0, "r2": 0, "growthRate": 0}

            # 配当（会計年度ベース、ir-bank 由来。yfinanceの暦年集計より正確）
            div_vals  = result.get("dividend", []) or []
            div_years = result.get("dividendYears", []) or []
            if len(div_vals) >= 2:
                # 直近10年に絞る
                div_vals_10  = div_vals[-10:]
                div_years_10 = div_years[-10:]
                cache[code]["dividend"]       = div_vals_10
                cache[code]["dividendYears"]  = div_years_10
                cache[code]["dividendStreak"] = consecutive_dividend_growth(div_vals_10)
                cache[code]["noDividendCut"]  = has_no_dividend_cut(div_vals_10) if len(div_vals_10) >= 3 else None
                # 配当利回りを最新確定値ベースで再計算
                price = cache[code].get("currentPrice") or 0
                # 末尾が None/0 のときは一つ前を採用（予想未確定対策）
                latest_div = next((v for v in reversed(div_vals_10) if v), 0)
                if price > 0 and latest_div > 0:
                    yld = round(latest_div / price * 100, 2)
                    if yld <= 100:
                        cache[code]["dividendYield"] = yld

            cache[code]["irbank_enriched"] = True
            enriched += 1
        if (i + 1) % 50 == 0:
            print(f"  {i+1}/{total} enriched={enriched}")
            save_cache(cache)

    save_cache(cache)
    print(f"✅ ir-bank補完完了: {enriched}/{total}")


def enrich_business_summary(cache):
    """yfinanceの英語事業概要(longBusinessSummary)を取得し、日本語に機械翻訳して補完する"""
    from deep_translator import GoogleTranslator

    force = os.environ.get("FORCE_SUMMARY") == "1"
    if force:
        targets = list(cache.keys())
    else:
        targets = [code for code, d in cache.items() if not d.get("businessSummaryJa")]
    total = len(targets)
    print(f"事業概要補完: {total}銘柄対象")
    enriched = 0
    translator = GoogleTranslator(source="en", target="ja")

    for i, code in enumerate(targets):
        time.sleep(1.0)
        try:
            summary_en = cache[code].get("businessSummaryEn", "")
            if not summary_en:
                ticker = yf.Ticker(f"{code}.T")
                summary_en = ticker.info.get("longBusinessSummary", "") or ""
                if summary_en:
                    cache[code]["businessSummaryEn"] = summary_en

            if summary_en:
                summary_ja = translator.translate(summary_en)
                if summary_ja:
                    cache[code]["businessSummaryJa"] = summary_ja
                    enriched += 1
        except Exception as e:
            print(f"  ⚠️ {code}: {e}")

        if (i + 1) % 50 == 0:
            print(f"  {i+1}/{total} enriched={enriched}")
            save_cache(cache)

    save_cache(cache)
    print(f"✅ 事業概要補完完了: {enriched}/{total}")


# ============================================================
def diag_prices(codes):
    """yfinance が実際に何を返しているかをそのまま出力する。

    「終値が取れない」ときに、Yahoo にデータが無いのか、こちらの取り出し方
    (期間指定・タイムゾーン・NaN・auto_adjust など)が悪いのかを推測せず
    切り分けるための診断。取得結果を加工せずそのまま出す。
    """
    print(f"yfinance version : {getattr(yf, '__version__', '?')}")
    print(f"now UTC          : {datetime.now(timezone.utc).isoformat()}")
    print(f"now JST          : {datetime.now(JST).isoformat()}")
    variants = [
        ("period=10d",                {"period": "10d"}),
        ("period=1mo",                {"period": "1mo"}),
        ("period=10d,auto_adjust=F",  {"period": "10d", "auto_adjust": False}),
        ("start=-14d",                {"start": (datetime.now(JST) - timedelta(days=14)).strftime("%Y-%m-%d")}),
    ]
    for code in codes:
        print("=" * 55)
        print(f"[{code}.T]")
        for label, kwargs in variants:
            try:
                hist = yf.Ticker(f"{code}.T").history(**kwargs)
                tz = getattr(hist.index, "tz", None)
                print(f"  -- {label}: {len(hist)}行 index.tz={tz}")
                if hist.empty:
                    continue
                for ts, row in hist.tail(4).iterrows():
                    close = row.get("Close")
                    vol = row.get("Volume")
                    print(f"       ts={ts} date={ts.date()} Close={close!r} Volume={vol!r}")
            except Exception as e:
                print(f"  -- {label}: 失敗 {type(e).__name__}: {e}")


def diag_irbank(code):
    """1銘柄のir-bank /results /dividend ページ構造をダンプ（パーサ調整用）"""
    from bs4 import BeautifulSoup
    ua_headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "ja,en;q=0.9",
    }
    for path in ("results", "dividend"):
        url = f"https://irbank.net/{code}/{path}"
        print(f"\n=========== {url} ===========")
        try:
            r = requests.get(url, headers=ua_headers, timeout=10)
            print(f"status: {r.status_code}, len={len(r.text)}")
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            tables = soup.find_all("table")
            print(f"tables found: {len(tables)}")
            for ti, t in enumerate(tables[:6]):
                rs = t.find_all("tr")
                if not rs: continue
                header = [c.get_text(strip=True) for c in rs[0].find_all(["th","td"])]
                print(f"\n--- table[{ti}] header ({len(rs)} rows) ---")
                print("  HEADER:", header)
                for ri, row in enumerate(rs[1:8]):
                    cells = [c.get_text(strip=True) for c in row.find_all(["th","td"])]
                    print(f"  row[{ri}]:", cells)
        except Exception as e:
            print(f"ERROR: {e}")


def main():
    print("=" * 55)
    print("  日本株スクリーナー - データ更新")
    print("=" * 55)

    cache = load_cache()
    set_cache(cache)

    update_only = os.environ.get("UPDATE_PRICES_ONLY") == "1"
    do_irbank = os.environ.get("ENRICH_IRBANK") == "1"
    do_summary = os.environ.get("ENRICH_SUMMARY") == "1"
    fetch_limit = int(os.environ.get("FETCH_LIMIT") or 0) or None
    diag_code = os.environ.get("DIAG_IRBANK") or ""
    diag_price_codes = os.environ.get("DIAG_PRICES") or ""

    if diag_price_codes:
        diag_prices([c.strip() for c in diag_price_codes.split(",") if c.strip()])
        return

    if diag_code:
        diag_irbank(diag_code.strip())
        return

    if update_only:
        if not cache:
            print("❌ 既存キャッシュなし。先にフルフェッチが必要です")
            sys.exit(1)
        update_prices_only(cache)
    elif do_irbank:
        if not cache:
            print("❌ 既存キャッシュなし。先にフルフェッチが必要です")
            sys.exit(1)
        enrich_irbank(cache)
        # 配当を更新したあとに株価も最新化（dividendYield を最新株価ベースで再計算）
        print("\n--- 株価も最新化します ---")
        update_prices_only(cache)
    elif do_summary:
        if not cache:
            print("❌ 既存キャッシュなし。先にフルフェッチが必要です")
            sys.exit(1)
        enrich_business_summary(cache)
    else:
        codes = load_tse_codes()
        if not codes:
            print("❌ 銘柄リスト取得失敗")
            sys.exit(1)
        if fetch_limit:
            codes = codes[:fetch_limit]
            print(f"  取得制限: {fetch_limit}社")
        cache = fetch_all(codes)

    # ポートフォリオ内のETF/REITは個別株と別ルートで毎回更新（1343 東証REIT指数ETF など）
    update_portfolio_etfs(cache)

    # J-Quants 業種情報を反映
    jq = fetch_jquants_info()
    if jq:
        cache = get_cache() or cache
        updated = apply_jquants_info(cache, jq)
        save_cache(cache)
        print(f"🇯🇵 業種情報反映: {updated}銘柄")

    print("=" * 55)
    print("  完了")
    print("=" * 55)


if __name__ == "__main__":
    main()
