"""
ポートフォリオ評価額の日次スナップショットを data/history.json に追記する。
GitHub Actions の株価更新後に実行され、資産推移グラフの元データになる。

- data/portfolio.json（保有銘柄の初期データ）と data/stock_cache.json（株価・配当）から
  評価額・取得額・年間配当（税引後）を計算する。
- フロントの enrichHolding と同じロジック（priceOverride は考慮しない = 初期データベース）。
- 同一日付のエントリが既にあれば上書き、なければ追記する。
"""

import os
import json
from datetime import datetime, timezone, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "data")
PORTFOLIO_FILE = os.path.join(DATA_DIR, "portfolio.json")
CACHE_FILE = os.path.join(DATA_DIR, "stock_cache.json")
HISTORY_FILE = os.path.join(DATA_DIR, "history.json")

JST = timezone(timedelta(hours=9))


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def compute_snapshot(portfolio, cache):
    tax_rate = portfolio.get("taxRate", 0.20315)
    stocks = cache.get("stocks", {})
    total_value = total_cost = total_div_at = 0.0
    priced = 0
    for h in portfolio.get("holdings", []):
        code = str(h.get("code"))
        shares = h.get("shares") or 0
        avg = h.get("avgPrice") or 0
        nisa = min(h.get("nisaShares") or 0, shares)
        m = h.get("manual") or {}
        s = stocks.get(code) or {}
        price = s.get("currentPrice") if (s.get("currentPrice") or 0) > 0 else m.get("price", 0)
        price = price or 0
        div_list = s.get("dividend") or []
        div_ps = div_list[-1] if div_list else m.get("dividendPerShare", 0)
        div_ps = div_ps or 0
        if price > 0:
            priced += 1
        total_value += price * shares
        total_cost += avg * shares
        total_div_at += div_ps * nisa + div_ps * (shares - nisa) * (1 - tax_rate)
    return {
        "date": datetime.now(JST).strftime("%Y-%m-%d"),
        "value": round(total_value),
        "cost": round(total_cost),
        "dividendAfterTax": round(total_div_at),
        "pricedCount": priced,
        "holdingsCount": len(portfolio.get("holdings", [])),
    }


def main():
    portfolio = load_json(PORTFOLIO_FILE, None)
    if not portfolio:
        print("❌ portfolio.json が読めません。スキップします。")
        return
    cache = load_json(CACHE_FILE, {})
    snap = compute_snapshot(portfolio, cache)

    # 株価が1件も取れていない場合は記録しない（不正な0円スナップショットを防ぐ）
    if snap["value"] <= 0 or snap["pricedCount"] == 0:
        print(f"⚠ 評価額が0円のため記録をスキップ: {snap}")
        return

    history = load_json(HISTORY_FILE, [])
    if not isinstance(history, list):
        history = []
    # 同日エントリは上書き
    history = [e for e in history if e.get("date") != snap["date"]]
    history.append(snap)
    history.sort(key=lambda e: e["date"])
    # 直近5年分（約1260営業日）に制限
    history = history[-1300:]

    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=1)
        f.write("\n")
    print(f"✅ 資産推移を記録: {snap}（全{len(history)}件）")


if __name__ == "__main__":
    main()
