"""配当利回り算出ロジックの自己テスト

  python scripts/test_dividend_yield.py

外部通信はしない。update_data.py の resolve_dividend_yield / latest_annual_dividend が
「年間配当（会計年度ベース）を本命、TTM（直近12ヶ月実績）を控え」で動くことを確認する。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from update_data import latest_annual_dividend, resolve_dividend_yield  # noqa: E402

REF_YEAR = 2026

# (説明, entry, 株価, TTM配当, 期待する(利回り, 根拠))
CASES = [
    # 本件: 期末50円+中間25円がTTMの窓に入り75円と合算されていた（3.82%→2.55%）
    ("重複合算を年間配当で是正",
     {"irbank_enriched": True, "dividend": [35, 50], "dividendYears": [2024, 2025]},
     1964, 75, (2.55, "annual")),
    ("TTMと一致していればそのまま",
     {"irbank_enriched": True, "dividend": [90, 95], "dividendYears": [2025, 2026]},
     3096, 95, (3.07, "annual")),
    # ir-bank が中間配当しか拾えていない銘柄はTTMのほぼ半分になる → 年間配当を信用しない
    ("年間配当がTTMの半分＝中間値の疑い→TTM維持",
     {"irbank_enriched": True, "dividend": [75, 88], "dividendYears": [2025, 2026]},
     4634, 175, (3.78, "ttm")),
    # 株式分割の遡及調整がされていない1株配当（株価は調整済み）
    ("年間配当がTTMを大きく上回る（分割未調整）→TTM維持",
     {"irbank_enriched": True, "dividend": [155, 170], "dividendYears": [2024, 2025]},
     1715, 21, (1.22, "ttm")),
    ("進行中年度の部分値は前年の確定値を使う",
     {"irbank_enriched": True, "dividend": [100, 50], "dividendYears": [2025, 2026]},
     1000, 100, (10.0, "annual")),
    ("進行中年度でも増配なら当年を使う",
     {"irbank_enriched": True, "dividend": [100, 120], "dividendYears": [2025, 2026]},
     1000, 120, (12.0, "annual")),
    ("年度が飛んでいる系列では前年に戻さない",
     {"irbank_enriched": True, "dividend": [150, 15], "dividendYears": [2018, 2026]},
     300, 15, (5.0, "annual")),
    ("年度データが古い銘柄はTTMに委ねる",
     {"irbank_enriched": True, "dividend": [22], "dividendYears": [2017]},
     9939, 228, (2.29, "ttm")),
    ("ETF/REITは年度データを持たないのでTTM",
     {"dividend": [50, 60], "dividendYears": [2025, 2026]},
     2000, 60, (3.0, "ttm")),
    ("TTMが無ければ年間配当で算出",
     {"irbank_enriched": True, "dividend": [30, 40], "dividendYears": [2025, 2026]},
     1600, 0, (2.5, "annual")),
    ("パース異常（1株78000円）は採用しない",
     {"irbank_enriched": True, "dividend": [78000], "dividendYears": [2017]},
     943, 0, (0, None)),
    ("株価が無ければ算出しない",
     {"irbank_enriched": True, "dividend": [50], "dividendYears": [2025]},
     0, 50, (0, None)),
    ("無配銘柄",
     {"irbank_enriched": True, "dividend": [0, 0], "dividendYears": [2025, 2026]},
     1000, 0, (0, None)),
]

ANNUAL_CASES = [
    ("末尾が未確定(0)なら一つ前の年度", [30, 35, 0], [2024, 2025, 2026], (35, 2025)),
    ("確定年度をそのまま返す", [30, 35, 50], [2023, 2024, 2025], (50, 2025)),
    ("配当データ無し", [], [], (0, None)),
    ("年と値の本数が合わない", [10, 20], [2025], (0, None)),
]


def main():
    failed = 0
    for name, values, years, expected in ANNUAL_CASES:
        got = latest_annual_dividend(values, years, ref_year=REF_YEAR)
        ok = got == expected
        failed += not ok
        print(f"{'OK ' if ok else 'NG '} latest_annual_dividend: {name} -> {got} (期待 {expected})")

    for name, entry, price, ttm, expected in CASES:
        got = resolve_dividend_yield(entry, price, ttm, ref_year=REF_YEAR)
        ok = got == expected
        failed += not ok
        print(f"{'OK ' if ok else 'NG '} resolve_dividend_yield: {name} -> {got} (期待 {expected})")

    print("-" * 50)
    if failed:
        print(f"❌ {failed}件失敗")
        return 1
    print(f"✅ 全{len(ANNUAL_CASES) + len(CASES)}件パス")
    return 0


if __name__ == "__main__":
    sys.exit(main())
