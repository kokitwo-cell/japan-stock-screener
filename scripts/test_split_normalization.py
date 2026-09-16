"""株式分割による基準ずれ補正の自己テスト

  python scripts/test_split_normalization.py

外部通信はしない。update_data.py の分割補正ロジックが、実際に観測されたデータの形
（2163 アルトナー / 4452 花王 / 6592 マブチモーター）で期待どおりに動くことを確認する。
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd  # noqa: E402

from update_data import (  # noqa: E402
    apply_split_normalization,
    infer_fiscal_year_end_month,
    match_factor,
    normalize_dividend_basis,
    normalize_eps_basis,
    split_factor_candidates,
    unreported_split_factor,
    yahoo_dividends_by_fiscal_year,
    yearly_average_prices,
)

TODAY = date(2026, 9, 16)
failed = 0


def check(name, got, expected):
    global failed
    ok = got == expected
    failed += not ok
    print(f"{'OK ' if ok else 'NG '} {name} -> {got}" + ("" if ok else f" (期待 {expected})"))


def close_series(points):
    """[(日付文字列, 終値)] → 終値 Series（tz付き index）"""
    idx = pd.to_datetime([d for d, _ in points]).tz_localize("Asia/Tokyo")
    return pd.Series([v for _, v in points], index=idx, dtype=float)


# ------------------------------------------------------------------
# 候補倍率と丸め
# ------------------------------------------------------------------
check("候補倍率: 1:2 が2回 → {1,2,4}", split_factor_candidates([(date(2024, 1, 1), 2.0), (date(2026, 1, 1), 2.0)]), [1.0, 2.0, 4.0])
check("候補倍率: 併合0.1 と分割2 → {0.1,0.2,1,2}", split_factor_candidates([(date(2018, 10, 1), 0.1), (date(2025, 4, 1), 2.0)]), [0.1, 0.2, 1.0, 2.0])
check("候補倍率: 分割なし → {1}", split_factor_candidates([]), [1.0])
check("倍率照合: 1.9 → 2", match_factor(1.9, [1.0, 2.0, 4.0]), 2.0)
check("倍率照合: 1.4 は 1 とも 2 とも言えない → None", match_factor(1.4, [1.0, 2.0, 4.0]), None)
check("倍率照合: 自社株買いで 0.85 → 1", match_factor(0.85, [1.0, 2.0]), 1.0)

# ------------------------------------------------------------------
# Yahoo 配当の年度合算（1月決算: 中間=7月末権利落ち、期末=1月末権利落ち）
# ------------------------------------------------------------------
events = [(date(2025, 1, 30), 20.5), (date(2025, 7, 30), 21), (date(2026, 1, 29), 21), (date(2026, 7, 30), 21.5)]
ref = yahoo_dividends_by_fiscal_year(events, [2025, 2026, 2027], fy_end_month=1, today=TODAY)
check("年度合算: 2026年1月期 = 2025/7 + 2026/1", ref[2026], 42.0)
check("年度合算: 期末未到来の2027年1月期は None", ref[2027], None)
check("年度合算: 決算月不明なら暦年（2026 = 1/29 + 7/30）",
      yahoo_dividends_by_fiscal_year(events, [2026], today=TODAY)[2026], None)  # 12月末が未到来
check("年度合算: 暦年2025", yahoo_dividends_by_fiscal_year(events, [2025], today=TODAY)[2025], 41.5)

# ------------------------------------------------------------------
# 配当の基準合わせ
# ------------------------------------------------------------------
C2 = [1.0, 2.0]
# 2163 アルトナー（分割効力発生後）: ir-bank は全年度調整済み、Yahoo も調整済み → 変更なし
yrs = [2022, 2023, 2024, 2025, 2026]
vals, f = normalize_dividend_basis(yrs, [17, 30, 38, 41, 42], {2022: 17, 2023: 30, 2024: 38, 2025: 41, 2026: 42}, C2)
check("配当: 両者一致なら無変更", (vals, f), ([17, 30, 38, 41, 42], [1.0] * 5))

# ir-bank が未調整（分割前の実額）、Yahoo は調整済み → 全年度 1/2
vals, f = normalize_dividend_basis(yrs, [34, 60, 75, 82, 84], {2022: 17, 2023: 30, 2024: 38, 2025: 41, 2026: 42}, C2)
check("配当: 全年度未調整 → 半分に", vals, [17, 30, 37.5, 41, 42])

# 6592 マブチモーター: 2016〜2023 は 2026/1/1 の分割が未調整、2024〜2025 は調整済み
m_years = list(range(2016, 2026))
m_ir = [60, 60, 68, 68, 68, 58, 68, 75, 38, 53]
m_yf = dict(zip(m_years, [30, 30, 34, 34, 34, 29, 34, 38, 38, 53]))
vals, f = normalize_dividend_basis(m_years, m_ir, m_yf, [1.0, 2.0, 4.0])
check("配当: 古い年度だけ未調整（マブチ型）→ 2023年以前を半分に", vals, [30, 30, 34, 34, 34, 29, 34, 37.5, 38, 53])
check("配当: マブチ型の倍率列", f, [2.0] * 8 + [1.0, 1.0])

# Yahoo が1年だけ欠測（半分しか拾えていない）→ 単発なので採用しない
vals, f = normalize_dividend_basis(yrs, [17, 30, 38, 41, 42], {2022: 17, 2023: 30, 2024: 19, 2025: 41, 2026: 42}, C2)
check("配当: Yahoo の単発欠測は無視", vals, [17, 30, 38, 41, 42])

# 予想年度（Yahoo 未確定=None）は直近実績の倍率を引き継ぐ
vals, f = normalize_dividend_basis([2024, 2025, 2026], [75, 82, 84], {2024: 38, 2025: 41, 2026: None}, C2)
check("配当: 期末未到来の年度は直近実績と同じ倍率", vals, [37.5, 41, 42])

# 分割履歴が無ければ、比が合わなくても触らない
vals, f = normalize_dividend_basis([2024, 2025], [80, 84], {2024: 40, 2025: 42}, [1.0])
check("配当: 分割履歴なしなら無変更", vals, [80, 84])

# 併合（1:0.1）: ir-bank の古い年度が併合前の小さい額のまま → 10倍に
vals, f = normalize_dividend_basis([2017, 2018, 2019, 2020], [3, 3, 30, 32], {2017: 30, 2018: 30, 2019: 30, 2020: 32}, [0.1, 1.0])
check("配当: 株式併合の未調整年度は10倍に", vals, [30, 30, 30, 32])

# 新しい年度だけ未調整（古い年度は調整済み）の形も、2年度そろっていれば補正する
vals, f = normalize_dividend_basis([2022, 2023, 2024, 2025], [40, 40, 84, 84], {2022: 40, 2023: 40, 2024: 42, 2025: 42}, C2)
check("配当: 直近2年度だけ未調整のときも補正", vals, [40, 40, 42, 42])
# 最古の年度が単独で外れていても採用しない（10年窓の端は Yahoo 側が年度の途中からしか無い）
vals, f = normalize_dividend_basis([2016, 2017, 2018], [30, 30, 30], {2016: 15, 2017: 30, 2018: 30}, C2)
check("配当: 最古の年度だけの不一致は無視", vals, [30, 30, 30])

# ------------------------------------------------------------------
# EPS の基準合わせ（純利益 ÷ EPS = 当時の株式数 と 現在の株式数 を比較）
# ------------------------------------------------------------------
# 2163: 純利益 12.7億 / EPS 118.64 → 10.7M株。分割後の現在株式数 21.4M → ÷2
eps_years = [2024, 2025, 2026, 2027]
eps_vals = [84.24, 98.99, 118.64, 117.74]
ni = {2024: 900, 2025: 1060, 2026: 1270, 2027: 1260}
new_eps, f = normalize_eps_basis(eps_years, eps_vals, ni, 21_400_000, C2, fy_end_month=1, today=TODAY)
check("EPS: 純利益ベースで全年度が分割前基準 → ÷2", new_eps, [42.12, 49.49, 59.32, 58.87])

# 自社株買いで株式数が1割減っている程度なら 1 のまま
new_eps, f = normalize_eps_basis([2024, 2025], [100.0, 110.0], {2024: 1000, 2025: 1100}, 9_000_000, C2, today=TODAY)
check("EPS: 株式数の1割差は無変更", new_eps, [100.0, 110.0])

# 純利益が無くても、決算発表後に効力が生じた分割の分は必ず補正（発表済み年度＋予想年度）
splits_2163 = [(date(2026, 8, 1), 2.0)]
floor = unreported_split_factor(splits_2163, [2024, 2025, 2026, 2027], fy_end_month=1, today=TODAY)
check("未報告分割: 2026年1月期発表後の 8/1 分割 → 2", floor, 2.0)
new_eps, f = normalize_eps_basis(eps_years, eps_vals, {}, None, C2, floor_factor=floor, fy_end_month=1, today=TODAY)
check("EPS: 純利益なしでも未報告分割分は ÷2（予想年度も同じ基準）", new_eps, [42.12, 49.49, 59.32, 58.87])

# 決算発表済みの年度内に効力が生じた分割は「未報告」ではない
check("未報告分割: 期中の分割は短信で調整済み扱い",
      unreported_split_factor([(date(2025, 7, 1), 2.0)], [2024, 2025, 2026], fy_end_month=1, today=TODAY), 1.0)
# Yahoo の分割日付は権利落ち日。1/1 効力の分割は 12/29 付で載るが、前期の短信では調整されない（マブチ 2026/1/1）
check("未報告分割: 期末2日前の日付（翌期初効力）は未報告扱い",
      unreported_split_factor([(date(2023, 12, 28), 2.0), (date(2025, 12, 29), 2.0)], list(range(2016, 2028)),
                              fy_end_month=12, today=TODAY), 2.0)

# 決算月の推定: 3月決算（9月末・3月末に権利落ち）で、暦年合算だと年度境界がずれて一致しない形
mar_events = []
for y in range(2017, 2027):
    mar_events += [(date(y, 3, 29), 10 + (y - 2017) * 3), (date(y, 9, 28), 10 + (y - 2017) * 3 + 1)]
mar_years = list(range(2018, 2027))
mar_vals = [(10 + (y - 2018) * 3 + 1) + (10 + (y - 2017) * 3) for y in mar_years]  # 前年9月中間 + 当年3月期末
check("決算月推定: 3月決算を当てる", infer_fiscal_year_end_month(mar_events, mar_years, mar_vals, [1.0, 2.0], today=TODAY), 3)
check("決算月推定: 一致しなければ None", infer_fiscal_year_end_month([], mar_years, mar_vals, [1.0, 2.0], today=TODAY), None)

# マブチ型: EPS は決算短信の遡及修正分しか調整されないので、2026/1/1 の分割は全年度未調整
# （2016〜2023 は 2024 分割のみ調整済み、2024〜2025 も 2026 分割は未調整＝株式数が現在の半分）
m_eps = [40, 40, 45, 45, 45, 38, 45, 75.3, 50.5, 105.9]
m_ni = {y: v for y, v in zip(m_years, [5300, 5300, 6000, 6000, 6000, 5000, 6000, 10000, 6700, 14000])}
new_eps, f = normalize_eps_basis(m_years, m_eps, m_ni, 265_000_000, [1.0, 2.0, 4.0], today=TODAY)
check("EPS: マブチ型の倍率列（純利益÷EPS が全年度で現在の半分の株式数）", f, [2.0] * 10)
check("EPS: マブチ型 2023 は 75.3→37.65、2024 は 50.5→25.25", (new_eps[7], new_eps[8]), (37.65, 25.25))
# 古い年度だけ株式数が 1/4（2つの分割とも未調整）なら 4 で割る
new_eps, f = normalize_eps_basis(m_years, [80, 80, 90, 90, 90, 76, 90, 150.6, 50.5, 105.9],
                                 {**m_ni}, 265_000_000, [1.0, 2.0, 4.0], today=TODAY)
check("EPS: 古い年度だけ2段階未調整 → 4 と 2 が混在", f, [4.0] * 8 + [2.0, 2.0])

# ------------------------------------------------------------------
# 平均株価: 異常値の除外
# ------------------------------------------------------------------
pts = [(f"2025-01-{d:02d}", 1000 + d) for d in range(1, 21)] + [("2025-02-03", 6_728_271_271.5)]
pts += [(f"2026-03-{d:02d}", 2000) for d in range(1, 11)]
avg = yearly_average_prices(close_series(pts), [2025, 2026, 2027])
check("平均株価: 桁違いの終値を除外して平均", avg, [1010.5, 2000.0, 0.0])

# ------------------------------------------------------------------
# 1銘柄まるごと（apply_split_normalization, 通信なし）: 2163 の実データ形
# ------------------------------------------------------------------
entry = {
    "code": "2163", "irbank_enriched": True, "fyEndMonth": 1,
    "years": [2024, 2025, 2026, 2027],
    "netIncomeM": [900, 1060, 1270, 1260],
    "eps": [84.24, 98.99, 118.64, 117.74],
    "epsYears": [2024, 2025, 2026, 2027], "epsValues": [84.24, 98.99, 118.64, 117.74],
    "dividendYears": [2024, 2025, 2026], "dividend": [38, 41, 42],
    "yearlyPrices": [1887.7, 1860.0, 2024.3],
    "currentPrice": 953, "dividendTTM": 64,
}
closes = close_series([(f"2024-06-{d:02d}", 950) for d in range(1, 11)]
                      + [(f"2025-06-{d:02d}", 930) for d in range(1, 11)]
                      + [(f"2026-06-{d:02d}", 1012) for d in range(1, 11)])
ctx = {
    "closes": closes,
    "dividends": [(date(2024, 1, 30), 19), (date(2024, 7, 30), 20), (date(2025, 1, 30), 21),
                  (date(2025, 7, 30), 21), (date(2026, 1, 29), 21), (date(2026, 7, 30), 21.5)],
    "splits": [(date(2026, 8, 1), 2.0)],
    "shares": 21_400_000,
}
s = apply_split_normalization("2163", entry, ctx=ctx, today=TODAY)
check("2163: 配当は Yahoo と一致するので無変更", entry["dividend"], [38, 41, 42])
check("2163: EPS は分割前基準なので ÷2", entry["epsValues"], [42.12, 49.49, 59.32, 58.87])
check("2163: years に揃えた eps 配列も同じ倍率", entry["eps"], [42.12, 49.49, 59.32, 58.87])
check("2163: 平均株価を調整済み終値で引き直し", entry["yearlyPrices"], [950.0, 930.0, 1012.0])
check("2163: 現在利回りは 42÷953", (entry["dividendYield"], entry["dividendYieldBasis"]), (4.41, "annual"))
check("2163: 補正記録", s["eps"], {"2024": 2.0, "2025": 2.0, "2026": 2.0, "2027": 2.0})
# 2回目は何も変わらない（冪等）
s2 = apply_split_normalization("2163", entry, ctx=ctx, today=TODAY)
check("2163: 再実行で二重補正しない", (entry["epsValues"], s2["eps"], s2["dividend"]), ([42.12, 49.49, 59.32, 58.87], {}, {}))
# 補正済みの系列に、新しい分割（前回以降）が効力発生 → Yahoo の株式数がまだ古くても全年度 ÷2
ctx2 = dict(ctx, splits=[(date(2026, 8, 1), 2.0), (date(2027, 2, 1), 2.0)])
s3 = apply_split_normalization("2163", entry, ctx=ctx2, today=date(2027, 2, 2))
check("2163: 前回以降の新規分割は株式数が未更新でも ÷2", entry["epsValues"], [21.06, 24.75, 29.66, 29.43])
s4 = apply_split_normalization("2163", entry, ctx=ctx2, today=date(2027, 2, 3))
check("2163: 新規分割の反映も1回きり", (entry["epsValues"], s4["eps"]), ([21.06, 24.75, 29.66, 29.43], {}))

# ir-bank から取り直した直後（epsBasisNormalized=False）の花王型: 配当は全年度調整済み・EPS未調整
kao = {
    "code": "4452", "irbank_enriched": True, "fyEndMonth": 12,
    "years": [2023, 2024, 2025, 2026], "netIncomeM": [43870, 107600, 118000, 125000],
    "eps": [94.0, 231.9, 260.3, 287.4], "epsYears": [2023, 2024, 2025, 2026], "epsValues": [94.0, 231.9, 260.3, 287.4],
    "dividendYears": [2023, 2024, 2025], "dividend": [75, 76, 77],
    "currentPrice": 3443, "dividendTTM": 78, "epsBasisNormalized": False,
}
kao_ctx = {
    "closes": close_series([("2025-06-02", 6000), ("2025-06-03", 6100)]),
    "dividends": [(date(2023, 6, 29), 37.5), (date(2023, 12, 28), 37.5), (date(2024, 6, 27), 38), (date(2024, 12, 27), 38),
                  (date(2025, 6, 27), 38.5), (date(2025, 12, 29), 38.5), (date(2026, 6, 29), 39)],
    "splits": [(date(2026, 7, 1), 2.0)],
    "shares": 930_000_000,
}
apply_split_normalization("4452", kao, ctx=kao_ctx, today=TODAY)
check("花王: 配当は調整済みなので無変更", kao["dividend"], [75, 76, 77])
check("花王: EPS は ÷2（純利益ベース）", kao["epsValues"], [47.0, 115.95, 130.15, 143.7])
check("花王: 配当に変更が無ければ減配判定も触らない", ("noDividendCut" in kao, "dividendStreak" in kao), (False, False))
check("花王: 補正後は補正済みフラグが立つ", kao.get("epsBasisNormalized"), True)

print("-" * 50)
if failed:
    print(f"❌ {failed}件失敗")
    sys.exit(1)
print("✅ 全件パス")
