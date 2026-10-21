"""ir-bank /results ページのパース自己テスト

  python scripts/test_irbank_parse.py

外部通信はしない（requests.get を差し替えて固定HTMLを返す）。
銀行など営業利益の列が無い業態でも、経常利益を利益として業績を取り込めることを確認する。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import update_data  # noqa: E402


def _table(header, rows):
    tr = lambda cells: "<tr>" + "".join(f"<td>{c}</td>" for c in cells) + "</tr>"
    return "<table>" + tr(header) + "".join(tr(r) for r in rows) + "</table>"


def _page(fin_header, fin_rows, div_rows):
    return ("<html><body>" + _table(fin_header, fin_rows)
            + _table(["年度", "一株配当", "配当性向"], div_rows) + "</body></html>")


YEARS = ["2021/03", "2022/03", "2023/03", "2024/03", "2025/03"]
DIVS = [[y, d, "40%"] for y, d in zip(YEARS, ["25", "28", "32", "41", "64"])]

CASES = [
    ("一般企業（営利あり）",
     _page(["年度", "売上", "営利", "経常", "純利", "EPS"],
           [[y, f"{1000 + i}億", f"{100 + i}億", f"{110 + i}億", f"{70 + i}億", f"{50 + i}.0"]
            for i, y in enumerate(YEARS)], DIVS),
     ([1000, 1001, 1002, 1003, 1004], [100, 101, 102, 103, 104])),
    ("銀行（営利なし→経常利益で代用）",
     _page(["年度", "経常収益", "経常利益", "純利", "EPS"],
           [[y, f"{6 + i}兆", f"{1 + i}兆", f"{8000 + i}億", f"{80 + i}.0"]
            for i, y in enumerate(YEARS)], DIVS),
     ([60000, 70000, 80000, 90000, 100000], [10000, 20000, 30000, 40000, 50000])),
]


class _Resp:
    status_code = 200

    def __init__(self, text):
        self.text = text


def main():
    failed = 0
    orig = update_data.requests.get
    try:
        for name, html, (exp_rev, exp_prof) in CASES:
            update_data.requests.get = lambda *a, _h=html, **k: _Resp(_h)
            r = update_data.fetch_irbank_financials("0000")
            got = (r["revenue"], r["profit"], r["dividend"], len(r["epsValues"])) if r else None
            exp = (exp_rev, exp_prof, [25, 28, 32, 41, 64], 5)
            ok = got == exp
            failed += not ok
            print(f"{'OK ' if ok else 'NG '} {name} -> {got} (期待 {exp})")
    finally:
        update_data.requests.get = orig
    print("-" * 50)
    if failed:
        print(f"❌ {failed}件失敗")
        return 1
    print(f"✅ 全{len(CASES)}件パス")
    return 0


if __name__ == "__main__":
    sys.exit(main())
