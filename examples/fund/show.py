from daten import extract

ROWS = [
    "令和5年度末現在高",
    "令和6年度*歳出決算額",
    "令和6年度*取崩し額",
    "令和6年度*歳計剰余金処分によるもの",
]

COLUMNS = [
    "積立基金*1財政調整基金",
    "積立基金*2減債基金",
    "積立基金*その他特定目的基金3",
    "積立基金*合計1～3",
    "定額運用基金*1土地開発基金",
    "定額運用基金*その他定額運用基金2",
    "定額運用基金*合計1～2",
]

# Formula results remain unverified caches; dictionary access rejects formulas.
records = extract("proxy-result.xlsx", sheet="97", table="Daten_Table_1")
values = {(record["row"], record["column"]): record["value"] for record in records}
for row in ROWS:
    print(row)
    for column in COLUMNS:
        print(f"  {column}: {values[row, column]}")
    inputs = [values[row, column] for column in COLUMNS[0:3] + COLUMNS[4:6]]
    if any(value is None for value in inputs):
        print("  合計: 確認不能（入力に空セルがあります）")
        continue
    fund = sum(values[row, column] for column in COLUMNS[0:3])
    fixed = sum(values[row, column] for column in COLUMNS[4:6])
    ok = values[row, COLUMNS[3]] == fund and values[row, COLUMNS[6]] == fixed
    print("  合計:", "期待値と一致（キャッシュは未検証）" if ok else f"未再計算（期待値 {fund} / {fixed}）")
