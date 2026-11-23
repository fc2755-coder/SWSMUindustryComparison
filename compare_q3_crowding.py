from pathlib import Path
import json
import re

ROOT = Path(r"C:/Users/Owner/WorkBuddy/2026-10-26-13-16-06")
BASELINE_HTML = ROOT / "baseline_h1_crowding.html"
CURRENT_HTML = Path(r"C:/Users/Owner/Desktop/行业比较与个股筛选模型_20261025.html")

OLD_BLOCK = r'''# 2026H1主动偏股基金全持仓基线，仅用于公募更新影响归因。
hwb = openpyxl.load_workbook(HOLDINGS_SOURCE, data_only=True, read_only=True)
assert "26H1" in hwb.sheetnames
sub_map = {s["code"]: s["subindustry"] for s in stocks if s["subindustry"]}
current_holdings = {}
hs = hwb["26H1"]
for row in hs.iter_rows(min_row=3, max_col=13, values_only=True):
    code = str(row[1] or "")
    if not code.endswith((".SH", ".SZ", ".BJ")):
        continue
    weight = num(row[11]) or 0.0
    current_holdings[code] = {
        "fundCount": num(row[3]) or 0.0,
        "fundCompanies": num(row[4]) or 0.0,
        "freeFloat": num(row[8]) or 0.0,
        "holdingValue": (num(row[9]) or 0.0) / 10000,
        "fundWeight": weight,
        "floatMarketCap": (num(row[9]) / (num(row[7]) / 100) / 10000) if num(row[9]) is not None and num(row[7]) is not None and num(row[7]) > 0 else None,
        "l1": str(row[12] or ""),
        "l2": sub_map.get(code, ""),
    }

def build_overweight(level):
    rows = []
    for code, h in current_holdings.items():
        group = h.get("l1" if level == "L1" else "l2")
        cap = h.get("floatMarketCap")
        if group and cap is not None and cap > 0:
            rows.append((code, h, group))
    total_weight = sum(h["fundWeight"] for _, h, _ in rows)
    total_cap = sum(h["floatMarketCap"] for _, h, _ in rows)
    group_weight, group_cap = {}, {}
    for _, h, group in rows:
        group_weight[group] = group_weight.get(group, 0.0) + h["fundWeight"]
        group_cap[group] = group_cap.get(group, 0.0) + h["floatMarketCap"]
    result = {}
    for group in group_cap:
        fund_a = 100 * group_weight.get(group, 0.0) / total_weight if total_weight > 0 else None
        market = 100 * group_cap[group] / total_cap if total_cap > 0 else None
        result[group] = {"fundWeightA": fund_a, "marketCapPct": market, "overweight": fund_a - market if fund_a is not None and market is not None else None}
    if level == "L1":
        for _, h, _ in rows:
            h["fundWeightA"] = 100 * h["fundWeight"] / total_weight if total_weight > 0 else None
            h["marketCapPct"] = 100 * h["floatMarketCap"] / total_cap if total_cap > 0 else None
            h["overweight"] = h["fundWeightA"] - h["marketCapPct"] if h["fundWeightA"] is not None and h["marketCapPct"] is not None else None
    return result

industry_overweight = {"L1": build_overweight("L1"), "L2": build_overweight("L2")}
for s in stocks:
    s.update({"fundCount": 0.0, "fundCompanies": 0.0, "freeFloat": 0.0, "holdingValue": 0.0, "fundWeight": 0.0, "fundWeightA": None, "marketCapPct": None, "overweight": None})
    s.update(current_holdings.get(s["code"], {}))
    s["hasFundHolding"] = s["code"] in current_holdings
for key in ("overweight", "fundWeightA", "fundCount", "freeFloat"):
    cross_scores(stocks, key, higher=True, missing=0)
for s in stocks:
    s["crowdRaw"] = 0.40 * s["overweightScore"] + 0.20 * s["fundCountScore"] + 0.20 * s["fundWeightAScore"] + 0.20 * s["freeFloatScore"]
    s["crowdScore"] = 100 - s["crowdRaw"]
for level in ("L1", "L2"):
    group = [x for x in industries if x["level"] == level]
    for x in group:
        x.update(industry_overweight[level].get(x["name"], {"fundWeightA": None, "marketCapPct": None, "overweight": None}))
    for key in ("overweight", "fundWeightA"):
        cross_scores(group, key, higher=True, missing=0)
    for x in group:
        x["crowdRaw"] = 0.70 * x["overweightScore"] + 0.30 * x["fundWeightAScore"]
        x["crowdScore"] = 100 - x["crowdRaw"]
'''

source = (ROOT / "build_scoring_model.py").read_text(encoding="utf-8")
start = source.index("# 2026Q3主动偏股基金前十大重仓")
end = source.index("for x in stocks + industries:", start)
old_source = source[:start] + OLD_BLOCK + "\n" + source[end:]
old_source = old_source.replace('OUTPUT = Path(r"C:/Users/Owner/Desktop/行业比较与个股筛选模型_20261025.html")', f'OUTPUT = Path(r"{BASELINE_HTML.as_posix()}")')
old_source = old_source.replace('HOLDINGS_Q3_SOURCE = Path(r"C:/Users/Owner/Desktop/重仓持股26Q3.xlsx")', 'HOLDINGS_SOURCE = Path(r"C:/Users/Owner/Desktop/主动偏股基金全持仓数据_26H1.xlsx")')
old_source = old_source.replace('    x["crowdChange"] = round(x["crowdRaw"] - x["prevCrowdRaw"], 2)\n    x["overweightChange"] = round(x["overweight"] - x["prevOverweight"], 6) if x["overweight"] is not None and x["prevOverweight"] is not None else None\n    x["fundWeightChange"] = round(x["fundWeightA"] - x["prevFundWeightA"], 6) if x["fundWeightA"] is not None and x["prevFundWeightA"] is not None else None\n', '    x["crowdChange"] = None\n    x["overweightChange"] = None\n    x["fundWeightChange"] = None\n')
exec(compile(old_source, str(ROOT / "baseline_h1_model.py"), "exec"), {"__name__": "__main__"})

def load_data(path):
    text = path.read_text(encoding="utf-8")
    return json.loads(re.search(r'<script id="modelData" type="application/json">(.*?)</script>', text, re.S).group(1))

before = load_data(BASELINE_HTML)
after = load_data(CURRENT_HTML)
comparison = {"before": before, "after": after}
(ROOT / "q3_crowding_comparison.json").write_text(json.dumps(comparison, ensure_ascii=False), encoding="utf-8")
print("BASELINE_AND_CURRENT_READY", len(before["stocks"]), len(after["stocks"]))
