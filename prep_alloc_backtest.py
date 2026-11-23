# -*- coding: utf-8 -*-
"""构建"当前配置思路（估值+拥挤代理口径）"十年季度调仓持仓表。
口径：score = 58.3%*估值便宜分(100-PE/PB十年滚动分位) + 41.7%*拥挤安全分(100-配置比例2010以来分位)
拥挤配置比例用披露滞后一季（t-1季度）；估值用≤t的最新年度值。
输出：factor_holdings.json（每季度末 top1/top5/top10 持仓+得分明细）
"""
import openpyxl, datetime as dt, json

SWS = r"C:/Users/Owner/Desktop/周报/【SWSMU】行业比较结果.xlsx"
XZ = r"C:/Users/Owner/Desktop/周报/主动偏股基金数据库_2026Q3.xlsx"
ASOF = dt.date(2026, 11, 13)

def num(v):
    return float(v) if isinstance(v, (int, float)) else None

# ---------- 1. 行业名单与代码（风格行业池：申万一级行业段，剔除综合） ----------
wb = openpyxl.load_workbook(SWS, read_only=True, data_only=True)
rowsst = list(wb["风格行业池"].iter_rows(values_only=True))
ind_code, method = {}, {}
in_l1 = False
for r in rowsst:
    tag = str(r[0]).strip() if r[0] else ""
    if tag == "申万一级行业":
        in_l1 = True; continue
    if tag.startswith("申万二级"):
        in_l1 = False; continue
    if in_l1 and r[1] and r[2]:
        n = str(r[1]).strip()
        if n != "综合":
            ind_code[n] = str(r[2]).strip()
            method[n] = "PB" if str(r[3]).strip() == "PB" else "PE"
assert len(ind_code) == 30, len(ind_code)

rows = list(wb["估值-PE"].iter_rows(values_only=True))

# 年度序列日期头（行索引2，列3起）
hdr_dates = []
for v in rows[2][3:]:
    if isinstance(v, dt.datetime):
        hdr_dates.append(v.date())
    elif isinstance(v, dt.date):
        hdr_dates.append(v)
    else:
        hdr_dates.append(None)
pe_hist = {}
for r in rows[3:]:
    n = str(r[1]).strip() if r[1] else ""
    if n in ind_code:
        pe_hist[n] = {hdr_dates[i]: num(v) for i, v in enumerate(r[3:len(hdr_dates)+3])}

rowspb = list(wb["估值-PB"].iter_rows(values_only=True))
pb_hist = {}
for r in rowspb[3:]:
    n = str(r[1]).strip() if r[1] else ""
    if n in ind_code:
        pb_hist[n] = {hdr_dates[i]: num(v) for i, v in enumerate(r[3:len(hdr_dates)+3])}

# ---------- 2. 兴证配置比例季度序列（L1） ----------
dwb = openpyxl.load_workbook(XZ, read_only=True, data_only=True)
drows = list(dwb["一级行业"].iter_rows(values_only=True))
headers = drows[0]
qdates = [h.date() if isinstance(h, dt.datetime) else h for h in headers[3:]]
qdates = [d for d in qdates if isinstance(d, dt.date)]
alloc = {}
for row in drows:
    if row[2] == "配置比例":
        seen = set()
        for r in drows[drows.index(row)+1:]:
            n = str(r[2]).strip() if r[2] else ""
            if n in ind_code and n not in seen:
                alloc[n] = {qdates[i]: num(v) for i, v in enumerate(r[3:3+len(qdates)])}
                seen.add(n)
        break
missing_a = [n for n in ind_code if n not in alloc]
assert not missing_a, missing_a

def pct_rank(pairs, cur):
    pts = [v for _, v in pairs if v is not None]
    if cur is None or len(pts) < 6:
        return None
    below = sum(1 for v in pts if v < cur)
    eq = sum(1 for v in pts if v == cur)
    return 100.0 * (below + 0.5 * eq) / len(pts)

# ---------- 3. 每季度末打分 ----------
quarters = []
y, q = 2016, 4
while (y, q) <= (2026, 3):
    m, d = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}[q]
    quarters.append(dt.date(y, m, d))
    q += 1
    if q > 4: q, y = 1, y + 1
out = []
for t in quarters:
    rec = {"date": t.isoformat(), "scores": {}}
    for n in ind_code:
        hist = pe_hist[n] if method[n] == "PE" else pb_hist[n]
        pts = sorted((d, v) for d, v in hist.items() if v is not None and d <= t)
        cur_pts = [(d, v) for d, v in pts if d > t.replace(year=t.year-10)]
        cur = cur_pts[-1][1] if cur_pts else None
        vp = pct_rank(cur_pts, cur)
        cheap = None if vp is None else 100 - vp
        # 拥挤：t之前最近一个已披露季度（滞后一季）
        cq = max((d for d in alloc[n] if d <= t and d + dt.timedelta(days=92) <= t), default=None)
        allpts = sorted((d, v) for d, v in alloc[n].items() if v is not None and d <= cq and d >= dt.date(2010,3,31))
        cp = pct_rank(allpts, alloc[n].get(cq))
        safety = None if cp is None else 100 - cp
        score = None
        if cheap is not None and safety is not None:
            score = round(58.3*cheap + 41.7*safety, 1)
        rec["scores"][n] = {"score": score, "cheap": None if cheap is None else round(cheap,1),
                            "safety": None if safety is None else round(safety,1), "method": method[n]}
    out.append(rec)

ranked = []
for rec in out:
    ss = sorted(((v["score"], n) for n, v in rec["scores"].items() if v["score"] is not None), reverse=True)
    ranked.append({"date": rec["date"], "top1": [n for _, n in ss[:1]], "top5": [n for _, n in ss[:5]],
                   "top10": [n for _, n in ss[:10]], "n_scored": len(ss)})
json.dump({"industries": ind_code, "quarters": out, "ranked": ranked},
          open("factor_holdings.json", "w", encoding="utf-8"), ensure_ascii=False)
print("季度数:", len(quarters), quarters[0], "->", quarters[-1])
for r in ranked[:4] + ranked[-2:]:
    print(r["date"], "n=", r["n_scored"], "T1:", r["top1"], "T5:", ",".join(r["top5"]))
