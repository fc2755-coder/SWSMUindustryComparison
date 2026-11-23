import bisect
import csv
import json
import math
import statistics
from datetime import datetime
from pathlib import Path
import openpyxl

SOURCE = Path(r"C:/Users/Owner/Desktop/多维数据比较1120.xlsx")
HOLDINGS_Q3_SOURCE = Path(r"C:/Users/Owner/Desktop/周报/重仓持股26Q3.xlsx")
HOLDINGS_Q2_SOURCE = Path(r"C:/Users/Owner/Desktop/每周更新/二季度机构重仓股持仓汇总.xlsx")  # 原重仓持股26Q2.xlsx已删除，此为同源替代（2026二季，表头结构一致）
HOLDINGS_DATABASE = Path(r"C:/Users/Owner/Desktop/周报/主动偏股基金数据库_2026Q3.xlsx")
RECENT_IPO_SOURCE = Path(r"C:/Users/Owner/WorkBuddy/2026-10-26-13-16-06/recent_ipos_20261120.csv")
OUTPUT = Path(r"C:/Users/Owner/Desktop/周报/行业比较与个股筛选模型_20261120.html")
ASOF = datetime(2026, 11, 20)
PB_INDUSTRIES = {"银行", "非银金融", "房地产"}


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(float(v)) else None


def pct(v):
    v = num(v)
    if v is None:
        return None
    return v * 100 if abs(v) <= 1.5 else v


def parse_date(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str) and v.strip():
        try:
            return datetime.strptime(v.strip()[:10], "%Y-%m-%d")
        except ValueError:
            return None
    return None


def add_two_months(d):
    m = d.month + 2
    return datetime(d.year + (1 if m > 12 else 0), m - 12 if m > 12 else m, min(d.day, 28))


def clamp(v):
    return max(0, min(100, v))


def cross_scores(items, key, higher=True, missing=50):
    values = sorted(x[key] for x in items if x.get(key) is not None)
    n = len(values)
    for x in items:
        v = x.get(key)
        if v is None or n == 0:
            x[key + "Score"] = missing
            continue
        if n == 1:
            x[key + "Score"] = 50
            continue
        left = bisect.bisect_left(values, v)
        right = bisect.bisect_right(values, v)
        rank = (left + right - 1) / 2
        score = 100 * rank / (n - 1)
        x[key + "Score"] = score if higher else 100 - score


def minmax_scores(items, key, missing=0):
    # 层级内线性缩放0—100：最大值100、最小值0，不做排名分位，保留原始差距。
    values = [x[key] for x in items if x.get(key) is not None]
    lo, hi = (min(values), max(values)) if values else (0.0, 0.0)
    for x in items:
        v = x.get(key)
        if v is None or hi <= lo:
            x[key + "Score"] = missing
        else:
            x[key + "Score"] = (v - lo) / (hi - lo) * 100


def own_level_score(values, current):
    vals = [num(v) for v in values if num(v) is not None]
    if current is None or len(vals) < 2:
        return 50
    less = sum(v < current for v in vals)
    equal = sum(v == current for v in vals)
    return clamp(100 * (less + (equal - 1) / 2) / (len(vals) - 1))


def roe_cycle(values):
    vals = [num(x) for x in values]
    r25, r26, r27 = vals[4], vals[5], vals[6]
    if r26 is None or r27 is None:
        return "数据不足", None
    delta = r27 - r26
    valid = [x for x in vals if x is not None]
    if delta < -0.1:
        label = "拐头向下"
    elif delta > 0.1:
        history = [x for x in vals[:6] if x is not None]
        trough = min(history) if history else r26
        if abs(r26 - trough) <= 0.5 or (r25 is not None and abs(r25 - trough) <= 0.5):
            label = "底部反弹"
        elif r25 is not None and r26 > r25 + 0.1:
            label = "持续向上"
        else:
            label = "修复向上"
    else:
        label = "高位平稳" if valid and r27 >= statistics.median(valid) else "低位平稳"
    return label, delta


def revision(cur, prev):
    cur, prev = num(cur), num(prev)
    if cur is None or prev is None:
        return None
    return (cur - prev) / abs(prev) * 100 if abs(prev) > 1e-12 else None


def weekly_streak(latest, wk1, wk2, wk3):
    # 尾部连续周度上修次数：最新>上周计1，再上周>上上周计2，再上上周>上上上周计3；任一缺失即中断。
    vals = [num(latest), num(wk1), num(wk2), num(wk3)]
    if any(v is None for v in vals):
        return 0
    streak = 0
    for a, b in ((0, 1), (1, 2), (2, 3)):
        if vals[a] > vals[b]:
            streak += 1
        else:
            break
    return streak


def digest(pe26, pe27, pe28):
    pes = [pe26, pe27, pe28]
    if all(v is None or v <= 0 for v in pes):
        return {"digestYear": None, "digestScore": 0, "digestLabel": "PE不适用", "digestRisk": "高"}
    for year, v, score, risk in zip((2026, 2027, 2028), pes, (100, 70, 30), ("低", "中", "高")):
        if v is not None and 0 < v < 30:
            return {"digestYear": year, "digestScore": score, "digestLabel": f"{str(year)[2:]}E已≤30x", "digestRisk": risk}
    return {"digestYear": None, "digestScore": 0, "digestLabel": "28E仍未消化", "digestRisk": "高"}


wb = openpyxl.load_workbook(SOURCE, data_only=True, read_only=True)
ws = wb["多维比较"]
# 只读流式读取，避免非只读模式下工作簿对象图占用过多内存。
head_rows = []
for row in ws.iter_rows(min_row=1, max_row=184, max_col=68, values_only=True):
    head_rows.append(row)
assert len(head_rows[0]) >= 68 and parse_date(head_rows[0][0]) == ASOF, "源表列数或数据截面与模型日期不一致"
assert head_rows[2][19] == "27E EPS（最新）" and head_rows[2][30] == "当前PE分位数"


def unpack(row, level, parent=None):
    code, name = str(row[0] or ""), str(row[1] or "")
    industry = parent or (str(row[2] or "") if level == "stock" else name)
    # 68列周度结构（1031版起）：EPS/ROE均为周度序列（最新/上周/上上周/上上上周），上修口径=周度（最新 vs 上周）。
    eps = [num(row[i]) for i in (10, 11, 12, 13, 14, 15, 19, 23)]
    roe = [num(row[i]) for i in (34, 35, 36, 37, 38, 39, 41, 45)]
    cycle, roe_delta = roe_cycle(roe[:7])
    method = "PB" if industry in PB_INDUSTRIES else "PE"
    pe_pct, pb_pct = num(row[30]), num(row[49])
    val_pct = pb_pct if method == "PB" else pe_pct
    pe26, pe27, pe28 = num(row[27]), num(row[28]), num(row[29])
    val_multiple = num(row[47]) if method == "PB" else pe27
    unlock_date = parse_date(row[65])
    unlock_ratio = num(row[67])
    unlock_amount = num(row[66]) * num(row[5]) / 10000 if num(row[66]) is not None and num(row[5]) is not None else None
    growth27 = num(row[25]) if num(row[15]) is not None and num(row[15]) > 0 else None
    growth_all = [num(row[24]) if num(row[14]) is not None and num(row[14]) > 0 else None,
                  growth27,
                  num(row[26]) if num(row[16]) is not None and num(row[16]) > 0 else None]
    item = {
        "code": code, "name": name, "level": level, "industry": industry, "subindustry": str(row[3] or ""),
        "marketCap": num(row[4]), "week": num(row[6]), "month": num(row[7]), "ytd": num(row[8]), "since924": num(row[9]),
        "eps": eps, "eps27": eps[6], "epsGrowth": growth27, "epsGrowthAll": growth_all, "epsRev": revision(row[19], row[20]),
        "epsStreak": weekly_streak(row[19], row[20], row[21], row[22]),
        "epsLevelRaw": own_level_score(eps[:7], eps[6]),
        "roe": roe, "roe26": roe[5], "roe27": roe[6], "roeDelta": roe_delta,
        "roeRev": roe[6] - num(row[42]) if roe[6] is not None and num(row[42]) is not None else None,
        "roeStreak": weekly_streak(row[41], row[42], row[43], row[44]),
        "cycle": cycle, "method": method, "valPct": val_pct, "pePct": pe_pct, "pbPct": pb_pct,
        "pe26": pe26, "pe27": pe27, "pe28": pe28, "pb27": num(row[47]), "valMultiple": val_multiple,
        "valPass": val_pct is not None and 0 <= val_pct < 50 and val_multiple is not None and val_multiple > 0,
        "staticYield": num(row[53]), "yield27": pct(row[62]), "payouts": [pct(row[55]), pct(row[56]), pct(row[57])],
        "stable": str(row[60] or ""), "coverage": num(row[64]), "unlockDate": unlock_date.strftime("%Y-%m-%d") if unlock_date else "",
        "unlockRatio": unlock_ratio, "unlockAmount": unlock_amount,
        "unlockAlert": bool(unlock_date and ASOF <= unlock_date <= add_two_months(ASOF) and unlock_ratio is not None and unlock_ratio > 10)
    }
    item["coverage"] = item["coverage"] * 100 if item["coverage"] is not None else None
    item.update(digest(pe26, pe27, pe28))
    return item


industries = [unpack(list(head_rows[r - 1]), "L1" if r <= 50 else "L2") for r in range(20, 185)]
# 剔除一级"综合"行业：行业间可比性差（成分混杂），不参与任何行业口径评分与展示
industries = [x for x in industries if not (x["level"] == "L1" and x["name"] == "综合")]
stocks = []
for row in ws.iter_rows(min_row=185, max_col=68, values_only=True):
    if str(row[0] or "").endswith((".SH", ".SZ", ".BJ")):
        stocks.append(unpack(list(row), "stock"))
del head_rows
wb.close()

# iFinD新股上市日期：剔除截至2026-11-13上市未满一年的标的，避免短历史估值分位失真。
recent_ipos = {}
with RECENT_IPO_SOURCE.open("r", encoding="utf-8-sig", newline="") as f:
    for row in csv.DictReader(f):
        d = datetime.strptime(row["新股上市日期"], "%Y%m%d")
        if datetime(2025, 11, 20) < d <= ASOF:
            recent_ipos[row["股票代码"]] = d
for s in stocks:
    d = recent_ipos.get(s["code"])
    s["listingDate"] = d.strftime("%Y-%m-%d") if d else ""
    s["recentIPO"] = d is not None

sub_parent = {s["subindustry"]: s["industry"] for s in stocks if s["subindustry"]}
for x in industries:
    if x["level"] == "L2":
        parent = sub_parent.get(x["name"], "")
        x["parent"] = parent
        x["industry"] = parent or x["name"]
        x["method"] = "PB" if x["industry"] in PB_INDUSTRIES else "PE"
        x["valPct"] = x["pbPct"] if x["method"] == "PB" else x["pePct"]
        x["valMultiple"] = x["pb27"] if x["method"] == "PB" else x["pe27"]
        x["valPass"] = x["valPct"] is not None and 0 <= x["valPct"] < 50 and x["valMultiple"] is not None and x["valMultiple"] > 0
    else:
        x["parent"] = x["name"]

industry_yield = {x["name"]: x["staticYield"] for x in industries if x["level"] == "L1"}
for s in stocks:
    trig = (s["staticYield"] is not None and s["staticYield"] > 4) or (industry_yield.get(s["industry"]) or 0) > 4
    s["divTrigger"] = trig
    s["stablePass"] = s["stable"] == "稳定"
    s["coveragePass"] = s["coverage"] >= 200 if s["coverage"] is not None and s["coverage"] > 0 else None
    s["yieldPass"] = s["yield27"] is not None and s["yield27"] > 3.5
    s["divPass"] = (s["yieldPass"] and s["stablePass"] and s["coveragePass"] is True) if trig else True
    s["divStatus"] = "非红利约束" if not trig else ("通过" if s["divPass"] else ("未通过" if not s["yieldPass"] else ("待补数据" if s["coveragePass"] is None else "未通过")))


# ---------- 红利资产筛选策略 v2：估值走廊框架 + 红马增强标签 ----------
# 框架来源：券商策略估值走廊研究框架及配套自动更新底稿(2026-11)；增强参照：红利"红马"选股思路。
def ddm_theory_pe(roe1, d1, disc=0.07, t1=5, t2=10, decay=0.02, d_step=0.1, d1_cap=0.9,
                  g3=0.02, d3_floor=0.7, d3_cap=0.95, g1_cap=0.065, g2_cap=0.04):
    """三阶段DDM理论PE（EPS1=1）。阶段规则沿用估值走廊底稿：阶段2 ROE衰减2pct、分红率+10pct(上限0.9)；
    阶段3永续，分红率钳制在0.7-0.95，g=2%。g_t = ROE_t × (1 - d_t)。"""
    d1 = min(d1, d1_cap)
    roe2, d2 = roe1 - decay, min(d1 + d_step, d1_cap)
    roe3, d3 = roe2 - decay, min(max(d2 + d_step, d3_floor), d3_cap)
    # g外推上限：框架案例库99只资产的g1最大6.4%（高留存成长股的留存收益不可按DDM无限外推）
    g1, g2 = min(roe1 * (1 - d1), g1_cap), min(roe2 * (1 - d2), g2_cap)
    pe, eps, t = 0.0, 1.0, 0
    for rr, dd, gg, tt, last in [(roe1, d1, g1, t1, False), (roe2, d2, g2, t2, False), (roe3, d3, g3, None, True)]:
        for _ in range(tt or 100):
            t += 1
            pe += eps * dd / (1 + disc) ** t
            eps *= (1 + gg)
            if last:
                break
        if last:
            pe += (eps * dd) / (1 + disc) ** t / (disc - g3)
    return pe

def pct_rank(vals, v):
    """v 在 vals 中的横截面分位(0-100)；vals 需已剔除 None。"""
    clean = [x for x in vals if x is not None]
    if not clean or v is None:
        return None
    return 100.0 * sum(1 for x in clean if x <= v) / len(clean)

def build_dividend_strategy(stocks):
    """估值走廊框架红利选股：
    初筛（价值资产池）：ROE27E≥7% + ROE变异系数≤全A 30%分位 + 静态与27E股息率均≥2% + 分红稳定
                      + OCF/DPS≥200% + 市值≥100亿 + 非次新；
    定价：三阶段DDM理论PE/PB → 估值走廊[0.8,1.2]×理论PB → 当前PB折溢价；
    补偿：应有股息率=3.1%+0.3×(ROE标准差-1.1pp)（水电基准），超额股息率=实际-应有；
    评分：估值折价35 + 股息安全边际25 + 盈利质量稳定20 + 分红可持续20；
    标签：红马增强（长江）=26E/27E/28E EPS同比均>-5%（成长护航）且CV≤全A中位（盈利稳态）。"""
    # 全A ROE变异系数分布（框架口径：与全A个股比）
    cvs_all = []
    for s in stocks:
        seq = [x for x in (s["roe"][:7] if s.get("roe") else []) if x is not None]
        if len(seq) >= 5:
            m = sum(seq) / len(seq)
            if m > 0:
                s["_roeCV"] = (sum((x - m) ** 2 for x in seq) / len(seq)) ** 0.5 / m
                s["_roeSD"] = (sum((x - m) ** 2 for x in seq) / len(seq)) ** 0.5
                cvs_all.append(s["_roeCV"])
    cvs_all.sort()
    cv_q30 = cvs_all[int(0.30 * (len(cvs_all) - 1))] if cvs_all else 0.5
    cv_med = cvs_all[len(cvs_all) // 2] if cvs_all else 0.5
    avg_payout = lambda s: (lambda ps: sum(ps) / len(ps) if ps else None)([p for p in s["payouts"] if p is not None and p > 0])
    pool = []
    for s in stocks:
        if s["recentIPO"] or (s["marketCap"] or 0) < 100:
            continue
        roe27, cv = s["roe27"], s.get("_roeCV")
        y_ttm, y27, pay, cov = s["staticYield"], s["yield27"], avg_payout(s), s["coverage"]
        if not (roe27 is not None and roe27 >= 7 and cv is not None and cv <= cv_q30):
            continue
        if not (y_ttm is not None and y_ttm >= 2 and y27 is not None and y27 >= 2):
            continue
        if s["stablePass"] is not True:
            continue
        if cov is not None and cov < 150:
            continue
        pb27 = s["pb27"]
        pe27 = s["pe27"]
        if pb27 is None and not (pe27 is not None and pe27 > 0):
            continue
        d1 = (pay / 100) if (pay is not None and 0 < pay <= 150) else 0.5  # payouts为百分数口径
        tp = ddm_theory_pe(roe27 / 100, d1)
        theory_pb = tp * roe27 / 100
        ratio = (pb27 / theory_pb) if (pb27 and theory_pb > 0) else ((pe27 / tp) if (pe27 and tp > 0) else None)
        sd = s.get("_roeSD") or 1.1
        req_yield = 3.1 + 0.3 * (sd - 1.1)
        pool.append({
            "code": s["code"], "name": s["name"], "industry": s["industry"],
            "roe27": roe27, "roeCV": round(cv, 3), "roeSD": round(sd, 2),
            "yieldTTM": round(y_ttm, 2), "yield27": round(y27, 2), "reqYield": round(req_yield, 2),
            "excessYield": round(y27 - req_yield, 2),
            "payout": round((d1 * 100) if d1 < 1.5 else d1, 1),
            "coverage": cov, "theoryPE": round(tp, 2), "theoryPB": round(theory_pb, 3),
            "pb27": pb27, "pe27": pe27, "ratio": round(ratio, 3) if ratio else None,
            "premium": round((ratio - 1) * 100, 1) if ratio else None,
            "corridor": ("低于下沿" if ratio < 0.8 else "走廊内" if ratio <= 1.2 else "高于上沿") if ratio else "—",
            "growthTag": all(g is not None and g > -5 for g in s.get("epsGrowthAll", [None, None, None])),
        })
    # 评分（仅池内横截面）
    if pool:
        r_val = [x["ratio"] for x in pool if x["ratio"]]
        r_exy = [x["excessYield"] for x in pool]
        r_roe = [x["roe27"] for x in pool]
        r_cv = [x["roeCV"] for x in pool]
        r_cov = [x["coverage"] for x in pool if x["coverage"] is not None]
        r_pay = [x["payout"] for x in pool]
        for x in pool:
            val_score = pct_rank(r_val, x["ratio"])
            val_score = 100 - val_score if val_score is not None else 50
            exy_score = pct_rank(r_exy, x["excessYield"]) or 50
            roe_s = pct_rank(r_roe, x["roe27"]) or 50
            cv_s = 100 - (pct_rank(r_cv, x["roeCV"]) or 50)
            cov_s = pct_rank(r_cov, x["coverage"]) if (x["coverage"] is not None and r_cov) else 50
            pay_s = pct_rank(r_pay, x["payout"]) or 50
            x["valScore"] = round(val_score, 1)
            x["yieldScore"] = round(exy_score, 1)
            x["qualityScore"] = round(0.5 * roe_s + 0.5 * cv_s, 1)
            x["sustainScore"] = round(0.5 * (cov_s or 50) + 0.3 * pay_s + 0.2 * (100 if x["coverage"] is not None else 50), 1)
            x["score"] = round(0.35 * val_score + 0.25 * exy_score + 0.20 * (0.5 * roe_s + 0.5 * cv_s) + 0.20 * x["sustainScore"], 1)
            x["redHorse"] = bool(x["growthTag"] and x["roeCV"] <= cv_med)
        pool.sort(key=lambda z: -z["score"])
        for i, x in enumerate(pool, 1):
            x["rank"] = i
    return {
        "params": {"disc": 0.07, "t1": 5, "t2": 10, "decay": 0.02, "g3": 0.02,
                    "reqBase": 3.1, "reqSlope": 0.3, "reqBaseSD": 1.1,
                    "corridor": [0.8, 1.2], "cvQ30": round(cv_q30, 3), "cvMedian": round(cv_med, 3)},
        "poolCount": len(pool),
        "belowFloor": sum(1 for x in pool if x["corridor"] == "低于下沿"),
        "inCorridor": sum(1 for x in pool if x["corridor"] == "走廊内"),
        "aboveCap": sum(1 for x in pool if x["corridor"] == "高于上沿"),
        "redHorseN": sum(1 for x in pool if x["redHorse"]),
        "list": pool,
    }

DIV_STRATEGY = build_dividend_strategy(stocks)

# 2026Q3主动偏股基金前十大重仓：Q3决定最新得分，Q2仅作同口径环比。
# 个股取重仓持股表；行业取数据库的配置与超配汇总，不与26H1全持仓混比。
sub_map = {s["code"]: s["subindustry"] for s in stocks if s["subindustry"]}
stock_industry_map = {s["code"]: s["industry"] for s in stocks}


def load_heavy_holdings(path):
    book = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = book["sheet1"]
    header_rows = list(sheet.iter_rows(min_row=2, max_row=2, max_col=12, values_only=True))
    headers = {str(v): i for i, v in enumerate(header_rows[0]) if v is not None}
    required = ["代码", "持有基金数", "持有公司家数", "持股占流通股比(%)", "持股总市值(万元)", "持股市值占基金股票投资市值比(%)"]
    assert all(k in headers for k in required), (path, headers)
    holdings = {}
    for row in sheet.iter_rows(min_row=3, max_col=12, values_only=True):
        code = str(row[headers["代码"]] or "")
        if not code.endswith((".SH", ".SZ", ".BJ")):
            continue
        value_wan = num(row[headers["持股总市值(万元)"]])
        float_pct = num(row[headers["持股占流通股比(%)"]])
        holdings[code] = {
            "fundCount": num(row[headers["持有基金数"]]) or 0.0,
            "fundCompanies": num(row[headers["持有公司家数"]]) or 0.0,
            "freeFloat": float_pct or 0.0,
            "holdingValue": (value_wan or 0.0) / 10000,
            "fundWeight": num(row[headers["持股市值占基金股票投资市值比(%)"]]) or 0.0,
            "floatMarketCap": value_wan / (float_pct / 100) / 10000 if value_wan is not None and float_pct is not None and float_pct > 0 else None,
            "l1": str(row[headers["所属行业"]] or "") or stock_industry_map.get(code, ""),
            "l2": sub_map.get(code, ""),
        }
    return holdings


def build_stock_overweight(holdings):
    rows = [(code, h) for code, h in holdings.items() if h.get("floatMarketCap") is not None and h["floatMarketCap"] > 0]
    total_weight = sum(h["fundWeight"] for _, h in rows)
    total_cap = sum(h["floatMarketCap"] for _, h in rows)
    for _, h in rows:
        h["fundWeightA"] = 100 * h["fundWeight"] / total_weight if total_weight > 0 else None
        h["marketCapPct"] = 100 * h["floatMarketCap"] / total_cap if total_cap > 0 else None
        h["overweight"] = h["fundWeightA"] - h["marketCapPct"] if h["fundWeightA"] is not None and h["marketCapPct"] is not None else None


current_holdings = load_heavy_holdings(HOLDINGS_Q3_SOURCE)
previous_holdings = load_heavy_holdings(HOLDINGS_Q2_SOURCE)
build_stock_overweight(current_holdings)
build_stock_overweight(previous_holdings)

# 数据库直接提供一级、二级行业的配置比例和超配比例；单位由小数转换为百分点。
dwb = openpyxl.load_workbook(HOLDINGS_DATABASE, data_only=True, read_only=True)


def read_industry_holdings(sheet_name, level):
    rows = list(dwb[sheet_name].iter_rows(values_only=True))
    headers = rows[0]
    q2_col = headers.index(datetime(2026, 6, 30))
    q3_col = headers.index(datetime(2026, 9, 30))
    names = {x["name"] for x in industries if x["level"] == level}
    result = {name: {} for name in names}
    marker_col = 2 if level == "L1" else 1
    for metric, marker in (("fundWeightA", "配置比例"), ("overweight", "超配比例")):
        start = next(i for i, row in enumerate(rows) if row[marker_col] == marker)
        seen = set()
        for row in rows[start + 1:]:
            name = row[2]
            if name not in names or name in seen:
                continue
            q3, q2 = num(row[q3_col]), num(row[q2_col])
            result[name][metric] = q3 * 100 if q3 is not None else None
            result[name]["prev" + metric[0].upper() + metric[1:]] = q2 * 100 if q2 is not None else None
            if metric == "fundWeightA":
                # 配置比例季度历史（2004Q1起，数据库一/二级行业sheet前段），用于2010年以来历史分位
                result[name]["allocHistDates"] = [str(headers[i])[:10] for i in range(3, q3_col + 1)]
                result[name]["allocHistVals"] = [num(v) * 100 if num(v) is not None else None for v in row[3:q3_col + 1]]
                assert len(result[name]["allocHistDates"]) == len(result[name]["allocHistVals"])
            seen.add(name)
            if len(seen) == len(names):
                break
        assert seen == names, (sheet_name, metric, names - seen)
    return result


def hist_percentile(hist, cur, start="2010-03-31", end=None):
    """配置比例在2010(可配)年以来序列中的分位（0-100，平均秩法）"""
    dates, vals = hist
    pts = [v for d, v in zip(dates, vals) if v is not None and d >= start and (end is None or d <= end)]
    if cur is None or len(pts) < 10:
        return None
    below = sum(1 for v in pts if v < cur)
    equal = sum(1 for v in pts if v == cur)
    return round(100 * (below + 0.5 * equal) / len(pts), 1)


industry_db = {"L1": read_industry_holdings("一级行业", "L1"), "L2": read_industry_holdings("二级行业", "L2")}

# 仓位周报动态数据（来源：公募基金行业仓位周报2026-11-13，占已披露A股前十大重仓股市值比例；主动变化为模型估算）
POS_WEEK_SOURCE = Path(__file__).parent / "posweek_20261120.json"
pos_week = {}
if POS_WEEK_SOURCE.exists():
    import json as _json
    _pw = _json.loads(POS_WEEK_SOURCE.read_text(encoding="utf-8"))
    pos_week = {w["name"]: {"posCurrent": round(w["current"], 2), "posActive": round(w["active_delta"], 2),
                            "posWeek": round(w["week_active"], 2)} for w in _pw["weekly"]}


for s in stocks:
    s.update({"fundCount": 0.0, "fundCompanies": 0.0, "freeFloat": 0.0, "holdingValue": 0.0, "fundWeight": 0.0, "fundWeightA": None, "marketCapPct": None, "overweight": None})
    s.update(current_holdings.get(s["code"], {}))
    s["hasFundHolding"] = s["code"] in current_holdings
    prev = previous_holdings.get(s["code"], {})
    s["hasPrevFundHolding"] = s["code"] in previous_holdings
    s["prevFundCount"] = prev.get("fundCount", 0.0)
    s["prevFreeFloat"] = prev.get("freeFloat", 0.0)
    s["prevFundWeightA"] = prev.get("fundWeightA")
    s["prevMarketCapPct"] = prev.get("marketCapPct")
    s["prevOverweight"] = prev.get("overweight")
for key in ("overweight", "fundWeightA", "fundCount", "freeFloat", "prevOverweight", "prevFundWeightA", "prevFundCount", "prevFreeFloat"):
    cross_scores(stocks, key, higher=True, missing=0)
for s in stocks:
    s["crowdRaw"] = 0.40 * s["overweightScore"] + 0.20 * s["fundCountScore"] + 0.20 * s["fundWeightAScore"] + 0.20 * s["freeFloatScore"]
    s["crowdScore"] = 100 - s["crowdRaw"]
    s["prevCrowdRaw"] = round(0.40 * s["prevOverweightScore"] + 0.20 * s["prevFundCountScore"] + 0.20 * s["prevFundWeightAScore"] + 0.20 * s["prevFreeFloatScore"], 2)

for level in ("L1", "L2"):
    group = [x for x in industries if x["level"] == level]
    for x in group:
        values = industry_db[level].get(x["name"], {})
        x["fundWeightA"] = values.get("fundWeightA")
        x["overweight"] = values.get("overweight")
        x["marketCapPct"] = x["fundWeightA"] - x["overweight"] if x["fundWeightA"] is not None and x["overweight"] is not None else None
        x["prevFundWeightA"] = values.get("prevFundWeightA")
        x["prevOverweight"] = values.get("prevOverweight")
        x["prevMarketCapPct"] = x["prevFundWeightA"] - x["prevOverweight"] if x["prevFundWeightA"] is not None and x["prevOverweight"] is not None else None
        # 因子2：配置比例2010年以来历史分位（超配历史缺流通市值序列，配置分位为其代理，兴证同口径）
        hist = (values.get("allocHistDates"), values.get("allocHistVals"))
        x["histPct"] = hist_percentile(hist, x["fundWeightA"])
        x["prevHistPct"] = hist_percentile(hist, x["prevFundWeightA"], end="2026-06-30")
        if level == "L1" and x["name"] in pos_week:
            x.update(pos_week[x["name"]])
    for key in ("overweight", "fundWeightA", "prevOverweight", "prevFundWeightA"):
        minmax_scores(group, key, missing=0)
    for x in group:
        # 拥挤度 = 50%超配比例(同层级min-max) + 50%配置比例2010年以来历史分位
        x["crowdRaw"] = 0.50 * x["overweightScore"] + 0.50 * (x["histPct"] if x["histPct"] is not None else 50.0)
        x["crowdScore"] = 100 - x["crowdRaw"]
        x["prevCrowdRaw"] = round(0.50 * x["prevOverweightScore"] + 0.50 * (x["prevHistPct"] if x["prevHistPct"] is not None else 50.0), 2)
for x in stocks + industries:
    x["crowdRaw"] = round(x["crowdRaw"], 2)
    x["crowdScore"] = round(x["crowdScore"], 2)
    x["crowdChange"] = round(x["crowdRaw"] - x["prevCrowdRaw"], 2)
    x["overweightChange"] = round(x["overweight"] - x["prevOverweight"], 6) if x["overweight"] is not None and x["prevOverweight"] is not None else None
    x["fundWeightChange"] = round(x["fundWeightA"] - x["prevFundWeightA"], 6) if x["fundWeightA"] is not None and x["prevFundWeightA"] is not None else None
    x["crowdRisk"] = "极高" if x["crowdRaw"] >= 80 else ("高" if x["crowdRaw"] >= 60 else ("中" if x["crowdRaw"] >= 40 else ("低" if x["crowdRaw"] >= 20 else "极低")))


def score_group(items, is_stock=False):
    for horizon in ("month", "ytd", "since924"):
        negative = [x for x in items if x.get(horizon) is not None and x[horizon] < 0]
        for x in negative:
            x[horizon + "Loss"] = -x[horizon]
        cross_scores(negative, horizon + "Loss", higher=True, missing=0)
        for x in items:
            x[horizon + "Score"] = x.get(horizon + "LossScore", 0)
    cross_scores(items, "epsGrowth", higher=True, missing=50)
    cross_scores(items, "epsRev", higher=True, missing=50)
    cross_scores(items, "roe27", higher=True, missing=50)
    cross_scores(items, "roeDelta", higher=True, missing=50)
    cross_scores(items, "roeRev", higher=True, missing=50)
    for x in items:
        x["priceScore"] = (x["monthScore"] + x["ytdScore"] + x["since924Score"]) / 3
        x["profitScore"] = 0.20 * x["epsLevelRaw"] + 0.15 * x["epsGrowthScore"] + 0.65 * x["epsRevScore"]
        x["roeScore"] = (x["roe27Score"] + x["roeDeltaScore"] + x["roeRevScore"]) / 3
        # 连续周度上修加分：尾部连续≥2次+3分、连续3次+5分；盈利分按EPS、ROE分按ROE各自适用，上限100。
        x["profitBonus"] = 5 if x["epsStreak"] >= 3 else (3 if x["epsStreak"] == 2 else 0)
        x["roeBonus"] = 5 if x["roeStreak"] >= 3 else (3 if x["roeStreak"] == 2 else 0)
        x["profitScore"] = clamp(x["profitScore"] + x["profitBonus"])
        x["roeScore"] = clamp(x["roeScore"] + x["roeBonus"])
        x["fundamentalScore"] = (x["profitScore"] + x["roeScore"]) / 2
        cheap = 100 - x["valPct"] if x["valPct"] is not None else 50
        x["cheapScore"] = clamp(cheap)
        if is_stock:
            x["matchScore"] = None
            x["valuationScore"] = 0.5 * x["cheapScore"] + 0.5 * x["digestScore"]
        else:
            x["matchScore"] = None
            x["valuationScore"] = x["cheapScore"]
        # 回测校准权重（2026-11-17）：拥挤25%/估值35%上调（方向显著），价格/ROE各降至5%（无信号/近5年无贡献）
        x["totalScore"] = 0.05 * x["priceScore"] + 0.30 * x["profitScore"] + 0.05 * x["roeScore"] + 0.35 * x["valuationScore"] + 0.25 * x["crowdScore"]
        for k in ("priceScore", "profitScore", "roeScore", "fundamentalScore", "cheapScore", "valuationScore", "totalScore"):
            x[k] = round(x[k], 2)
        if x["matchScore"] is not None:
            x["matchScore"] = round(x["matchScore"], 2)


score_group(stocks, True)
for level in ("L1", "L2"):
    score_group([x for x in industries if x["level"] == level], False)

for level in ("L1", "L2"):
    for x in [z for z in industries if z["level"] == level]:
        children = [s for s in stocks if (s["industry"] == x["name"] if level == "L1" else s["subindustry"] == x["name"])]
        eligible_children = [s for s in children if not s["recentIPO"]]
        dividend = [s for s in eligible_children if s["divTrigger"] and (s["marketCap"] or 0) >= 100]
        x["stockCount"] = len(eligible_children)
        x["recentIPOCount"] = len(children) - len(eligible_children)
        x["highScoreCount"] = sum(1 for s in eligible_children if s["totalScore"] >= 70 and s["valPass"])
        x["unlockCount"] = sum(1 for s in eligible_children if s["unlockAlert"])
        x["dividendQualityPct"] = 100 * sum(1 for s in dividend if s["divPass"]) / len(dividend) if dividend else None
        x["divTrigger"] = x["staticYield"] is not None and x["staticYield"] > 4

for group in (stocks, [x for x in industries if x["level"] == "L1"], [x for x in industries if x["level"] == "L2"]):
    for rank, x in enumerate(sorted(group, key=lambda z: z["totalScore"], reverse=True), 1):
        x["rank"] = rank

_asof_s = ASOF.strftime("%Y-%m-%d")
payload = {"asof": _asof_s, "monthStart": _asof_s, "unlockEnd": add_two_months(ASOF).strftime("%Y-%m-%d"), "industries": industries, "stocks": stocks, "divStrategy": DIV_STRATEGY}
summary = {
    "stocks": sum(not s["recentIPO"] for s in stocks), "recentIPOs": sum(s["recentIPO"] for s in stocks), "l1": 31, "l2": 134,
    "highIndustries": sum(x["level"] == "L1" and x["totalScore"] >= 60 for x in industries),
    "highStocks": sum(s["totalScore"] >= 70 and s["valPass"] and not s["recentIPO"] for s in stocks),
    "divPass": sum(s["divTrigger"] and s["divPass"] and not s["recentIPO"] and (s["marketCap"] or 0) >= 100 for s in stocks), "unlock": sum(s["unlockAlert"] and not s["recentIPO"] for s in stocks),
    "heldStocks": sum(s["hasFundHolding"] and not s["recentIPO"] for s in stocks),
    "crowdedL1": sum(x["level"] == "L1" and x["crowdRaw"] >= 60 for x in industries),
    "crowdedStocks": sum(s["crowdRaw"] >= 60 and not s["recentIPO"] for s in stocks),
    "topIndustries": [(x["name"], round(x["totalScore"], 1)) for x in sorted([z for z in industries if z["level"] == "L1"], key=lambda z: z["totalScore"], reverse=True)[:8]],
    "topStocks": [(s["name"], s["code"], round(s["totalScore"], 1), s["digestLabel"]) for s in sorted([z for z in stocks if z["valPass"] and not z["recentIPO"]], key=lambda z: z["totalScore"], reverse=True)[:10]]
}
DATA_JSON = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
SUMMARY_JSON = json.dumps(summary, ensure_ascii=False)

HTML = r'''<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>行业与个股性价比评分 · 2026-11-13</title>
<style>
:root{--ink:#18243a;--muted:#6b7688;--line:#e4e8ef;--bg:#f5f6f8;--navy:#162a49;--red:#c83c45;--green:#14806a;--blue:#3467c9;--amber:#a66716}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}button,input,select{font:inherit}.wrap{max-width:1440px;margin:auto;padding:0 26px 42px}.hero{background:linear-gradient(120deg,#14243e,#24446d);color:#fff;padding:28px 0 0}.hero .wrap{padding-bottom:0}.eyebrow{font-size:11px;letter-spacing:2px;color:#b7c9e4}.hero h1{font-size:28px;margin:6px 0}.hero p{margin:0;color:#d5e0ee}.heroLine{display:flex;justify-content:space-between;align-items:center}.date{text-align:right;color:#c5d3e6;font-size:12px}.date b{display:block;color:#fff;font-size:19px}.nav{display:flex;gap:5px;margin-top:22px;overflow:auto}.nav button{border:0;background:transparent;color:#cfdbec;padding:11px 17px;cursor:pointer;border-radius:7px 7px 0 0;white-space:nowrap}.nav button.active{background:var(--bg);color:var(--navy);font-weight:700}.page{display:none;padding-top:18px}.page.active{display:block}.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:11px;margin-bottom:16px}.card,.panel{background:#fff;border:1px solid var(--line);border-radius:10px}.card{padding:15px}.label,.sub{font-size:12px;color:var(--muted)}.metric{font-size:27px;font-weight:750;line-height:1.2;margin:4px 0}.panel{padding:19px;margin-bottom:15px}.panel h2{font-size:16px;margin:0 0 3px}.panelHead{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:13px}.weights{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:16px}.weight{background:#fff;border:1px solid var(--line);border-radius:9px;padding:13px 15px}.weight b{font-size:14px}.weight span{float:right;color:var(--red);font-weight:750}.weight p{margin:4px 0 0;font-size:12px;color:var(--muted)}.grid2{display:grid;grid-template-columns:1.55fr 1fr;gap:15px}.toolbar{display:flex;gap:8px;flex-wrap:wrap;margin:11px 0}.toolbar input,.toolbar select,.btn{border:1px solid #d5dce6;background:#fff;border-radius:7px;padding:8px 10px;color:var(--ink)}.toolbar input{min-width:200px}.btn{cursor:pointer}.btn.primary{background:var(--navy);color:#fff;border-color:var(--navy)}.tableWrap{overflow:auto;border:1px solid var(--line);border-radius:8px}table{width:100%;border-collapse:collapse;white-space:nowrap;font-size:13px}th{background:#f6f8fb;color:#5c687a;font-weight:600;text-align:right;padding:10px;position:sticky;top:0;z-index:1}td{padding:10px;text-align:right;border-top:1px solid #edf0f4}th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left}tbody tr:hover{background:#fafbfc}.click{cursor:pointer}.click:hover{color:var(--blue);text-decoration:underline}.score{font-weight:800;font-size:15px}.score.high{color:var(--red)}.score.mid{color:var(--blue)}.miniScore{display:inline-block;width:42px;height:5px;background:#edf0f4;border-radius:4px;margin-left:6px;vertical-align:middle;overflow:hidden}.miniScore i{display:block;height:100%;background:var(--blue)}.tag{display:inline-block;border-radius:5px;padding:2px 7px;font-size:12px;background:#eff2f5;color:#5d697a;white-space:nowrap}.tag.red{background:#fff0f1;color:var(--red)}.tag.green{background:#ecf8f4;color:var(--green)}.tag.amber{background:#fff6e8;color:var(--amber)}.tag.blue{background:#eef4ff;color:var(--blue)}.up{color:var(--red)}.down{color:var(--green)}.scatter{width:100%;height:385px}.scatter text{font-size:11px;fill:#748094}.scatter .grid{stroke:#e2e7ed;stroke-dasharray:3 4}.scatter .axis{stroke:#b9c3d0}.scatter circle{cursor:pointer;stroke:#fff;stroke-width:1.5}.topList{display:grid;gap:3px}.topItem{display:grid;grid-template-columns:28px 1fr auto;align-items:center;gap:9px;padding:9px 0;border-bottom:1px solid #edf0f4}.rank{width:23px;height:23px;border-radius:6px;background:#eff2f6;text-align:center;line-height:23px;color:#657186}.topScore{font-size:18px;font-weight:800;color:var(--red)}.legend{font-size:12px;color:var(--muted)}.note{padding:11px 14px;border:1px solid #eadfc8;background:#fff9ed;color:#72531f;border-radius:8px;margin-bottom:15px;font-size:13px}.empty{text-align:center;padding:30px;color:var(--muted)}.pager{display:flex;justify-content:flex-end;gap:8px;align-items:center;margin-top:11px}.drawer{position:fixed;right:-540px;top:0;width:520px;max-width:96vw;height:100vh;background:#fff;box-shadow:-12px 0 35px #17233b20;z-index:10;padding:24px;overflow:auto;transition:right .2s}.drawer.open{right:0}.overlay{display:none;position:fixed;inset:0;background:#14243f45;z-index:9}.overlay.open{display:block}.close{float:right;border:0;background:#eef1f5;border-radius:50%;width:30px;height:30px;cursor:pointer}.detailGrid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:16px 0}.detailGrid .card{padding:10px}.detailGrid .metric{font-size:20px}.dim{margin:10px 0}.dimHead{display:flex;justify-content:space-between;font-size:12px}.track{height:7px;background:#edf0f4;border-radius:6px;overflow:hidden;margin-top:4px}.track i{display:block;height:100%;background:var(--blue)}.miniChart{width:100%;height:190px}.rules{display:grid;grid-template-columns:1fr 1fr;gap:12px}.rule{border:1px solid var(--line);border-radius:8px;padding:14px}.rule b{font-size:14px}.rule p{margin:5px 0 0;color:var(--muted);font-size:13px}.unlockGrid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.unlockItem{border:1px solid var(--line);border-radius:8px;padding:13px}.foot{font-size:12px;color:var(--muted);margin-top:15px}@media(max-width:1000px){.cards{grid-template-columns:repeat(3,1fr)}.grid2{grid-template-columns:1fr}.unlockGrid{grid-template-columns:repeat(2,1fr)}}@media(max-width:650px){.wrap{padding:0 13px 30px}.hero h1{font-size:23px}.heroLine{display:block}.date{text-align:left;margin-top:10px}.cards{grid-template-columns:repeat(2,1fr)}.weights{grid-template-columns:1fr 1fr}.rules{grid-template-columns:1fr}.unlockGrid{grid-template-columns:1fr}.panel{padding:13px}.detailGrid{grid-template-columns:repeat(2,1fr)}}
.scatter{height:auto;background:#fff;border:1px solid #edf0f4;border-radius:8px}.scatter svg{display:block}.scatter text{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;fill:#4b586b}.scatter .tick{font-size:12px;fill:#697586}.scatter .axis{stroke:#8d99aa;stroke-width:1.2}.scatter .grid{stroke:#cbd3de;stroke-width:1;stroke-dasharray:4 4}.scatter .axisTitle{font-size:12px;font-weight:600;fill:#556174}.scatter .indLabel{font-size:12px;font-weight:650;fill:#243754;paint-order:stroke;stroke:#fff;stroke-width:4px;stroke-linejoin:round}.scatter .otherLabel{font-size:11px;font-weight:500;fill:#536174}.scatter circle{cursor:pointer;transition:fill-opacity .15s}.scatter circle:hover{fill-opacity:1;stroke:#14243e;stroke-width:2}.scatter .leader{stroke:#9aa7b8;stroke-width:1}.chartFoot{display:flex;justify-content:space-between;gap:12px;margin-top:8px;font-size:12px;color:#697586}.chartKey{display:flex;gap:14px;flex-wrap:wrap}.chartKey i{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px}.panelHead .btn{white-space:nowrap;padding:6px 10px;font-size:12px}
.metricPill{display:inline-flex;align-items:center;justify-content:center;min-width:52px;height:27px;padding:0 9px;border:1px solid transparent;border-radius:6px;font-size:12px;font-weight:750;font-variant-numeric:tabular-nums}.toneGood{background:#e8f5f0;color:#14715e;border-color:#cce8dd}.toneMid{background:#edf3fd;color:#3467a9;border-color:#dce7f7}.toneWatch{background:#fff5df;color:#9a6713;border-color:#f3e3bb}.toneRisk{background:#fff0f0;color:#b84a52;border-color:#f2d4d6}.cycleTag{display:inline-flex;align-items:center;min-height:25px;padding:2px 9px;border:1px solid transparent;border-radius:6px;font-size:12px;font-weight:650}.valuationTag{display:inline-flex;align-items:center;min-height:25px;padding:2px 8px;border:1px solid transparent;border-radius:6px;font-size:12px;font-weight:650;font-variant-numeric:tabular-nums}.scoreLegend{display:flex;gap:13px;align-items:center;flex-wrap:wrap;margin:-3px 0 12px;font-size:11px;color:#707c8d}.scoreLegend span{display:inline-flex;align-items:center;gap:5px}.scoreLegend i{width:8px;height:8px;border-radius:2px;display:inline-block}
</style></head><body>
<header class="hero"><div class="wrap"><div class="heroLine"><div><div class="eyebrow">MULTI-FACTOR VALUATION SCORE</div><h1>行业与个股性价比评分</h1><p>盈利动能 × ROE质量 × 估值分位 × 价格回撤 × 公募拥挤度</p></div><div class="date">数据截面<b>2026-11-13</b>公募持仓截至2026Q3</div></div><nav class="nav"><button class="active" data-page="overview">行业评分</button><button data-page="stocks">个股评分</button><button data-page="crowding">公募拥挤度</button><button data-page="dividends">红利资产</button><button data-page="unlocks">解禁预警</button><button data-page="method">评分口径</button></nav></div></header>
<main class="wrap"><section id="overview" class="page active"><div class="cards" id="summaryCards"></div>
<div class="weights"><div class="weight"><b>盈利动能</b><span>30%</span><p>27E水平20% / 同比增速15% / 周度上修65%＋连续上修加分</p></div><div class="weight"><b>ROE质量</b><span>5%</span><p>27E水平 / 同比变化 / 周度上修＋连续上修加分</p></div><div class="weight"><b>估值性价比</b><span>35%</span><p>个股：估值分位与利润消化等权</p></div><div class="weight"><b>价格回撤</b><span>5%</span><p>近1月 / 年初以来 / 924以来等权</p></div><div class="weight"><b>公募拥挤度</b><span>25%</span><p>2026Q3 · 越拥挤安全分越低</p></div></div>
<div class="grid2"><div class="panel"><div class="panelHead"><div><h2>一级行业：估值分位与综合评分</h2><div class="sub">横轴越左越便宜 · 纵轴越上综合分越高 · 行业估值仅采用历史分位</div></div><button class="btn" id="toggleLabels">显示全部标签</button></div><div id="industryScatter" class="scatter"></div><div class="chartFoot"><div class="chartKey"><span><i style="background:#3467c9"></i>重点行业</span><span><i style="background:#a6b0bf"></i>其他行业</span></div><span>点击圆点查看行业详情</span></div></div><div class="panel"><h2>综合评分领先行业</h2><div class="sub">不再以各指标前1/3取交集</div><div id="topIndustries" class="topList" style="margin-top:12px"></div></div></div>
<div class="panel"><div class="panelHead"><div><h2>行业评分明细</h2><div class="sub">点击行业名称查看详细分解；“成分股”按钮跳转个股｜综合分 = 30%盈利 + 10%ROE + 30%估值 + 10%价格 + 20%公募拥挤度安全分</div></div></div><div class="toolbar"><select id="indLevel"><option value="L1">申万一级</option><option value="L2">申万二级</option></select><select id="indMin"><option value="0">全部评分</option><option value="60">综合分≥60</option><option value="70">综合分≥70</option><option value="80">综合分≥80</option></select><input id="indSearch" placeholder="搜索行业 / 代码"><button class="btn" id="exportInd">导出行业CSV</button></div><div class="scoreLegend"><span><i style="background:#14715e"></i>优势 / 向上</span><span><i style="background:#3467a9"></i>中性 / 平稳</span><span><i style="background:#9a6713"></i>偏弱 / 关注</span><span><i style="background:#b84a52"></i>向下 / 高风险</span><span>拥挤度与估值分位按风险反向着色</span></div><div class="tableWrap"><table><thead><tr><th>排名</th><th>行业</th><th>综合分</th><th>价格</th><th>盈利</th><th>ROE</th><th>估值分</th><th>公募拥挤</th><th>估值法/分位</th><th>ROE周期</th><th>高分股</th><th>风险</th></tr></thead><tbody id="indBody"></tbody></table></div></div></section>
<section id="stocks" class="page"><div class="panel"><div class="panelHead"><div><h2>个股综合评分</h2><div class="sub">按估值状态与利润消化分类；默认显示低估值A股，红利质量仅在独立页签判断</div></div><span class="tag blue" id="resultCount"></span></div><div class="toolbar"><input id="stockSearch" placeholder="搜索代码 / 简称 / 行业"><select id="stockIndustry"><option value="">全部一级行业</option></select><select id="stockMode"><option value="low">低估值（分位＜50%）</option><option value="all">全部A股</option><option value="mid">估值中性（50%—80%）</option><option value="high">高估值（分位≥80%）</option><option value="digest">远期利润透支（28E及以后）</option></select><select id="marketCapFilter"><option value="">全部市值</option><option value="small">0—100亿</option><option value="mid">100—500亿</option><option value="large">500亿以上</option></select><select id="listingFilter"><option value="">上市满一年（默认）</option><option value="all">包含上市未满一年</option><option value="recent">仅上市未满一年</option></select><select id="stockMin"><option value="0">全部评分</option><option value="60">综合分≥60</option><option value="70">综合分≥70</option><option value="80">综合分≥80</option></select><select id="cycleFilter"><option value="">全部ROE周期</option><option>底部反弹</option><option>持续向上</option><option>修复向上</option><option>高位平稳</option><option>低位平稳</option><option>拐头向下</option></select><button class="btn primary" id="exportStocks">导出CSV</button></div><div class="tableWrap"><table><thead><tr><th>排名</th><th>股票</th><th>行业</th><th>综合分</th><th>价格</th><th>盈利</th><th>ROE</th><th>估值分位</th><th>利润消化</th><th>27E PE</th><th>公募拥挤</th><th>解禁</th></tr></thead><tbody id="stockBody"></tbody></table></div><div class="pager"><button class="btn" id="prevPage">上一页</button><span class="sub" id="pageLabel"></span><button class="btn" id="nextPage">下一页</button></div></div></section>
<section id="crowding" class="page"><div class="cards" style="grid-template-columns:repeat(4,1fr)" id="crowdCards"></div><div class="panel"><div class="panelHead"><div><h2>行业公募持仓拥挤度</h2><div class="sub">行业拥挤度＝50%超配比例＋50%配置比例2010年以来历史分位｜超配在同层级内线性缩放0—100；历史分位取自身2010Q1以来季度序列（配置口径，数据库同源）</div></div></div><div class="toolbar"><select id="crowdIndLevel"><option value="L1">申万一级</option><option value="L2">申万二级</option></select><input id="crowdIndSearch" placeholder="搜索行业"></div><div class="tableWrap" style="max-height:520px"><table><thead><tr><th>行业</th><th>公募配置占比</th><th>流通市值占比</th><th>超配(pct)</th><th>超配环比<br>26Q3-26Q2</th><th>超配得分<br>权重50%</th><th>历史分位<br>2010以来·权重50%</th><th>当前仓位*<br>周报</th><th>930以来主动*<br>pct</th><th>拥挤度</th><th>风险</th><th>安全得分</th></tr></thead><tbody id="crowdIndBody"></tbody></table></div></div><div class="panel"><div class="panelHead"><div><h2>个股公募持仓拥挤度</h2><div class="sub">拥挤度＝40%超配比例分位＋20%持有基金数分位＋20%公募配置占比分位＋20%占流通股比例分位；数值越高越拥挤</div></div><span class="tag blue" id="crowdStockCount"></span></div><div class="toolbar"><input id="crowdStockSearch" placeholder="搜索代码 / 简称 / 行业"><select id="crowdRiskFilter"><option value="">全部风险等级</option><option>极高</option><option>高</option><option>中</option><option>低</option><option>极低</option></select></div><div class="tableWrap" style="max-height:620px"><table><thead><tr><th>股票</th><th>行业</th><th>持有基金数</th><th>公募配置占比</th><th>流通市值占比</th><th>超配(pct)</th><th>超配环比<br>26Q3-26Q2</th><th>占流通股<br>权重20%</th><th>拥挤度</th><th>风险</th></tr></thead><tbody id="crowdStockBody"></tbody></table></div></div><div class="note">口径：2026Q3主动偏股基金A股前十大重仓，截至2026-09-30；2026Q2仅用于同口径环比，不与26H1全持仓直接比较。个股配置占比按Q3重仓股合计归一为100%，流通市值由持股总市值÷持股占流通股比反推；行业配置与超配直接取数据库的一级、二级行业汇总。超配＝公募A股配置占比－流通市值占比，正值为超配、负值为低配。个股四项指标分别在全部A股中转为0—100分位，按超配40%/持有基金数20%/公募配置占比20%/占流通股比例20%加权；行业按超配比例50%和配置比例2010年以来历史分位50%加权，超配在同层级内线性缩放为0—100分（最高100、最低0），一级、二级分别计算；历史分位为该行业自身2010Q1以来配置比例序列的分位（配置口径代理超配）。未被重仓持有的股票，四项分位均计0。两份来源的持仓总额与行业汇总略有差异，因此不跨来源混算；季报披露有滞后，不代表当前实时持仓。</div></section>
<section id="dividends" class="page"><div class="panel"><h2>红利资产筛选（估值走廊框架）</h2><div class="sub">初筛＝ROE27E≥7%＋ROE变异系数≤全A 30%分位＋静态/27E股息率均≥2%＋分红稳定＋OCF/DPS≥150%＋市值≥100亿；定价＝三阶段DDM理论PB（贴现率7%），走廊[0.8,1.2]×理论PB；应有股息率=3.1%+0.3×(ROE标准差-1.1pp)。详见《红利资产筛选策略说明》</div><div class="cards" style="grid-template-columns:repeat(5,1fr);max-width:1100px" id="divCards"></div><div class="toolbar"><select id="divFilter"><option value="all">全部价值资产池</option><option value="low">低于走廊下沿（低估）</option><option value="in">走廊内</option><option value="high">高于走廊上沿</option><option value="rh">红马增强标签</option></select><button class="btn" id="exportDiv">导出CSV</button></div><div class="tableWrap"><table><thead><tr><th>排名</th><th>股票</th><th>行业</th><th>红利综合分</th><th>估值折价</th><th>股息边际</th><th>27PB/理论PB</th><th>走廊</th><th>理论PB</th><th>27E股息率</th><th>应有股息率</th><th>超额(pct)</th><th>ROE27E</th><th>ROE波动CV</th><th>分红率</th><th>OCF/DPS</th><th>标签</th></tr></thead><tbody id="divBody"></tbody></table></div></div></section>
<section id="unlocks" class="page"><div class="panel"><h2>解禁预警</h2><div class="sub">按解禁日期由近至远｜2026-11-13至2027-01-02，解禁比例＞10%｜金额按2026-11-13股价估算，仅预警、不影响评分</div><div class="unlockGrid" id="unlockList" style="margin-top:14px"></div></div></section>
<section id="method" class="page"><div class="note">“下半年以来”源表公式实际从2026/1/1起算，因此页面统一称为“年初以来”。所有分均转为0—100，缺失的盈利/ROE子项按中性50处理。</div><div class="rules"><div class="rule"><b>价格回撤分 · 5%</b><p>近1月、年初以来、924以来三个区间等权。回测显示行业截面动量/反转无显著前瞻信号（297周IC≈0），故权重由10%降至5%，仅作辅助。仅负收益参与跌幅横截面分位，跌得越深分越高；非负收益该区间计0分，避免把上涨较少当成超跌。</p></div><div class="rule"><b>盈利动能分 · 30%</b><p>内部权重：27E EPS在2021—2027E自身序列的位置20%、27E EPS同比增速的横截面分位15%、27E EPS较上周调整幅度的横截面分位65%。回测显示"当期增速最高"是次年反指（2021-2025五年IC全负），增速子项由30%降至15%，边际上修（变化而非水平）升为65%。源表为周度EPS序列（最新/上周/上上周/上上上周），周度上修取27E最新值较上周的变化。连续上修加分：尾部连续≥2次周度上修盈利分+3分、连续3次+5分（上限100分），奖励上修趋势的持续性。</p></div><div class="rule"><b>ROE质量分 · 5%</b><p>三等分：27E ROE绝对水平、27E－26E ROE变化、27E ROE较上周修正。ROE同样适用连续上修加分：尾部连续≥2次周度上修+3分、连续3次+5分（上限100分）。ROE周期标签仍结合2021—2027E路径判断。回测显示ROE水平与变化近5年无稳定正向贡献，权重由10%降至5%。</p></div><div class="rule"><b>行业估值分 · 35%</b><p>行业估值分＝100－PE/PB历史分位，不再使用估值匹配度。历史分位越低，估值得分越高；仅银行、非银金融、房地产用PB，其余所有行业及个股均用PE。行业不计算个股利润消化指标。</p></div><div class="rule"><b>个股估值与利润消化 · 35%</b><p>估值分位＜50%仍是默认硬筛选。个股估值分=50%便宜程度（100－估值历史分位）+50%利润消化分，两者等权；26E/27E/28E首次回到30x以下分别计100/70/30分，28E仍未消化计0分。到28E才消化给予较大惩罚，28E仍未消化给予更大惩罚。源表仅有年度EPS，消化时间只标到年份，不推算月份。</p></div><div class="rule"><b>公募持仓拥挤度 · 25%</b><p>报告期2026Q3，Q2仅用于同口径环比，不与26H1全持仓直接比较。个股拥挤度=40%超配比例的全A横截面分位+20%持有基金数的全A横截面分位+20%公募A股配置占比的全A横截面分位+20%占流通股比例的全A横截面分位。超配是核心，持有广度、绝对配置和流通盘持仓占比辅助确认。公募超配＝A股配置占比（A股合计归一100%）－同期流通市值占比。行业拥挤度=50%超配比例+50%配置比例2010年以来历史分位。超配（=公募A股配置占比－流通市值占比）在同层级行业内做0—100线性缩放，保留原始差距；历史分位取自身配置比例2010Q1以来的季度序列（数据库同源口径；严格超配历史因缺流通市值序列，以配置分位为代理）。安全得分=100－拥挤度，越拥挤得分越低。带*列为仓位周报动态（来源：公募基金行业仓位周报2026-11-13，占已披露A股前十大重仓股市值比例，主动变化为模型估算、仅一级行业），仅展示不参与打分。未持有个股的四项分位均计0。</p></div><div class="rule"><b>上市历史要求</b><p>默认剔除截至2026-11-13上市未满一年的股票，即上市日期晚于2025-11-13的标的，避免短历史PE/PB分位被误判为长期低估。上市日期来自同花顺iFinD，清单查询于2026-10-30、经2026-11-17复核新增钶锐锶（2026-11-10上市）、经2026-11-23复核新增皇冠新材（2026-11-03）、森峰激光（2026-11-06）、菊乐股份（2026-11-17）；个股页可通过“包含上市未满一年”或“仅上市未满一年”单独查看，不参与默认高分股和行业高分股统计。</p></div><div class="rule"><b>红利资产筛选 · 估值走廊框架</b><p>采用券商策略估值走廊研究框架，辅以"红马"增强标签。初筛门槛：27E ROE≥7%、ROE变异系数（2021—2027E）≤全A个股30%分位、静态股息率与27E股息率均≥2%、分红稳定、27E OCF/DPS≥150%，另要求总市值≥100亿元。定价采用三阶段DDM：一阶段5年（ROE=27E、分红率=近三年均值，g=ROE×留存率、上限6.5%）、二阶段10年（ROE每年衰减2pct、分红率+10pct、上限0.9、g上限4%）、三阶段永续（分红率钳制0.7—0.95、g=2%），贴现率7%；g上限参照框架案例库99只资产（g1最大6.4%），防止高留存成长股被无限外推。理论PB=理论PE×ROE，估值走廊=[0.8,1.2]×理论PB，低于下沿为低估、高于上沿为溢价。盈利波动补偿：应有股息率=3.1%+0.3×(ROE标准差-1.1pp)，水电为基准。综合分=估值折价35%+股息安全边际25%+盈利质量稳定20%+分红可持续20%（均为池内横截面分位）。红马增强标签：26/27/28E EPS同比均＞-5%（成长护航）且ROE变异系数≤全A中位（盈利稳态）。不参与普通个股分类或高分股统计。解禁仅作独立预警，不进入综合评分。</p></div></div></section><div class="foot">数据源：多维数据比较1113.xlsx · 多维比较（2026-11-13）；重仓持股26Q3.xlsx、重仓持股26Q2.xlsx、主动偏股基金数据库_2026Q3.xlsx（2026Q3/Q2前十大重仓）；同花顺iFinD上市日期（2026-10-30查询、2026-11-17复核）｜评分为相对比较工具，不构成投资建议</div></main>
<div class="overlay" id="overlay"></div><aside class="drawer" id="drawer"></aside><script id="modelData" type="application/json">__DATA__</script><script>
const DATA=JSON.parse(document.getElementById('modelData').textContent),SUMMARY=__SUMMARY__,$=id=>document.getElementById(id);
const fmt=(v,d=1)=>v===null||v===undefined||Number.isNaN(v)?'—':Number(v).toFixed(d),esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const tag=(s,t='')=>`<span class="tag ${t}">${esc(s)}</span>`,score=v=>`<span class="score ${v>=70?'high':v>=50?'mid':''}">${fmt(v)}</span>`,bar=v=>`<span class="miniScore"><i style="width:${v||0}%"></i></span>`;
const metricPill=(v,risk=false)=>{const n=Number(v)||0,tone=risk?(n>=80?'toneRisk':n>=60?'toneWatch':n>=40?'toneMid':'toneGood'):(n>=70?'toneGood':n>=40?'toneMid':'toneWatch');return `<span class="metricPill ${tone}" title="${risk?'拥挤度越高，风险越大':'分数越高，表现越优'}">${fmt(v)}</span>`};
const cyclePill=c=>{const tone=c==='拐头向下'?'toneRisk':c==='低位平稳'?'toneWatch':c==='高位平稳'?'toneMid':'toneGood';return `<span class="cycleTag ${tone}">${esc(c)}</span>`};
const valuationPill=(method,pct)=>{const n=Number(pct),tone=pct===null?'toneMid':n>=80?'toneRisk':n>=60?'toneWatch':n>=30?'toneMid':'toneGood';return `<span class="valuationTag ${tone}">${method} ${fmt(pct)}%</span>`};
const signed=(v,d=2)=>v===null||v===undefined||Number.isNaN(v)?'—':(v>0?'+':'')+fmt(v,d);
let stockPage=1,pageSize=60,currentStocks=[],showAllLabels=false;
function init(){
 $('summaryCards').innerHTML=[['A股样本',SUMMARY.stocks,'剔除上市未满一年 '+SUMMARY.recentIPOs+' 只'],['一级行业',SUMMARY.l1,'二级 '+SUMMARY.l2],['及格行业',SUMMARY.highIndustries,'一级综合分≥60'],['低估值高分股',SUMMARY.highStocks,'综合分≥70且估值分位＜50%'],['红利价值资产池',DATA.divStrategy.poolCount,'估值走廊框架·低估'+DATA.divStrategy.belowFloor+'只'],['解禁预警',SUMMARY.unlock,'独立风险标注']].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="metric">${x[1]}</div><div class="sub">${x[2]}</div></div>`).join('');
 $('stockIndustry').innerHTML+=[...new Set(DATA.stocks.map(x=>x.industry))].sort().map(x=>`<option>${esc(x)}</option>`).join('');
 const l1=DATA.industries.filter(x=>x.level==='L1');renderScatter(l1);renderTop(l1);renderIndustries();renderStocks();renderCrowding();renderRisks();
 document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>{document.querySelectorAll('.nav button,.page').forEach(x=>x.classList.remove('active'));b.classList.add('active');$(b.dataset.page).classList.add('active')});
 ['indLevel','indMin','indSearch'].forEach(id=>$(id).oninput=renderIndustries);['stockSearch','stockIndustry','stockMode','marketCapFilter','listingFilter','stockMin','cycleFilter'].forEach(id=>$(id).oninput=()=>{stockPage=1;renderStocks()});
 $('prevPage').onclick=()=>{if(stockPage>1){stockPage--;renderStocks()}};$('nextPage').onclick=()=>{if(stockPage*pageSize<currentStocks.length){stockPage++;renderStocks()}};$('exportStocks').onclick=exportStocks;$('exportInd').onclick=exportIndustries;$('divFilter').oninput=renderRisks;$('overlay').onclick=closeDrawer;
 $('toggleLabels').onclick=()=>{showAllLabels=!showAllLabels;$('toggleLabels').textContent=showAllLabels?'仅显示重点标签':'显示全部标签';renderScatter(DATA.industries.filter(x=>x.level==='L1'))};
 ['crowdIndLevel','crowdIndSearch','crowdStockSearch','crowdRiskFilter'].forEach(id=>$(id).oninput=renderCrowding);
}
function renderTop(items){$('topIndustries').innerHTML=[...items].sort((a,b)=>b.totalScore-a.totalScore).slice(0,9).map((x,i)=>`<div class="topItem"><span class="rank">${i+1}</span><div><b class="click" onclick="openIndustry('${esc(x.name)}')">${esc(x.name)}</b><div class="sub">盈利 ${fmt(x.profitScore)} · 估值 ${fmt(x.valuationScore)} · 拥挤 ${fmt(x.crowdRaw)}</div></div><span class="topScore">${fmt(x.totalScore)}</span></div>`).join('')}
function renderScatter(items){
 const W=820,H=455,L=58,R=24,T=24,B=52,pw=W-L-R,ph=H-T-B,x=v=>L+v/100*pw,y=v=>T+(100-v)/100*ph;
 const visible=items.filter(d=>d.valPct!==null),topNames=new Set([...items].sort((a,b)=>b.totalScore-a.totalScore).slice(0,8).map(d=>d.name));
 let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="一级行业估值分位与综合评分散点图"><rect x="${L}" y="${T}" width="${x(50)-L}" height="${y(50)-T}" fill="#fff7f6"/><rect x="${x(50)}" y="${T}" width="${W-R-x(50)}" height="${y(50)-T}" fill="#f7f9fc"/><line class="grid" x1="${x(50)}" y1="${T}" x2="${x(50)}" y2="${H-B}"/><line class="grid" x1="${L}" y1="${y(50)}" x2="${W-R}" y2="${y(50)}"/><line class="axis" x1="${L}" y1="${H-B}" x2="${W-R}" y2="${H-B}"/><line class="axis" x1="${L}" y1="${T}" x2="${L}" y2="${H-B}"/>`;
 [0,25,50,75,100].forEach(v=>{s+=`<line x1="${x(v)}" y1="${H-B}" x2="${x(v)}" y2="${H-B+5}" stroke="#8d99aa"/><text class="tick" x="${x(v)}" y="${H-B+23}" text-anchor="middle">${v}</text><line x1="${L-5}" y1="${y(v)}" x2="${L}" y2="${y(v)}" stroke="#8d99aa"/><text class="tick" x="${L-10}" y="${y(v)+4}" text-anchor="end">${v}</text>`});
 visible.forEach(d=>{const cx=x(d.valPct),cy=y(d.totalScore),r=5+d.totalScore/16,top=topNames.has(d.name);s+=`<circle cx="${cx}" cy="${cy}" r="${r}" fill="${top?'#3467c9':'#a6b0bf'}" fill-opacity="${top?'.9':'.72'}" onclick="openIndustry('${esc(d.name)}')"><title>${esc(d.name)}｜综合 ${fmt(d.totalScore)}｜估值分位 ${fmt(d.valPct)}%｜估值得分 ${fmt(d.valuationScore)}</title></circle>`});
 const placed=[],labels=visible.filter(d=>showAllLabels||topNames.has(d.name)).sort((a,b)=>b.totalScore-a.totalScore);
 labels.forEach(d=>{const cx=x(d.valPct),cy=y(d.totalScore),r=5+d.totalScore/16,w=d.name.length*(showAllLabels?11:12)+8,h=19,cands=[[cx+r+8,cy-r-h],[cx+r+8,cy+r+2],[cx-w-r-8,cy-r-h],[cx-w-r-8,cy+r+2],[cx-w/2,cy-r-h-9],[cx-w/2,cy+r+9],[cx+r+14,cy-h/2],[cx-w-r-14,cy-h/2]];[22,40,58,76].forEach(dy=>{cands.push([cx+r+8,cy-r-h-dy],[cx+r+8,cy+r+dy],[cx-w-r-8,cy-r-h-dy],[cx-w-r-8,cy+r+dy],[cx-w/2,cy-r-h-dy],[cx-w/2,cy+r+dy])});let best=null;
  cands.forEach((c,i)=>{const bx=Math.max(L+3,Math.min(W-R-w-3,c[0])),by=Math.max(T+2,Math.min(H-B-h-2,c[1])),rect={x:bx,y:by,w:w,h:h};let overlap=0;placed.forEach(p=>{const ox=Math.max(0,Math.min(bx+w+6,p.x+p.w+6)-Math.max(bx-6,p.x-6)),oy=Math.max(0,Math.min(by+h+4,p.y+p.h+4)-Math.max(by-4,p.y-4));overlap+=ox*oy});const score=overlap*100+i*.08+Math.hypot(bx+w/2-cx,by+h/2-cy)*.015;if(!best||score<best.score)best={bx,by,w,h,score}});
  placed.push(best);const tx=best.bx+4,ty=best.by+14,near=Math.abs(best.bx+4-cx)<r+12&&Math.abs(ty-cy)<r+12;if(!near)s+=`<line class="leader" x1="${cx}" y1="${cy}" x2="${Math.max(best.bx,Math.min(best.bx+best.w,cx))}" y2="${best.by+best.h/2}"/>`;s+=`<text class="${topNames.has(d.name)?'indLabel':'otherLabel'}" x="${tx}" y="${ty}" onclick="openIndustry('${esc(d.name)}')" style="cursor:pointer">${esc(d.name)}</text>`;
 });
 s+=`<text class="axisTitle" x="${W-R}" y="${H-8}" text-anchor="end">PE/PB历史分位（%）→</text><text class="axisTitle" x="16" y="${T}" transform="rotate(-90 16 ${T})" text-anchor="end">综合评分 →</text></svg>`;$('industryScatter').innerHTML=s;
}
function filteredIndustries(){let q=$('indSearch').value.trim().toLowerCase(),m=Number($('indMin').value);return DATA.industries.filter(x=>x.level===$('indLevel').value&&x.totalScore>=m&&(!q||(x.name+x.code).toLowerCase().includes(q))).sort((a,b)=>b.totalScore-a.totalScore)}
function renderIndustries(){let rows=filteredIndustries();$('indBody').innerHTML=rows.map(x=>`<tr><td>${x.rank}</td><td><b class="click" onclick="openIndustry('${esc(x.name)}')">${esc(x.name)}</b><div><button class="btn" style="padding:2px 7px;font-size:11px;margin-top:4px" onclick="openIndustryStocks('${esc(x.name)}')">成分股</button></div></td><td>${score(x.totalScore)}</td><td>${metricPill(x.priceScore)}</td><td>${metricPill(x.profitScore)}</td><td>${metricPill(x.roeScore)}</td><td>${metricPill(x.valuationScore)}</td><td>${metricPill(x.crowdRaw,true)}</td><td>${valuationPill(x.method,x.valPct)}</td><td>${cyclePill(x.cycle)}</td><td>${x.highScoreCount}/${x.stockCount}</td><td>${x.unlockCount?tag(x.unlockCount+'解禁','amber'):'—'}</td></tr>`).join('')||'<tr><td colspan="12" class="empty">无匹配行业</td></tr>'}
function filteredStocks(){let q=$('stockSearch').value.trim().toLowerCase(),ind=$('stockIndustry').value,cy=$('cycleFilter').value,mode=$('stockMode').value,cap=$('marketCapFilter').value,listing=$('listingFilter').value,m=Number($('stockMin').value);return DATA.stocks.filter(x=>{const validVal=x.valPct!==null&&x.valMultiple!==null&&x.valMultiple>0,modePass=mode==='all'||mode==='low'&&x.valPass||mode==='mid'&&validVal&&x.valPct>=50&&x.valPct<80||mode==='high'&&validVal&&x.valPct>=80||mode==='digest'&&(x.digestYear===2028||x.digestLabel==='28E仍未消化'),capPass=!cap||x.marketCap!==null&&(cap==='small'&&x.marketCap>=0&&x.marketCap<100||cap==='mid'&&x.marketCap>=100&&x.marketCap<500||cap==='large'&&x.marketCap>=500),listingPass=listing==='all'||listing==='recent'&&x.recentIPO||!listing&&!x.recentIPO;return x.totalScore>=m&&modePass&&capPass&&listingPass&&(!q||(x.code+x.name+x.industry+x.subindustry).toLowerCase().includes(q))&&(!ind||x.industry===ind)&&(!cy||x.cycle===cy)}).sort((a,b)=>b.totalScore-a.totalScore)}
function renderStocks(){currentStocks=filteredStocks();let pages=Math.max(1,Math.ceil(currentStocks.length/pageSize));if(stockPage>pages)stockPage=pages;let rows=currentStocks.slice((stockPage-1)*pageSize,stockPage*pageSize);$('resultCount').textContent=currentStocks.length+' 只';$('pageLabel').textContent=`${stockPage} / ${pages} 页`;$('stockBody').innerHTML=rows.map((x,i)=>`<tr><td>${(stockPage-1)*pageSize+i+1}</td><td><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b><div class="sub">${x.code}${x.recentIPO?' · 上市未满一年':''}</div></td><td>${esc(x.industry)}</td><td>${score(x.totalScore)}</td><td>${fmt(x.priceScore)}</td><td>${fmt(x.profitScore)}</td><td>${fmt(x.roeScore)}</td><td>${x.method} ${fmt(x.valPct)}%</td><td>${tag(x.digestLabel,x.digestRisk==='低'?'green':x.digestRisk==='中'?'blue':'amber')}</td><td>${fmt(x.pe27)}x</td><td>${tag(fmt(x.crowdRaw)+' · '+x.crowdRisk,x.crowdRaw>=60?'amber':'blue')}</td><td>${x.unlockAlert?tag(fmt(x.unlockRatio)+'%','amber'):'—'}</td></tr>`).join('')||'<tr><td colspan="12" class="empty">当前条件下无匹配股票</td></tr>'}
function renderCrowding(){
 $('crowdCards').innerHTML=[['报告期','2026Q3','截至2026-09-30 · 环比26Q2'],['有公募持仓',SUMMARY.heldStocks,'模型A股样本'],['高/极高拥挤行业',SUMMARY.crowdedL1,'申万一级'],['高/极高拥挤个股',SUMMARY.crowdedStocks,'拥挤度≥60']].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="metric">${x[1]}</div><div class="sub">${x[2]}</div></div>`).join('');
 const iq=$('crowdIndSearch').value.trim(),inds=DATA.industries.filter(x=>x.level===$('crowdIndLevel').value&&(!iq||x.name.includes(iq))).sort((a,b)=>b.crowdRaw-a.crowdRaw);
 $('crowdIndBody').innerHTML=inds.map(x=>`<tr><td><b>${esc(x.name)}</b><div class="sub">${esc(x.code)}</div></td><td>${fmt(x.fundWeightA,2)}%</td><td>${fmt(x.marketCapPct,2)}%</td><td class="${x.overweight>0?'up':'down'}"><b>${signed(x.overweight)}</b></td><td class="${x.overweightChange>0?'up':'down'}">${signed(x.overweightChange)}</td><td>${fmt(x.overweightScore)}</td><td>${fmt(x.histPct)}</td><td>${x.posCurrent!=null?fmt(x.posCurrent,2)+'%':'<span class="sub">—</span>'}</td><td class="${(x.posActive||0)<0?'down':'up'}">${x.posActive!=null?signed(x.posActive):'<span class="sub">—</span>'}</td><td><b>${fmt(x.crowdRaw)}</b>${bar(x.crowdRaw)}</td><td>${tag(x.crowdRisk,x.crowdRaw>=60?'amber':'blue')}</td><td>${fmt(x.crowdScore)}</td></tr>`).join('')||'<tr><td colspan="10" class="empty">无匹配行业</td></tr>';
 const q=$('crowdStockSearch').value.trim().toLowerCase(),risk=$('crowdRiskFilter').value,rows=DATA.stocks.filter(x=>!x.recentIPO&&(!risk||x.crowdRisk===risk)&&(!q||(x.code+x.name+x.industry+x.subindustry).toLowerCase().includes(q))).sort((a,b)=>b.crowdRaw-a.crowdRaw);
 const dist=['极高','高','中','低','极低'].map(k=>k+(rows.filter(x=>x.crowdRisk===k).length)).join(' / ');$('crowdStockCount').textContent=rows.length+' 只｜'+dist;$('crowdStockBody').innerHTML=rows.map(x=>`<tr><td><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b><div class="sub">${x.code}</div></td><td>${esc(x.industry)}</td><td>${fmt(x.fundCount,0)}</td><td>${fmt(x.fundWeightA,3)}%</td><td>${fmt(x.marketCapPct,3)}%</td><td class="${x.overweight>0?'up':'down'}"><b>${signed(x.overweight,3)}</b></td><td class="${x.overweightChange>0?'up':'down'}">${x.hasFundHolding&&!x.hasPrevFundHolding?'新进':signed(x.overweightChange,3)}</td><td>${fmt(x.freeFloat,2)}%</td><td><b>${fmt(x.crowdRaw)}</b></td><td>${tag(x.crowdRisk,x.crowdRaw>=60?'amber':'blue')}</td></tr>`).join('')||'<tr><td colspan="10" class="empty">无匹配股票</td></tr>';
}
function renderRisks(){const DS=DATA.divStrategy,list=DS.list||[];$('divCards').innerHTML=[['价值资产池',DS.poolCount,'六道门槛初筛'],['低于走廊下沿',DS.belowFloor,'PB＜0.8×理论PB'],['走廊内',DS.inCorridor,'0.8—1.2×理论PB'],['高于走廊上沿',DS.aboveCap,'PB＞1.2×理论PB'],['红马增强',DS.redHorseN,'成长护航＋盈利稳态']].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="metric">${x[1]}</div><div class="sub">${x[2]}</div></div>`).join('');const f=$('divFilter').value,rows=list.filter(x=>f==='all'||f==='low'&&x.corridor==='低于下沿'||f==='in'&&x.corridor==='走廊内'||f==='high'&&x.corridor==='高于上沿'||f==='rh'&&x.redHorse);const corrTag=c=>c==='低于下沿'?tag(c,'green'):c==='高于上沿'?tag(c,'red'):tag(c,'blue');$('divBody').innerHTML=rows.map(x=>`<tr><td>${x.rank}</td><td><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b></td><td>${esc(x.industry)}</td><td><b class="${x.score>=70?'up':''}">${fmt(x.score,1)}</b></td><td>${fmt(x.valScore,0)}</td><td>${fmt(x.yieldScore,0)}</td><td>${fmt(x.ratio,2)}</td><td>${corrTag(x.corridor)}</td><td>${fmt(x.theoryPB,2)}</td><td class="up"><b>${fmt(x.yield27,2)}%</b></td><td>${fmt(x.reqYield,2)}%</td><td class="${x.excessYield>=0?'up':'down'}">${signed(x.excessYield)}</td><td>${fmt(x.roe27,1)}%</td><td>${fmt(x.roeCV,2)}</td><td>${fmt(x.payout,0)}%</td><td>${x.coverage!=null?fmt(x.coverage,0)+'%':'—'}</td><td>${x.redHorse?tag('红马','green'):x.growthTag?tag('成长护航','amber'):''}</td></tr>`).join('')||'<tr><td colspan="17" class="empty">无匹配股票</td></tr>';$('exportDiv').onclick=()=>csvDownload('红利资产池_20261120.csv',[['排名','代码','简称','行业','红利综合分','估值折价分','股息边际分','盈利质量分','分红可持续分','27PB/理论PB','走廊位置','理论PE','理论PB','27PB','27E股息率%','应有股息率%','超额pct','ROE27E%','ROE波动CV','ROE标准差','平均分红率%','OCF/DPS%','红马标签'],...rows.map(x=>[x.rank,x.code,x.name,x.industry,fmt(x.score,1),fmt(x.valScore,0),fmt(x.yieldScore,0),fmt(x.qualityScore,0),fmt(x.sustainScore,0),fmt(x.ratio,2),x.corridor,fmt(x.theoryPE,1),fmt(x.theoryPB,2),fmt(x.pb27,2),fmt(x.yield27,2),fmt(x.reqYield,2),signed(x.excessYield),fmt(x.roe27,1),fmt(x.roeCV,2),fmt(x.roeSD,2),fmt(x.payout,0),x.coverage!=null?fmt(x.coverage,0):'',x.redHorse?'是':'否'])]);$('unlockList').innerHTML=DATA.stocks.filter(x=>x.unlockAlert&&!x.recentIPO).sort((a,b)=>a.unlockDate.localeCompare(b.unlockDate)||b.unlockRatio-a.unlockRatio).map(x=>`<div class="unlockItem"><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b><span style="float:right">${tag(fmt(x.unlockRatio,2)+'%','amber')}</span><div class="sub">${x.code} · ${esc(x.industry)}</div><div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px;margin-top:8px"><b>${x.unlockDate}</b><span class="sub">解禁金额 <b style="color:#a66716">${fmt(x.unlockAmount,1)}亿元</b></span></div></div>`).join('')}
function openIndustryStocks(name){closeDrawer();document.querySelector('[data-page="stocks"]').click();let isL1=DATA.stocks.some(s=>s.industry===name);$('stockIndustry').value=isL1?name:'';$('stockSearch').value=isL1?'':name;$('stockMode').value='all';$('listingFilter').value='';stockPage=1;renderStocks()}
function openIndustry(name){
 const x=DATA.industries.find(d=>d.name===name);if(!x)return;
 const level=x.level==='L1'?'申万一级':'申万二级',parent=x.level==='L2'&&x.parent?` · 所属 ${esc(x.parent)}`:'';
 $('drawer').innerHTML=`<button class="close" onclick="closeDrawer()">×</button><div class="eyebrow" style="color:#748094">INDUSTRY SCORE</div><h2>${esc(x.name)} <span class="sub">${esc(x.code)}</span></h2><div>${tag(level,'blue')} ${tag(x.cycle,'blue')} ${tag('拥挤'+x.crowdRisk,x.crowdRaw>=60?'amber':'blue')}</div><p class="sub">行业综合评分第 ${x.rank} 名${parent}</p><button class="btn primary" onclick="openIndustryStocks('${esc(x.name)}')">查看成分股（${x.stockCount}只）</button><div class="detailGrid" style="margin-top:14px"><div class="card"><div class="label">综合分</div><div class="metric up">${fmt(x.totalScore)}</div></div><div class="card"><div class="label">${x.method}估值分位</div><div class="metric">${fmt(x.valPct)}%</div></div><div class="card"><div class="label">公募拥挤度</div><div class="metric">${fmt(x.crowdRaw)}</div></div><div class="card"><div class="label">低估值高分股</div><div class="metric">${x.highScoreCount}/${x.stockCount}</div></div></div>${dimensions(x)}<h3>ROE周期路径（%）</h3><div class="miniChart">${roeChart(x)}</div><div class="panel" style="padding:13px"><b>ROE质量分解 · ${fmt(x.roeScore)}分</b><p class="sub">26E / 27E ROE：${fmt(x.roe26)}% / ${fmt(x.roe27)}%；27E－26E：${signed(x.roeDelta)}pct；27E较上周：${signed(x.roeRev)}pct；连续周度上修 ${x.roeStreak||0} 次。</p><p class="sub">27E水平分位 ${fmt(x.roe27Score)}＋同比变化分位 ${fmt(x.roeDeltaScore)}＋周度修正分位 ${fmt(x.roeRevScore)}，三项等权；连续上修加分 +${fmt(x.roeBonus,0)}。</p></div><div class="panel" style="padding:13px"><b>盈利动能分解 · ${fmt(x.profitScore)}分</b><p class="sub">27E EPS：${fmt(x.eps27,3)}；27E同比增速：${signed(x.epsGrowth)}%；27E EPS较上周：${signed(x.epsRev)}%；连续周度上修 ${x.epsStreak||0} 次。</p><p class="sub">绝对景气 ${fmt(x.epsLevelRaw)} × 20%＋盈利增速分位 ${fmt(x.epsGrowthScore)} × 15%＋周度上修分位 ${fmt(x.epsRevScore)} × 65%；连续上修加分 +${fmt(x.profitBonus,0)}（连续≥2次+3分、3次+5分）。</p></div><div class="panel" style="padding:13px"><b>估值详情 · ${fmt(x.valuationScore)}分</b><p class="sub">估值方法：${x.method}；当前${x.method}：${fmt(x.valMultiple)}x；历史分位：${fmt(x.valPct)}%。</p><p class="sub">行业估值得分＝100－历史分位，不计算个股利润消化指标。仅银行、非银金融、房地产采用PB，其余采用PE。</p></div><div class="panel" style="padding:13px"><b>价格回撤分解 · ${fmt(x.priceScore)}分</b><p class="sub">近1月：${signed(x.month)}%，得分 ${fmt(x.monthScore)}；年初以来：${signed(x.ytd)}%，得分 ${fmt(x.ytdScore)}；924以来：${signed(x.since924)}%，得分 ${fmt(x.since924Score)}。</p><p class="sub">三个区间等权；仅实际下跌区间参与跌幅分位，非负收益区间计0分。</p></div><div class="panel" style="padding:13px"><b>2026Q3公募持仓拥挤度</b><p class="sub">公募A股配置占比 ${fmt(x.fundWeightA,2)}%；流通市值占比 ${fmt(x.marketCapPct,2)}%；超配 <b class="${x.overweight>0?'up':'down'}">${signed(x.overweight)}pct</b>。</p><p class="sub">较26Q2：配置 ${signed(x.fundWeightChange)}pct；超配 ${signed(x.overweightChange)}pct；拥挤度 ${signed(x.crowdChange)}（26Q2为 ${fmt(x.prevCrowdRaw)}）。</p><p class="sub">拥挤度 ${fmt(x.crowdRaw)}＝超配得分 ${fmt(x.overweightScore)} × 50%＋历史分位 ${fmt(x.histPct)} × 50%（超配在同层级内线性缩放0—100；历史分位=自身配置比例2010Q1以来季度序列分位，配置口径代理超配）；安全得分 ${fmt(x.crowdScore)}。</p>${x.posCurrent!=null?`<p class="sub">仓位周报动态（仅展示）：当前仓位 ${fmt(x.posCurrent,2)}%（占已披露前十大重仓市值）；930以来主动变化 ${signed(x.posActive)}pct；本周主动 ${signed(x.posWeek)}pct。来源：公募基金行业仓位周报（2026-11-13）。</p>`:''}</div><div class="panel" style="padding:13px"><b>成分股与风险统计</b><p class="sub">上市满一年成分股 ${x.stockCount}只；综合分≥70且估值分位＜50%的低估值高分股 ${x.highScoreCount}只；未来两个月解禁预警 ${x.unlockCount}只；另剔除上市未满一年 ${x.recentIPOCount}只。</p></div>`;
 $('drawer').classList.add('open');$('overlay').classList.add('open')
}
function dimensions(x){return[['价格回撤',x.priceScore],['盈利动能',x.profitScore],['ROE质量',x.roeScore],['估值性价比',x.valuationScore],['公募拥挤安全分',x.crowdScore]].map(d=>`<div class="dim"><div class="dimHead"><span>${d[0]}</span><b>${fmt(d[1])}</b></div><div class="track"><i style="width:${d[1]}%"></i></div></div>`).join('')}
function roeChart(x){const years=['21','22','23','24','25','26E','27E','28E'],vals=x.roe,valid=vals.filter(v=>v!==null),min=Math.min(...valid,0),max=Math.max(...valid,1),W=450,H=175,P=27,y=v=>H-P-(v-min)/(max-min)*(H-2*P);let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="100%"><line x1="${P}" y1="${y(0)}" x2="${W-P}" y2="${y(0)}" stroke="#ccd3dd"/>`;let pts=vals.map((v,i)=>v===null?null:`${P+i*(W-2*P)/7},${y(v)}`).filter(Boolean).join(' ');s+=`<polyline points="${pts}" fill="none" stroke="#3467c9" stroke-width="2.5"/>`;vals.forEach((v,i)=>{if(v===null)return;let cx=P+i*(W-2*P)/7;s+=`<circle cx="${cx}" cy="${y(v)}" r="4" fill="${i===6?'#c83c45':'#3467c9'}"/><text x="${cx}" y="${y(v)-8}" text-anchor="middle" font-size="10" fill="#596579">${fmt(v)}</text><text x="${cx}" y="${H-6}" text-anchor="middle" font-size="10" fill="#596579">${years[i]}</text>`});return s+'</svg>'}
function openStock(code){const x=DATA.stocks.find(s=>s.code===code);if(!x)return;$('drawer').innerHTML=`<button class="close" onclick="closeDrawer()">×</button><div class="eyebrow" style="color:#748094">STOCK SCORE</div><h2>${esc(x.name)} <span class="sub">${x.code}</span></h2><div>${tag(x.cycle,'blue')} ${tag(x.digestLabel,x.digestRisk==='低'?'green':'amber')} ${tag('拥挤'+x.crowdRisk,x.crowdRaw>=60?'amber':'blue')} ${x.recentIPO?tag('上市未满一年 '+x.listingDate,'amber'):''}</div><div class="detailGrid"><div class="card"><div class="label">综合分</div><div class="metric up">${fmt(x.totalScore)}</div></div><div class="card"><div class="label">估值分位</div><div class="metric">${fmt(x.valPct)}%</div></div><div class="card"><div class="label">27E PE</div><div class="metric">${fmt(x.pe27)}x</div></div><div class="card"><div class="label">28E PE</div><div class="metric">${fmt(x.pe28)}x</div></div></div>${dimensions(x)}<div class="panel" style="padding:13px"><b>利润消化路径</b><p class="sub">26E / 27E / 28E PE：${fmt(x.pe26)}x / ${fmt(x.pe27)}x / ${fmt(x.pe28)}x</p><p class="sub">首次低于30x：${x.digestYear||'28E仍未消化'}；消化分 ${fmt(x.digestScore)}。估值分=50%估值分位+50%利润消化；源表仅有年度EPS，不推算月份。</p></div><div class="panel" style="padding:13px"><b>周度上修与连续性加分</b><p class="sub">27E EPS较上周 <b class="${x.epsRev>0?'up':'down'}">${signed(x.epsRev)}%</b>；尾部连续周度上修 ${x.epsStreak||0} 次 → 盈利分加分 +${fmt(x.profitBonus,0)}（连续≥2次+3分、3次+5分）。</p><p class="sub">27E ROE较上周 <b class="${x.roeRev>0?'up':'down'}">${signed(x.roeRev)}pct</b>；尾部连续周度上修 ${x.roeStreak||0} 次 → ROE分加分 +${fmt(x.roeBonus,0)}。</p></div><div class="panel" style="padding:13px"><b>2026Q3公募持仓拥挤度</b><p class="sub">持有基金 ${fmt(x.fundCount,0)}只；公募A股配置占比 ${fmt(x.fundWeightA,3)}%；流通市值占比 ${fmt(x.marketCapPct,3)}%</p><p class="sub">公募超配 <b class="${x.overweight>0?'up':'down'}">${signed(x.overweight,3)}pct</b>＝配置占比－流通市值占比；占流通股 ${fmt(x.freeFloat,2)}%</p><p class="sub">四项全A分位：超配 ${fmt(x.overweightScore)} × 40%＋持有基金数 ${fmt(x.fundCountScore)} × 20%＋配置占比 ${fmt(x.fundWeightAScore)} × 20%＋占流通股比例 ${fmt(x.freeFloatScore)} × 20%。</p><p class="sub">较26Q2：持有基金数 ${fmt(x.prevFundCount,0)} → ${fmt(x.fundCount,0)}只；配置 ${signed(x.fundWeightChange,3)}pct；超配 ${signed(x.overweightChange,3)}pct；拥挤度 ${signed(x.crowdChange)}（26Q2为 ${fmt(x.prevCrowdRaw)}）。</p><p class="sub">拥挤度 ${fmt(x.crowdRaw)}（${x.crowdRisk}）；安全得分 ${fmt(x.crowdScore)}＝100－拥挤度。</p></div><h3>ROE周期路径（%）</h3><div class="miniChart">${roeChart(x)}</div><div class="panel" style="padding:13px"><b>风险检查</b><p>红利：${tag(x.divStatus,x.divPass?'green':'amber')}　27E股息率 ${fmt(x.yield27,2)}%</p><p>OCF/DPS：${x.coveragePass===null?'待补':fmt(x.coverage,0)+'%'}；分红：${esc(x.stable||'待补')}</p><p>解禁：${x.unlockDate||'无日期'}；比例 ${fmt(x.unlockRatio,2)}%；金额 ${fmt(x.unlockAmount,1)}亿元</p></div>`;$('drawer').classList.add('open');$('overlay').classList.add('open')}
function closeDrawer(){$('drawer').classList.remove('open');$('overlay').classList.remove('open')}
function csvDownload(name,rows){let csv='\ufeff'+rows.map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\n'),a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download=name;a.click();URL.revokeObjectURL(a.href)}
function exportStocks(){csvDownload('个股性价比评分_20261113.csv',[['代码','简称','行业','总市值(亿元)','上市日期','上市未满一年','综合分','价格分','盈利分','盈利加分','EPS连续上修(次)','ROE分','ROE加分','ROE连续上修(次)','EPS周度上修(%)','27E EPS较上周(元)','估值分','公募拥挤度','26Q2拥挤度','拥挤度环比','拥挤风险','公募配置环比(pct)','公募超配环比(pct)','持有基金数','公募原口径持仓权重','公募A股配置占比','流通市值占比','公募超配(pct)','流通股持仓占比','估值法','估值分位','利润消化','消化分','26PE','27PE','28PE','红利状态','解禁日期','解禁比例','解禁金额(亿元)'],...currentStocks.map(x=>[x.code,x.name,x.industry,fmt(x.marketCap),x.listingDate,x.recentIPO?'是':'否',fmt(x.totalScore),fmt(x.priceScore),fmt(x.profitScore),fmt(x.profitBonus,0),x.epsStreak||0,fmt(x.roeScore),fmt(x.roeBonus,0),x.roeStreak||0,signed(x.epsRev),fmt(x.eps27,3),fmt(x.valuationScore),fmt(x.crowdRaw),fmt(x.prevCrowdRaw),fmt(x.crowdChange),x.crowdRisk,fmt(x.fundWeightChange),fmt(x.overweightChange),fmt(x.fundCount),fmt(x.fundWeight),fmt(x.fundWeightA),fmt(x.marketCapPct),fmt(x.overweight),fmt(x.freeFloat),x.method,fmt(x.valPct),x.digestLabel,fmt(x.digestScore),fmt(x.pe26),fmt(x.pe27),fmt(x.pe28),x.divStatus,x.unlockDate,fmt(x.unlockRatio),fmt(x.unlockAmount)])])}
function exportIndustries(){csvDownload('行业性价比评分_20261113.csv',[['行业','代码','综合分','价格分','盈利分','盈利加分','EPS连续上修(次)','ROE分','ROE加分','ROE连续上修(次)','EPS周度上修(%)','估值分','公募拥挤度','26Q2拥挤度','拥挤度环比','拥挤风险','公募配置环比(pct)','公募超配环比(pct)','公募A股配置占比','流通市值占比','公募超配(pct)','超配得分','历史分位(2010以来)','当前仓位%(周报)','930以来主动(pct)','估值法','估值分位','ROE周期'],...filteredIndustries().map(x=>[x.name,x.code,fmt(x.totalScore),fmt(x.priceScore),fmt(x.profitScore),fmt(x.profitBonus,0),x.epsStreak||0,fmt(x.roeScore),fmt(x.roeBonus,0),x.roeStreak||0,signed(x.epsRev),fmt(x.valuationScore),fmt(x.crowdRaw),fmt(x.prevCrowdRaw),fmt(x.crowdChange),x.crowdRisk,fmt(x.fundWeightChange),fmt(x.overweightChange),fmt(x.fundWeightA),fmt(x.marketCapPct),fmt(x.overweight),fmt(x.overweightScore),fmt(x.histPct),x.posCurrent!=null?fmt(x.posCurrent,2):'',x.posActive!=null?signed(x.posActive):'',x.method,fmt(x.valPct),x.cycle])])}
init();</script></body></html>'''
# 模板日期令牌：全部由 ASOF/源文件名/posweek 文件名派生，换数据源无需手改
_posweek_s = POS_WEEK_SOURCE.stem.replace("posweek_", "")
_posweek_s = f"{_posweek_s[:4]}-{_posweek_s[4:6]}-{_posweek_s[6:]}"
_cutoff_s = (ASOF.replace(year=ASOF.year - 1)).strftime("%Y-%m-%d")
_unlock_s = add_two_months(ASOF).strftime("%Y-%m-%d")
HTML = (HTML
    .replace("行业性价比评分_20261113.csv", f"行业性价比评分_{ASOF.strftime('%Y%m%d')}.csv")
    .replace("公募基金行业仓位周报2026-11-13", f"公募基金行业仓位周报{_posweek_s}")
    .replace("公募基金行业仓位周报（2026-11-13）", f"公募基金行业仓位周报（{_posweek_s}）")
    .replace("多维数据比较1113.xlsx", SOURCE.name)
    .replace("2026-11-13", _asof_s)
    .replace("2025-11-13", _cutoff_s)
    .replace("2027-01-02", _unlock_s))
assert "2026-11-13" not in HTML and "2025-11-13" not in HTML and "1113.xlsx" not in HTML
OUTPUT.write_text(HTML.replace("__DATA__", DATA_JSON).replace("__SUMMARY__", SUMMARY_JSON), encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False, indent=2))
print(OUTPUT)
