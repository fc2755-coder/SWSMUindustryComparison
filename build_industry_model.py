import json
import math
import statistics
from datetime import datetime
from pathlib import Path
import openpyxl

SOURCE = Path(r"C:/Users/Owner/Desktop/新建 XLSX 工作表 (2).xlsx")
OUTPUT = Path(r"C:/Users/Owner/Desktop/行业比较与个股筛选模型_20261025.html")
ASOF = datetime(2026, 10, 25)
PB_INDUSTRIES = {"银行", "非银金融", "房地产", "钢铁", "煤炭", "石油石化", "有色金属", "基础化工", "建筑材料", "建筑装饰", "交通运输"}


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
    return datetime(d.year + (1 if d.month == 12 else 0), 1 if d.month == 12 else d.month + 2, min(d.day, 28))


def roe_cycle(values):
    vals = [num(x) for x in values]
    r25, r26, r27 = vals[4], vals[5], vals[6]
    if r26 is None or r27 is None:
        return "数据不足", None, False
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
    return label, delta, delta >= -0.1


def revision(cur, prev):
    cur, prev = num(cur), num(prev)
    if cur is None or prev is None:
        return None, False
    rev = (cur - prev) / abs(prev) * 100 if abs(prev) > 1e-12 else None
    return rev, cur > prev


wb = openpyxl.load_workbook(SOURCE, data_only=True, read_only=False)
ws = wb["多维比较"]
headers = [ws.cell(3, c).value for c in range(1, 57)]
assert headers[18] == "27E EPS（上月）" and headers[26] == "当前PE分位数" and headers[40] == "当前PB分位数"
assert parse_date(ws["A1"].value).date() == ASOF.date()


def unpack(row, level, parent=None):
    code, name = str(row[0] or ""), str(row[1] or "")
    industry = parent or (str(row[2] or "") if level == "stock" else name)
    rev, eps_up = revision(row[17], row[18])
    roe_values = [row[27], row[28], row[29], row[30], row[31], row[32], row[34], row[36]]
    cycle, roe_delta, roe_ok = roe_cycle(roe_values[:7])
    method = "PB" if industry in PB_INDUSTRIES else "PE"
    val_pct = num(row[40]) if method == "PB" else num(row[26])
    val_multiple = num(row[38]) if method == "PB" else num(row[24])
    val_pass = val_pct is not None and 0 <= val_pct < 50 and val_multiple is not None and val_multiple > 0
    unlock_date = parse_date(row[53])
    unlock_ratio = num(row[55])
    unlock_alert = bool(unlock_date and ASOF <= unlock_date <= add_two_months(ASOF) and unlock_ratio is not None and unlock_ratio > 10)
    coverage = num(row[52])
    coverage = coverage * 100 if coverage is not None else None
    return {
        "code": code, "name": name, "level": level, "industry": industry, "subindustry": str(row[3] or ""),
        "month": num(row[7]), "eps27": num(row[17]), "eps27m": num(row[18]), "epsRev": rev, "epsUp": eps_up,
        "roe": [num(x) for x in roe_values], "roe26": num(row[32]), "roe27": num(row[34]), "roeDelta": roe_delta,
        "cycle": cycle, "roeOk": roe_ok, "method": method, "valPct": val_pct, "valMultiple": val_multiple,
        "pePct": num(row[26]), "pbPct": num(row[40]), "pe27": num(row[24]), "pb27": num(row[38]), "valPass": val_pass,
        "staticYield": num(row[41]), "yield27": pct(row[50]), "payouts": [pct(row[43]), pct(row[44]), pct(row[45])],
        "stable": str(row[48] or ""), "coverage": coverage, "unlockDate": unlock_date.strftime("%Y-%m-%d") if unlock_date else "",
        "unlockRatio": unlock_ratio, "unlockAlert": unlock_alert, "mcap": num(row[4]), "price": num(row[5])
    }


industries = []
for r in range(20, 185):
    row = [ws.cell(r, c).value for c in range(1, 57)]
    item = unpack(row, "L1" if r <= 50 else "L2")
    item["parent"] = item["industry"] if r <= 50 else ""
    industries.append(item)

stocks = []
for row in ws.iter_rows(min_row=185, max_row=ws.max_row, max_col=56, values_only=True):
    code = str(row[0] or "")
    if not code.endswith((".SH", ".SZ", ".BJ")):
        continue
    item = unpack(row, "stock")
    stocks.append(item)

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

industry_yield = {x["name"]: x["staticYield"] for x in industries if x["level"] == "L1"}
for s in stocks:
    trig = (s["staticYield"] is not None and s["staticYield"] > 4) or (industry_yield.get(s["industry"]) or 0) > 4
    s["divTrigger"] = trig
    s["stablePass"] = s["stable"] == "稳定"
    s["coveragePass"] = s["coverage"] >= 200 if s["coverage"] is not None and s["coverage"] > 0 else None
    s["yieldPass"] = s["yield27"] is not None and s["yield27"] > 3.5
    s["divPass"] = (s["yieldPass"] and s["stablePass"] and s["coveragePass"] is True) if trig else True
    s["divStatus"] = "非红利约束" if not trig else ("通过" if s["divPass"] else ("未通过" if not s["yieldPass"] else ("待补数据" if s["coveragePass"] is None else "未通过")))


def assign_month_percentile(items):
    valid = sorted([x for x in items if x["month"] is not None], key=lambda x: x["month"])
    n = len(valid)
    for i, x in enumerate(valid):
        x["monthPct"] = 100 * i / max(1, n - 1)
    for x in items:
        if "monthPct" not in x:
            x["monthPct"] = None


assign_month_percentile(stocks)
for level in ("L1", "L2"):
    assign_month_percentile([x for x in industries if x["level"] == level])

for s in stocks:
    s["basePass"] = all([s["monthPct"] is not None and s["monthPct"] <= 20, s["epsUp"], s["roeOk"], s["valPass"]])
    s["fullPass"] = s["basePass"] and s["divPass"]
for x in industries:
    x["oversold"] = x["monthPct"] is not None and x["monthPct"] <= 100 / 3
    x["basePass"] = all([x["oversold"], x["epsUp"], x["roeOk"], x["valPass"]])

for level in ("L1", "L2"):
    members = [x for x in industries if x["level"] == level]
    for x in members:
        if level == "L1":
            children = [s for s in stocks if s["industry"] == x["name"]]
        else:
            children = [s for s in stocks if s["subindustry"] == x["name"]]
        dividend = [s for s in children if s["divTrigger"]]
        x["stockCount"] = len(children)
        x["candidateCount"] = sum(1 for s in children if s["fullPass"])
        x["unlockCount"] = sum(1 for s in children if s["unlockAlert"])
        x["dividendCount"] = len(dividend)
        x["dividendQualityPct"] = 100 * sum(1 for s in dividend if s["divPass"]) / len(dividend) if dividend else None
        x["divTrigger"] = x["staticYield"] is not None and x["staticYield"] > 4

payload = {
    "asof": "2026-10-25", "monthStart": "2026-09-25", "unlockEnd": add_two_months(ASOF).strftime("%Y-%m-%d"),
    "industries": industries, "stocks": stocks, "pbIndustries": sorted(PB_INDUSTRIES)
}

summary = {
    "stocks": len(stocks), "industries_l1": sum(x["level"] == "L1" for x in industries),
    "industries_l2": sum(x["level"] == "L2" for x in industries),
    "base_candidates": sum(s["basePass"] for s in stocks), "full_candidates": sum(s["fullPass"] for s in stocks),
    "dividend_stocks": sum(s["divTrigger"] for s in stocks), "unlock_alerts": sum(s["unlockAlert"] for s in stocks),
    "industry_full_l1": [x["name"] for x in industries if x["level"] == "L1" and x["basePass"]],
    "top_candidates": [(s["code"], s["name"], s["industry"], round(s["month"], 2), round(s["epsRev"], 2) if s["epsRev"] is not None else None, s["cycle"], round(s["valPct"], 1)) for s in sorted([s for s in stocks if s["fullPass"]], key=lambda z: (z["monthPct"], -(z["epsRev"] or -999)))[:20]]
}

DATA_JSON = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
SUMMARY_JSON = json.dumps(summary, ensure_ascii=False)

HTML = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>行业比较与个股筛选模型 · 2026-10-25</title>
<style>
:root{--ink:#17233b;--muted:#687386;--line:#e4e8ef;--paper:#f5f6f8;--card:#fff;--navy:#172b4d;--red:#c83c45;--redbg:#fff1f1;--green:#16826c;--greenbg:#edf8f4;--amber:#a76512;--amberbg:#fff7e8;--blue:#3467c9;--bluebg:#eff4ff}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}button,input,select{font:inherit}.wrap{max-width:1480px;margin:auto;padding:0 28px 48px}.hero{background:linear-gradient(118deg,#14243f,#213d64);color:#fff;padding:34px 0 26px}.hero .wrap{padding-bottom:0}.eyebrow{font-size:12px;letter-spacing:2px;color:#b8c9e4}.hero h1{font-size:30px;margin:7px 0 8px;letter-spacing:.5px}.hero p{margin:0;color:#d6e0ef}.asof{margin-left:auto;text-align:right;font-size:12px;color:#c4d1e4}.asof b{display:block;font-size:20px;color:#fff}.heroFlex{display:flex;align-items:center}.nav{display:flex;gap:6px;margin-top:25px}.nav button{border:1px solid #ffffff25;background:transparent;color:#dce6f4;padding:9px 16px;border-radius:7px;cursor:pointer}.nav button.active{background:#fff;color:var(--navy);font-weight:700}.notice{background:#fff8e9;border:1px solid #f0dfbd;color:#72501b;padding:12px 16px;border-radius:9px;margin:20px 0}.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin:20px 0}.card{background:var(--card);border:1px solid var(--line);border-radius:11px;padding:17px}.metric{font-size:28px;font-weight:750;line-height:1.15;margin:5px 0}.label{font-size:12px;color:var(--muted)}.sub{font-size:12px;color:var(--muted)}.grid2{display:grid;grid-template-columns:1.55fr 1fr;gap:16px}.panel{background:#fff;border:1px solid var(--line);border-radius:11px;padding:20px;margin-bottom:16px}.panel h2{font-size:17px;margin:0 0 4px}.panelHead{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:14px}.legend{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:var(--muted)}.dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:5px}.up{color:var(--red)}.down{color:var(--green)}.tag{display:inline-flex;align-items:center;border-radius:5px;padding:2px 7px;font-size:12px;white-space:nowrap;background:#f0f2f5;color:#556174}.tag.pass{background:var(--redbg);color:var(--red)}.tag.ok{background:var(--greenbg);color:var(--green)}.tag.warn{background:var(--amberbg);color:var(--amber)}.tag.info{background:var(--bluebg);color:var(--blue)}.toolbar{display:flex;gap:9px;flex-wrap:wrap;align-items:center;margin:12px 0}.toolbar input,.toolbar select{border:1px solid #d7dde7;background:#fff;border-radius:7px;padding:8px 10px;color:var(--ink)}.toolbar input{min-width:210px}.btn{border:1px solid #d5dce7;background:#fff;color:var(--ink);border-radius:7px;padding:8px 12px;cursor:pointer}.btn.primary{background:var(--navy);color:#fff;border-color:var(--navy)}.tableWrap{overflow:auto;border:1px solid var(--line);border-radius:8px}table{border-collapse:collapse;width:100%;white-space:nowrap;font-size:13px}th{position:sticky;top:0;background:#f6f8fb;color:#596579;text-align:right;font-weight:600;padding:10px;border-bottom:1px solid var(--line);z-index:1}td{padding:10px;border-bottom:1px solid #edf0f4;text-align:right}th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left}tbody tr:hover{background:#fafbfc}.click{cursor:pointer}.click:hover{color:var(--blue);text-decoration:underline}.check{font-weight:700}.check.yes{color:var(--red)}.check.no{color:#a0a8b5}.scatter{width:100%;height:390px}.scatter text{font-size:11px;fill:#748094}.scatter .axis{stroke:#cbd2dc;stroke-width:1}.scatter .zero{stroke:#9ca8b8;stroke-dasharray:4 4}.scatter circle{cursor:pointer;stroke:#fff;stroke-width:1.5}.cycleBars{display:grid;gap:13px;margin-top:20px}.barRow{display:grid;grid-template-columns:72px 1fr 30px;gap:10px;align-items:center}.barTrack{height:10px;background:#eef1f5;border-radius:8px;overflow:hidden}.barFill{height:100%;border-radius:8px}.flow{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;margin:18px 0}.flowItem{border:1px solid var(--line);border-radius:8px;padding:13px;background:#fafbfc}.flowItem b{display:block;font-size:13px;margin-bottom:5px}.flowItem span{font-size:12px;color:var(--muted)}.page{display:none}.page.active{display:block}.foot{color:var(--muted);font-size:12px;margin-top:18px}.drawer{position:fixed;right:-520px;top:0;width:500px;max-width:95vw;height:100vh;background:#fff;box-shadow:-12px 0 35px #17233b20;z-index:10;padding:25px;overflow:auto;transition:right .2s}.drawer.open{right:0}.overlay{display:none;position:fixed;inset:0;background:#14243f45;z-index:9}.overlay.open{display:block}.close{float:right;border:0;background:#eef1f5;border-radius:50%;width:30px;height:30px;cursor:pointer}.detailGrid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:18px 0}.detailGrid .card{padding:12px}.miniChart{width:100%;height:210px}.rules{counter-reset:rule}.rule{position:relative;padding:14px 14px 14px 52px;border-bottom:1px solid var(--line)}.rule:before{counter-increment:rule;content:counter(rule);position:absolute;left:12px;top:14px;width:26px;height:26px;border-radius:50%;background:var(--navy);color:#fff;text-align:center;line-height:26px}.rule b{font-size:14px}.rule p{margin:4px 0 0;color:var(--muted);font-size:13px}.empty{text-align:center;padding:35px;color:var(--muted)}.pager{display:flex;justify-content:flex-end;gap:8px;align-items:center;margin-top:12px}.topList{display:grid;gap:8px}.topItem{display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #edf0f4;padding:8px 0}.rank{width:24px;height:24px;background:#eff2f6;border-radius:6px;text-align:center;line-height:24px;color:#637087}.nowrap{white-space:nowrap}
@media(max-width:1000px){.cards{grid-template-columns:repeat(3,1fr)}.grid2{grid-template-columns:1fr}.flow{grid-template-columns:repeat(3,1fr)}.heroFlex{display:block}.asof{text-align:left;margin-top:15px}}@media(max-width:640px){.wrap{padding:0 14px 30px}.hero h1{font-size:24px}.cards{grid-template-columns:repeat(2,1fr)}.flow{grid-template-columns:1fr 1fr}.nav{overflow:auto}.panel{padding:14px}.scatter{height:320px}}
</style>
</head>
<body>
<header class="hero"><div class="wrap"><div class="heroFlex"><div><div class="eyebrow">INDUSTRY COMPARISON · STOCK SCREEN</div><h1>行业比较与个股筛选模型</h1><p>超跌 × 盈利上修 × ROE周期 × 估值分位｜红利质量与解禁风险独立校验</p></div><div class="asof">数据截面<b>2026-10-25</b>过去一月起点 2026-09-25</div></div><nav class="nav"><button class="active" data-page="overview">行业总览</button><button data-page="stocks">个股筛选</button><button data-page="dividends">红利质量</button><button data-page="unlocks">解禁预警</button><button data-page="method">模型口径</button></nav></div></header>
<main class="wrap">
<div class="notice"><b>读数提示：</b>超跌、上修、ROE、估值均按原表最新数据直接计算；“完整候选”已执行红利约束，解禁&gt;10%仅预警不剔除。PE/PB历史分位直接采用源表“当前分位数”，非重新回算。</div>
<section id="overview" class="page active">
<div class="cards" id="summaryCards"></div>
<div class="flow"><div class="flowItem"><b>① 超跌</b><span>行业月跌幅前1/3<br>个股默认前20%</span></div><div class="flowItem"><b>② 盈利上修</b><span>27E EPS最新值<br>&gt; 上月值</span></div><div class="flowItem"><b>③ ROE非下行</b><span>27E－26E<br>≥－0.1pct</span></div><div class="flowItem"><b>④ 低估值分位</b><span>行业适配PE/PB<br>2016年以来&lt;50%</span></div><div class="flowItem"><b>⑤ 红利质量</b><span>27E股息率＞3.5%<br>分红稳定且OCF/DPS≥200%</span></div><div class="flowItem"><b>⑥ 解禁预警</b><span>未来2个月<br>解禁比例&gt;10%</span></div></div>
<div class="grid2"><div class="panel"><div class="panelHead"><div><h2>一级行业：超跌与盈利上修象限</h2><div class="sub">横轴为过去一月涨跌幅，纵轴为27E EPS月度调整；点击行业可查看成分股</div></div><div class="legend"><span><i class="dot" style="background:var(--red)"></i>EPS上修</span><span><i class="dot" style="background:var(--green)"></i>EPS下修/持平</span></div></div><div id="industryScatter" class="scatter"></div></div>
<div class="panel"><h2>一级行业 ROE 周期分布</h2><div class="sub">依据21—27E ROE路径分类</div><div id="cycleBars" class="cycleBars"></div><div style="margin-top:25px"><h2>四项条件同时通过</h2><div id="industryPassList" class="topList"></div></div></div></div>
<div class="panel"><div class="panelHead"><div><h2>行业比较明细</h2><div class="sub">四项条件：超跌、27E EPS上修、ROE非下行、适配估值分位&lt;50%</div></div></div><div class="toolbar"><select id="indLevel"><option value="L1">申万一级</option><option value="L2">申万二级</option></select><select id="indFilter"><option value="all">全部行业</option><option value="pass">四项通过</option><option value="oversoldUp">超跌且上修</option><option value="dividend">红利行业（股息率&gt;4%）</option></select><input id="indSearch" placeholder="搜索行业名称 / 代码"><button class="btn" id="exportInd">导出当前行业CSV</button></div><div class="tableWrap"><table><thead><tr><th>行业</th><th>代码</th><th>近1月</th><th>27E EPS月调</th><th>ROE周期</th><th>27E ROE</th><th>估值法</th><th>历史分位</th><th>四项</th><th>股息率</th><th>红利质量覆盖</th><th>候选股</th><th>解禁预警</th></tr></thead><tbody id="indBody"></tbody></table></div></div>
</section>
<section id="stocks" class="page">
<div class="panel"><div class="panelHead"><div><h2>个股筛选器</h2><div class="sub">完整候选 = 前四项 + 红利股质量约束；解禁仅标注，不作为剔除项</div></div><div><span class="tag info" id="resultCount"></span></div></div><div class="toolbar"><input id="stockSearch" placeholder="搜索代码 / 简称 / 行业"><select id="stockIndustry"><option value="">全部一级行业</option></select><select id="oversoldThreshold"><option value="10">超跌前10%</option><option value="20" selected>超跌前20%</option><option value="30">超跌前30%</option><option value="100">不限跌幅排名</option></select><select id="stockMode"><option value="full">完整候选</option><option value="base">前四项通过</option><option value="all">全部A股</option><option value="unlock">解禁预警</option></select><select id="cycleFilter"><option value="">全部ROE周期</option><option>底部反弹</option><option>持续向上</option><option>修复向上</option><option>高位平稳</option><option>低位平稳</option><option>拐头向下</option></select><button class="btn primary" id="exportStocks">导出筛选CSV</button></div><div class="tableWrap"><table><thead><tr><th>代码</th><th>简称</th><th>一级/二级行业</th><th>近1月</th><th>跌幅分位</th><th>27E EPS月调</th><th>ROE周期</th><th>估值法/分位</th><th>红利约束</th><th>27E股息率</th><th>解禁</th></tr></thead><tbody id="stockBody"></tbody></table></div><div class="pager"><button class="btn" id="prevPage">上一页</button><span id="pageLabel" class="sub"></span><button class="btn" id="nextPage">下一页</button></div></div>
</section>
<section id="dividends" class="page"><div class="panel"><h2>红利股质量筛查</h2><div class="sub">个股或所属一级行业静态股息率&gt;4%即触发；通过须同时满足27E股息率＞3.5%、分红稳定、27E OCF/DPS≥200%</div><div class="cards" style="grid-template-columns:repeat(4,1fr);max-width:960px" id="divCards"></div><div class="toolbar"><select id="divFilter"><option value="all">全部红利约束股</option><option value="pass">质量通过</option><option value="fail">未通过</option><option value="missing">OCF待补</option></select></div><div class="tableWrap"><table><thead><tr><th>简称</th><th>行业</th><th>静态股息率</th><th>27E股息率</th><th>分红</th><th>OCF/DPS</th><th>结论</th></tr></thead><tbody id="divBody"></tbody></table></div></div></section>
<section id="unlocks" class="page"><div class="panel"><h2>未来两个月解禁预警</h2><div class="sub">2026-10-25至2026-12-25，解禁占总股本比例&gt;10%；仅作风险标注，不参与红利质量判断</div><div id="unlockList" style="margin-top:14px"></div></div></section>
<section id="method" class="page"><div class="panel"><h2>六条规则的计算口径</h2><div class="rules"><div class="rule"><b>过去一个月跌幅靠前</b><p>用“过去一月涨跌幅”在同层样本内升序排名。个股默认取前20%，可切换10%/30%；一级与二级行业分别取各自样本前1/3。负收益越多，排名越靠前。</p></div><div class="rule"><b>过去一个月27E EPS上修</b><p>最新27E EPS＞上月27E EPS。展示调整幅度 =(最新－上月)/|上月|；判定上修只看两者大小，避免负EPS导致方向误读。</p></div><div class="rule"><b>ROE处在非下行期</b><p>27E ROE－26E ROE≥－0.1个百分点。若27E较26E下降超过0.1pct，标为拐头向下；结合21—27E路径识别底部反弹、持续向上、修复向上及高/低位平稳。</p></div><div class="rule"><b>2016年以来估值分位低于50%</b><p>直接采用源表当前PE/PB分位。金融、地产及强周期行业使用PB，其余使用PE；同时要求27E对应估值倍数&gt;0，避免负盈利下PE失真。</p></div><div class="rule"><b>红利板块的额外约束</b><p>个股静态股息率或所属一级行业静态股息率&gt;4%即触发。要求27E股息率＞3.5%、“分红是否稳定=稳定”且27E每股OCF/DPS≥200%；27E股息率≤3.5%直接不通过，OCF为0或空值视为待补，不当作通过。</p></div><div class="rule"><b>未来两个月限售解禁</b><p>下个解禁日处于2026-10-25至2026-12-25且解禁占总股本比例&gt;10%时预警，仅标注、不剔除。</p></div></div></div><div class="panel"><h2>行业估值方法映射</h2><p class="sub">PB：银行、非银金融、房地产、钢铁、煤炭、石油石化、有色金属、基础化工、建筑材料、建筑装饰、交通运输。其余行业使用PE。</p><p class="sub">行业红利质量覆盖率 = 行业内触发红利约束且质量通过的股票数 / 行业内触发红利约束股票数；行业原始行没有分红稳定性及OCF字段，因此不从行业行臆造。</p><p class="sub">本模型是规则化研究筛选，不构成投资建议。数据日期为周末时，行情字段对应数据源最近有效收盘。</p></div></section>
<div class="foot">数据源：新建 XLSX 工作表 (2).xlsx · 多维比较｜计算截面：2026-10-25｜离线单文件，无需外部图表库</div>
</main>
<div class="overlay" id="overlay"></div><aside class="drawer" id="drawer"></aside>
<script id="modelData" type="application/json">__DATA__</script>
<script>
const DATA=JSON.parse(document.getElementById('modelData').textContent);
const SUMMARY=__SUMMARY__;
const $=id=>document.getElementById(id);
const fmt=(v,d=2)=>v===null||v===undefined||Number.isNaN(v)?'—':Number(v).toFixed(d);
const sign=v=>v===null||v===undefined?'—':(v>0?'+':'')+fmt(v,2)+'%';
const cls=v=>v>0?'up':v<0?'down':'';
const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const tag=(text,type='')=>`<span class="tag ${type}">${esc(text)}</span>`;
const yn=b=>`<span class="check ${b?'yes':'no'}">${b?'✓':'—'}</span>`;
const valText=x=>`${x.method} ${fmt(x.valPct,1)}%`;
let stockPage=1,pageSize=60,currentStocks=[];
const cycleColors={'底部反弹':'#c83c45','持续向上':'#df7b56','修复向上':'#d6a043','高位平稳':'#3467c9','低位平稳':'#8290a5','拐头向下':'#16826c','数据不足':'#b5bdc9'};
function init(){
  $('summaryCards').innerHTML=[['A股样本',SUMMARY.stocks,'不含港股及行业指数'],['四项通过',SUMMARY.base_candidates,'超跌前20%口径'],['完整候选',SUMMARY.full_candidates,'含红利质量约束'],['红利约束股',SUMMARY.dividend_stocks,'个股或行业股息率>4%'],['解禁预警',SUMMARY.unlock_alerts,'未来2个月且比例>10%'],['一级行业通过',SUMMARY.industry_full_l1.length,'四项条件同时满足']].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="metric">${x[1]}</div><div class="sub">${x[2]}</div></div>`).join('');
  const l1=DATA.industries.filter(x=>x.level==='L1');
  $('stockIndustry').innerHTML+=[...new Set(DATA.stocks.map(x=>x.industry))].sort().map(x=>`<option>${esc(x)}</option>`).join('');
  renderScatter(l1);renderCycles(l1);renderIndustries();renderStocks();renderRisks();
  $('industryPassList').innerHTML=l1.filter(x=>x.basePass).length?l1.filter(x=>x.basePass).map((x,i)=>`<div class="topItem"><div style="display:flex;gap:9px;align-items:center"><span class="rank">${i+1}</span><b class="click" onclick="openIndustry('${esc(x.name)}')">${esc(x.name)}</b></div><div>${tag(x.cycle,'info')} ${tag(valText(x),'pass')}</div></div>`).join(''):'<div class="empty">当前没有一级行业同时通过四项条件</div>';
  document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>{document.querySelectorAll('.nav button,.page').forEach(x=>x.classList.remove('active'));b.classList.add('active');$(b.dataset.page).classList.add('active')});
  ['indLevel','indFilter','indSearch'].forEach(id=>$(id).oninput=renderIndustries);
  ['stockSearch','stockIndustry','oversoldThreshold','stockMode','cycleFilter'].forEach(id=>$(id).oninput=()=>{stockPage=1;renderStocks()});
  $('prevPage').onclick=()=>{if(stockPage>1){stockPage--;renderStocks()}};$('nextPage').onclick=()=>{if(stockPage*pageSize<currentStocks.length){stockPage++;renderStocks()}};
  $('exportStocks').onclick=exportStocks;$('exportInd').onclick=exportIndustries;$('divFilter').oninput=renderRisks;
  $('overlay').onclick=closeDrawer;
}
function renderScatter(items){
  const valid=items.filter(x=>x.month!==null&&x.epsRev!==null);const W=760,H=390,P=48;
  const xmin=Math.min(...valid.map(x=>x.month),-1),xmax=Math.max(...valid.map(x=>x.month),1),ymin=Math.min(...valid.map(x=>x.epsRev),-1),ymax=Math.max(...valid.map(x=>x.epsRev),1);
  const x=v=>P+(v-xmin)/(xmax-xmin)*(W-2*P),y=v=>H-P-(v-ymin)/(ymax-ymin)*(H-2*P);
  let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="100%"><line class="axis" x1="${P}" y1="${H-P}" x2="${W-P}" y2="${H-P}"/><line class="axis" x1="${P}" y1="${P}" x2="${P}" y2="${H-P}"/>`;
  if(xmin<0&&xmax>0)s+=`<line class="zero" x1="${x(0)}" y1="${P}" x2="${x(0)}" y2="${H-P}"/>`;if(ymin<0&&ymax>0)s+=`<line class="zero" x1="${P}" y1="${y(0)}" x2="${W-P}" y2="${y(0)}"/>`;
  for(let i=0;i<=4;i++){let xv=xmin+(xmax-xmin)*i/4,yv=ymin+(ymax-ymin)*i/4;s+=`<text x="${x(xv)}" y="${H-P+20}" text-anchor="middle">${fmt(xv,0)}%</text><text x="${P-8}" y="${y(yv)+4}" text-anchor="end">${fmt(yv,0)}%</text>`}
  valid.forEach(d=>{let r=6+Math.max(0,Math.min(7,(d.roe27||0)/4));s+=`<circle cx="${x(d.month)}" cy="${y(d.epsRev)}" r="${r}" fill="${d.epsUp?'#c83c45':'#16826c'}" fill-opacity=".78" onclick="openIndustry('${esc(d.name)}')"><title>${esc(d.name)}｜近1月 ${sign(d.month)}｜EPS月调 ${sign(d.epsRev)}｜${d.cycle}</title></circle><text x="${x(d.month)+r+3}" y="${y(d.epsRev)-r}" style="font-size:10px">${esc(d.name)}</text>`});
  s+=`<text x="${W-P}" y="${H-8}" text-anchor="end">过去一月涨跌幅 →</text><text x="14" y="${P}" transform="rotate(-90 14 ${P})" text-anchor="end">27E EPS月度调整 →</text></svg>`;$('industryScatter').innerHTML=s;
}
function renderCycles(items){const counts={};items.forEach(x=>counts[x.cycle]=(counts[x.cycle]||0)+1);const order=['底部反弹','持续向上','修复向上','高位平稳','低位平稳','拐头向下','数据不足'];$('cycleBars').innerHTML=order.filter(k=>counts[k]).map(k=>`<div class="barRow"><span>${k}</span><div class="barTrack"><div class="barFill" style="width:${counts[k]/items.length*100}%;background:${cycleColors[k]}"></div></div><b>${counts[k]}</b></div>`).join('')}
function filteredIndustries(){const q=$('indSearch').value.trim().toLowerCase(),f=$('indFilter').value;return DATA.industries.filter(x=>x.level===$('indLevel').value&&(!q||(x.name+x.code).toLowerCase().includes(q))&&(f==='all'||f==='pass'&&x.basePass||f==='oversoldUp'&&x.oversold&&x.epsUp||f==='dividend'&&x.divTrigger)).sort((a,b)=>Number(b.basePass)-Number(a.basePass)||a.monthPct-b.monthPct)}
function renderIndustries(){const rows=filteredIndustries();$('indBody').innerHTML=rows.map(x=>`<tr><td><b>${esc(x.name)}</b> ${x.divTrigger?tag('红利','warn'):''}</td><td>${x.code}</td><td class="${cls(x.month)}">${sign(x.month)}</td><td class="${cls(x.epsRev)}">${sign(x.epsRev)}</td><td>${tag(x.cycle,x.roeOk?'info':'')}</td><td>${fmt(x.roe27,1)}%</td><td>${x.method}</td><td class="${x.valPass?'up':''}">${fmt(x.valPct,1)}%</td><td>${yn(x.basePass)}</td><td>${fmt(x.staticYield,2)}%</td><td>${x.dividendQualityPct===null?'—':fmt(x.dividendQualityPct,0)+'%'}</td><td><span class="click" onclick="openIndustry('${esc(x.name)}')">${x.candidateCount}/${x.stockCount}</span></td><td>${x.unlockCount?tag(x.unlockCount+'只','warn'):'—'}</td></tr>`).join('')||'<tr><td colspan="13" class="empty">无匹配行业</td></tr>'}
function filteredStocks(){const q=$('stockSearch').value.trim().toLowerCase(),ind=$('stockIndustry').value,cy=$('cycleFilter').value,mode=$('stockMode').value,t=Number($('oversoldThreshold').value);return DATA.stocks.filter(x=>{const base=x.monthPct!==null&&x.monthPct<=t&&x.epsUp&&x.roeOk&&x.valPass;return(!q||(x.code+x.name+x.industry+x.subindustry).toLowerCase().includes(q))&&(!ind||x.industry===ind)&&(!cy||x.cycle===cy)&&(mode==='all'||mode==='base'&&base||mode==='full'&&base&&x.divPass||mode==='unlock'&&x.unlockAlert)}).sort((a,b)=>Number(b.fullPass)-Number(a.fullPass)||a.monthPct-b.monthPct||(b.epsRev??-999)-(a.epsRev??-999))}
function renderStocks(){currentStocks=filteredStocks();const pages=Math.max(1,Math.ceil(currentStocks.length/pageSize));if(stockPage>pages)stockPage=pages;const rows=currentStocks.slice((stockPage-1)*pageSize,stockPage*pageSize);$('resultCount').textContent=`${currentStocks.length} 只`;$('pageLabel').textContent=`${stockPage} / ${pages} 页`;$('stockBody').innerHTML=rows.map(x=>`<tr><td>${x.code}</td><td><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b></td><td>${esc(x.industry)} / ${esc(x.subindustry)}</td><td class="${cls(x.month)}">${sign(x.month)}</td><td>${fmt(x.monthPct,1)}%</td><td class="${cls(x.epsRev)}">${sign(x.epsRev)}</td><td>${tag(x.cycle,x.roeOk?'info':'')}</td><td class="${x.valPass?'up':''}">${x.method} ${fmt(x.valPct,1)}%</td><td>${tag(x.divStatus,x.divStatus==='通过'?'ok':x.divStatus==='未通过'?'warn':'')}</td><td>${fmt(x.yield27,2)}%</td><td>${x.unlockAlert?tag(fmt(x.unlockRatio,1)+'% · '+x.unlockDate,'warn'):'—'}</td></tr>`).join('')||'<tr><td colspan="11" class="empty">当前条件下无匹配股票</td></tr>'}
function renderRisks(){const div=DATA.stocks.filter(x=>x.divTrigger),passed=div.filter(x=>x.divPass),missing=div.filter(x=>!x.divPass&&x.coveragePass===null&&x.yieldPass),lowYield=div.filter(x=>!x.yieldPass);$('divCards').innerHTML=[['红利约束股',div.length],['质量通过',passed.length],['27E股息率≤3.5%',lowYield.length],['OCF待补',missing.length]].map(x=>`<div class="card"><div class="label">${x[0]}</div><div class="metric">${x[1]}</div></div>`).join('');const f=$('divFilter').value;const rows=div.filter(x=>f==='all'||f==='pass'&&x.divPass||f==='fail'&&!x.divPass&&!(x.coveragePass===null&&x.yieldPass)||f==='missing'&&!x.divPass&&x.coveragePass===null&&x.yieldPass).sort((a,b)=>(b.yield27||0)-(a.yield27||0)).slice(0,250);$('divBody').innerHTML=rows.map(x=>`<tr><td><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b></td><td>${esc(x.industry)}</td><td>${fmt(x.staticYield,2)}%</td><td class="${x.yieldPass?'up':'down'}"><b>${fmt(x.yield27,2)}%</b></td><td>${tag(x.stable||'待补',x.stablePass?'ok':'warn')}</td><td>${x.coveragePass===null?'—':fmt(x.coverage,0)+'%'}</td><td>${tag(x.divStatus,x.divPass?'ok':'warn')}</td></tr>`).join('')||'<tr><td colspan="7" class="empty">无匹配股票</td></tr>';const unlock=DATA.stocks.filter(x=>x.unlockAlert).sort((a,b)=>b.unlockRatio-a.unlockRatio);$('unlockList').innerHTML=unlock.length?unlock.map(x=>`<div class="topItem"><div><b class="click" onclick="openStock('${x.code}')">${esc(x.name)}</b><div class="sub">${x.code} · ${esc(x.industry)}</div></div><div style="text-align:right">${tag(fmt(x.unlockRatio,2)+'%','warn')}<div class="sub">${x.unlockDate}</div></div></div>`).join(''):'<div class="empty">没有触发阈值的解禁</div>'}
function openIndustry(name){document.querySelector('[data-page="stocks"]').click();$('stockIndustry').value=DATA.stocks.some(s=>s.industry===name)?name:'';$('stockSearch').value=DATA.stocks.some(s=>s.industry===name)?'':name;$('stockMode').value='all';stockPage=1;renderStocks()}
function openStock(code){const x=DATA.stocks.find(s=>s.code===code);if(!x)return;const years=['2021','2022','2023','2024','2025','2026E','2027E','2028E'];const vals=x.roe,valid=vals.filter(v=>v!==null),min=Math.min(...valid,0),max=Math.max(...valid,1),W=430,H=180,P=28;const y=v=>H-P-(v-min)/(max-min)*(H-2*P);let chart=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="100%"><line x1="${P}" y1="${y(0)}" x2="${W-P}" y2="${y(0)}" stroke="#ccd3dd"/>`;const pts=vals.map((v,i)=>v===null?null:`${P+i*(W-2*P)/7},${y(v)}`).filter(Boolean).join(' ');chart+=`<polyline points="${pts}" fill="none" stroke="#3467c9" stroke-width="2.5"/>`;vals.forEach((v,i)=>{if(v===null)return;let cx=P+i*(W-2*P)/7;chart+=`<circle cx="${cx}" cy="${y(v)}" r="4" fill="${i===6?'#c83c45':'#3467c9'}"/><text x="${cx}" y="${y(v)-9}" text-anchor="middle" font-size="10" fill="#596579">${fmt(v,1)}</text><text x="${cx}" y="${H-7}" text-anchor="middle" font-size="10" fill="#596579">${years[i]}</text>`});chart+='</svg>';$('drawer').innerHTML=`<button class="close" onclick="closeDrawer()">×</button><div class="eyebrow" style="color:#748094">STOCK DETAIL</div><h2>${esc(x.name)} <span class="sub">${x.code}</span></h2><div>${tag(x.cycle,'info')} ${tag(x.method+'分位 '+fmt(x.valPct,1)+'%',x.valPass?'pass':'')} ${x.unlockAlert?tag('解禁预警','warn'):''}</div><div class="detailGrid"><div class="card"><div class="label">近1月涨跌幅</div><div class="metric ${cls(x.month)}">${sign(x.month)}</div></div><div class="card"><div class="label">27E EPS月调</div><div class="metric ${cls(x.epsRev)}">${sign(x.epsRev)}</div></div><div class="card"><div class="label">27E ROE</div><div class="metric">${fmt(x.roe27,1)}%</div></div><div class="card"><div class="label">27E股息率</div><div class="metric">${fmt(x.yield27,2)}%</div></div></div><h3>ROE周期路径（%）</h3><div class="miniChart">${chart}</div><div class="panel" style="padding:14px"><b>规则检查</b><p>超跌前20% ${yn(x.monthPct<=20)}　EPS上修 ${yn(x.epsUp)}</p><p>ROE非下行 ${yn(x.roeOk)}　估值分位合格 ${yn(x.valPass)}</p><p>红利约束：${tag(x.divStatus,x.divPass?'ok':'warn')}　OCF/DPS：${x.coveragePass===null?'待补':fmt(x.coverage,0)+'%'}</p><p>分红稳定性：${esc(x.stable||'待补')}；23/24/25年分红率：${x.payouts.map(v=>fmt(v,1)+'%').join(' / ')}</p><p>解禁：${x.unlockDate||'无日期'}；比例：${fmt(x.unlockRatio,2)}%</p></div>`;$('drawer').classList.add('open');$('overlay').classList.add('open')}
function closeDrawer(){$('drawer').classList.remove('open');$('overlay').classList.remove('open')}
function csvDownload(name,rows){const csv='\ufeff'+rows.map(r=>r.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download=name;a.click();URL.revokeObjectURL(a.href)}
function exportStocks(){csvDownload('个股筛选_20261025.csv',[['代码','简称','一级行业','二级行业','近1月涨跌幅%','跌幅分位%','27E_EPS月调%','ROE周期','27E_ROE%','估值法','估值历史分位%','红利约束','27E股息率%','解禁日期','解禁比例%'],...currentStocks.map(x=>[x.code,x.name,x.industry,x.subindustry,fmt(x.month),fmt(x.monthPct),fmt(x.epsRev),x.cycle,fmt(x.roe27),x.method,fmt(x.valPct),x.divStatus,fmt(x.yield27),x.unlockDate,fmt(x.unlockRatio)])])}
function exportIndustries(){csvDownload('行业比较_20261025.csv',[['行业','代码','层级','近1月涨跌幅%','27E_EPS月调%','ROE周期','27E_ROE%','估值法','估值分位%','四项通过','静态股息率%','红利质量覆盖%','候选股数'],...filteredIndustries().map(x=>[x.name,x.code,x.level,fmt(x.month),fmt(x.epsRev),x.cycle,fmt(x.roe27),x.method,fmt(x.valPct),x.basePass?'是':'否',fmt(x.staticYield),fmt(x.dividendQualityPct),x.candidateCount])])}
init();
</script></body></html>'''
html = HTML.replace("__DATA__", DATA_JSON).replace("__SUMMARY__", SUMMARY_JSON)
OUTPUT.write_text(html, encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False, indent=2))
print(str(OUTPUT))
