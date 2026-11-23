# -*- coding: utf-8 -*-
"""双创指数（科创综指/创业板综）权重前30个股 × 2026Q3公募重仓持仓 对比页。
逻辑：每指数取权重最大30只，直观展示公募持仓谁相对较少（组内配置占比后1/3标"持仓偏低"）。
点击个股名弹出详情抽屉（盈利·估值·ROE，同主模型口径，仅展示不打分）。
权重：桌面《双创标的.xlsx》；公募与基本面：复用主模型 build_scoring_model.py 同口径计算结果。
模型外个股：公募用Q3重仓原始值在主模型全A分布上补算；基本面无同源数据，抽屉中说明。
"""
import io
import json
import bisect
from pathlib import Path

import openpyxl

ROOT = Path(r"C:/Users/Owner/WorkBuddy/2026-10-26-13-16-06")
WEIGHT_SOURCE = Path(r"C:/Users/Owner/Desktop/双创标的.xlsx")
MODEL_SCRIPT = ROOT / "build_scoring_model.py"
OUTPUT = Path(r"C:/Users/Owner/Desktop/双创权重前30_公募持仓.html")
TOP_N = 30


def num(v):
    return float(v) if isinstance(v, (int, float)) else None


def revision(cur, prev):
    cur, prev = num(cur), num(prev)
    if cur is None or prev is None:
        return None
    # 源表把"无最新盈利预测"写成0（26E/27E/28E同时归零），并非真实下修，按缺失处理。
    if cur == 0 and prev:
        return None
    # EPS跨零或同为负时百分比无意义，改用文字标注。
    if cur * prev < 0:
        return "转亏" if cur < 0 else "扭亏"
    if cur < 0 and prev < 0:
        if abs(cur) > abs(prev):
            return "亏损扩大"
        if abs(cur) < abs(prev):
            return "亏损收窄"
        return 0.0
    return (cur - prev) / abs(prev) * 100 if abs(prev) > 1e-12 else None


# ---------- 执行主模型脚本，取同口径个股数据（含基本面与Q3公募） ----------
src = io.open(MODEL_SCRIPT, encoding="utf-8").read()
src = src.replace(
    'OUTPUT = Path(r"C:/Users/Owner/Desktop/周报/行业比较与个股筛选模型_20261102.html")',
    'OUTPUT = Path(r"C:/Users/Owner/WorkBuddy/2026-10-26-13-16-06/_model_tmp_for_dualinnov.html")',
)
assert "_model_tmp_for_dualinnov" in src, "主模型OUTPUT重定向失败"
g = {"__name__": "__main__"}
exec(compile(src, str(MODEL_SCRIPT), "exec"), g)
model = {s["code"]: s for s in g["stocks"]}
q3 = g["current_holdings"]
print("主模型个股:", len(model))

dist = {}
for key in ("overweight", "fundWeightA", "fundCount", "freeFloat"):
    dist[key] = sorted(s[key] for s in model.values() if s.get(key) is not None)


def percentile_score(key, value):
    if value is None:
        return 0.0
    values = dist[key]
    n = len(values)
    if n <= 1:
        return 50.0
    left = bisect.bisect_left(values, value)
    right = bisect.bisect_right(values, value)
    rank = (left + right - 1) / 2
    return 100 * rank / (n - 1)


# ---------- 权重：每指数取前TOP_N ----------
wb = openpyxl.load_workbook(WEIGHT_SOURCE, data_only=True, read_only=True)
indices = {}
total_count = {}
for sheet in wb.sheetnames:
    name = "科创综指" if "科创" in sheet else "创业板综"
    rows = []
    for row in wb[sheet].iter_rows(min_row=3, max_col=4, values_only=True):
        code, cname, w = row[1], row[2], row[3]
        if isinstance(code, str) and code.endswith((".SH", ".SZ")) and isinstance(w, (int, float)):
            rows.append({"code": code.strip(), "name": str(cname), "weight": round(float(w), 4)})
    rows.sort(key=lambda r: -r["weight"])
    total_count[name] = len(rows)
    indices[name] = rows[:TOP_N]
wb.close()
constituent_codes = {r["code"] for rows in indices.values() for r in rows}

# ---------- 源表周度EPS：近三周各自的上修幅度（仅前30成分） ----------
weekly_rev = {}
swb = openpyxl.load_workbook(g["SOURCE"], data_only=True, read_only=True)
sws = swb["多维比较"]
for row in sws.iter_rows(min_row=185, max_col=23, values_only=True):
    code = str(row[0] or "")
    if code in constituent_codes:
        weekly_rev[code] = [revision(row[19], row[20]), revision(row[20], row[21]), revision(row[21], row[22])]
swb.close()

payload = {}
for idx_name, rows in indices.items():
    items = []
    for rank, r in enumerate(rows, 1):
        code = r["code"]
        m = model.get(code)
        if m:
            item = {
                "industry": m["industry"], "fundCount": m["fundCount"],
                "fundWeightA": m["fundWeightA"], "overweight": m["overweight"],
                "marketCapPct": m["marketCapPct"], "recentIPO": m["recentIPO"], "inModel": True,
                "hasQ3": m["hasFundHolding"], "ytd": m["ytd"], "since924": m["since924"],
                # 详情抽屉用的完整基本面（不打分）
                "eps": m["eps"], "eps27": m["eps27"], "epsGrowth": m["epsGrowth"],
                "epsStreak": m["epsStreak"], "roe": m["roe"], "roe27": m["roe27"],
                "roeDelta": m["roeDelta"], "roeRev": m["roeRev"], "cycle": m["cycle"],
                "method": m["method"], "valPct": m["valPct"], "pe26": m["pe26"],
                "pe27": m["pe27"], "pe28": m["pe28"], "pb27": m["pb27"],
                "digestYear": m["digestYear"], "digestLabel": m["digestLabel"], "digestScore": m["digestScore"],
                "month": m["month"], "marketCap": m["marketCap"],
            }
        else:
            h = q3.get(code)
            if h:
                item = {
                    "industry": h.get("l1") or "—", "fundCount": h["fundCount"],
                    "fundWeightA": h.get("fundWeightA"), "overweight": h.get("overweight"),
                    "marketCapPct": h.get("marketCapPct"), "recentIPO": False, "inModel": False,
                    "hasQ3": True, "ytd": None, "since924": None,
                }
            else:
                item = {"industry": "—", "fundCount": 0.0, "fundWeightA": None, "overweight": None,
                        "marketCapPct": None, "recentIPO": False, "inModel": False,
                        "hasQ3": False, "ytd": None, "since924": None}
            item.update({k: None for k in ("eps", "eps27", "epsGrowth", "epsStreak", "roe", "roe27", "roeDelta",
                                           "roeRev", "method", "valPct", "pe26", "pe27", "pe28", "pb27",
                                           "digestYear", "digestLabel", "digestScore", "month", "marketCap")})
            item["cycle"] = ""
        item.update(r)
        item["rank"] = rank
        item["revs"] = weekly_rev.get(code, [None, None, None])
        items.append(item)
    # 组内"持仓偏低"：公募配置占比最低的1/3（10只）标tag
    q3items = [x for x in items if x["hasQ3"]]
    fw_sorted = sorted(x["fundWeightA"] if x["fundWeightA"] is not None else -1e9 for x in q3items)
    cut = fw_sorted[max(0, len(fw_sorted) // 3 - 1)] if fw_sorted else -1e9
    fw_vals = [x["fundWeightA"] for x in items if x["fundWeightA"] is not None]
    fw_min, fw_max = min(fw_vals), max(fw_vals)
    for x in items:
        fw = x["fundWeightA"] if x["fundWeightA"] is not None else -1e9
        x["lowHold"] = bool(x["hasQ3"] and fw <= cut and cut > -1e8)
        # 组内0-100相对水平（用于迷你bar直观对比）
        x["fwBar"] = round(100 * (x["fundWeightA"] - fw_min) / (fw_max - fw_min), 1) if (fw_vals and x["fundWeightA"] is not None and fw_max > fw_min) else (100.0 if x["fundWeightA"] is not None else 0.0)
    payload[idx_name] = items

summary = {}
for idx_name, items in payload.items():
    summary[idx_name] = {
        "count": len(items), "total": total_count[idx_name],
        "weightSum": round(sum(x["weight"] for x in items), 2),
        "lowHold": sum(1 for x in items if x["lowHold"]),
        "noQ3": sum(1 for x in items if not x["hasQ3"]),
        "outModel": sum(1 for x in items if not x["inModel"]),
    }
    print(idx_name, json.dumps(summary[idx_name], ensure_ascii=False))

DATA_JSON = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
SUMMARY_JSON = json.dumps(summary, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>双创权重前30 × 公募持仓</title>
<style>
:root{--ink:#1b2430;--sub:#6b7684;--line:#e5e9ef;--bg:#f5f7fa;--card:#fff;--teal:#0e9f8a;--blue:#2f6fed;--amber:#c77d0a;--red:#d0453e;--green:#1d9e57}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;background:var(--bg);color:var(--ink);font-size:13px;line-height:1.55}
.wrap{max-width:1280px;margin:0 auto;padding:20px 24px 40px}
h1{font-size:22px;letter-spacing:.5px}.sub{color:var(--sub);font-size:12px;margin-top:4px}
.head{display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:10px;margin-bottom:16px}
.tabs{display:flex;gap:8px}.tab{padding:7px 18px;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:pointer;font-size:13px;color:var(--sub)}
.tab.active{background:var(--blue);border-color:var(--blue);color:#fff;font-weight:600}
.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 16px}
.card .v{font-size:22px;font-weight:700}.card .l{color:var(--sub);font-size:12px}
.panel{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:16px}
.panelHead{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px}
h2{font-size:15px}.toolbar{display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin:8px 0 12px;color:var(--sub);font-size:12px}
input[type=text]{padding:6px 10px;border:1px solid var(--line);border-radius:6px;width:180px;font-size:12px}
button,select{padding:6px 12px;border:1px solid var(--line);border-radius:6px;background:#fff;cursor:pointer;font-size:12px;color:var(--ink)}
button.primary{background:var(--blue);color:#fff;border-color:var(--blue)}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:right;color:var(--sub);font-weight:500;padding:6px 8px;border-bottom:1px solid var(--line);white-space:nowrap}
td{text-align:right;padding:7px 8px;border-bottom:1px solid var(--line);white-space:nowrap}
thead tr.grp th{text-align:center;font-weight:600;color:var(--ink);background:#f2f5f9;font-size:12px}
th.c1,td.c1,th.c2,td.c2,th.c3,td.c3,th.c0,td.c0{text-align:left}
thead{position:sticky;top:0;background:#fff;z-index:2}
.name{cursor:pointer}.name:hover{text-decoration:underline;color:var(--blue)}
.tag{display:inline-block;padding:1px 8px;border-radius:20px;font-size:11px}
.tag.blue{background:#e8effd;color:var(--blue)}.tag.teal{background:#e2f5f1;color:var(--teal)}.tag.amber{background:#fbf1de;color:var(--amber)}.tag.red{background:#fbe7e5;color:var(--red)}.tag.gray{background:#eef1f5;color:var(--sub)}
.up{color:var(--red)}.down{color:var(--green)}
.note{color:var(--sub);font-size:12px;line-height:1.9;background:#fff;border:1px dashed var(--line);border-radius:10px;padding:12px 16px}
.empty{color:var(--sub);text-align:center;padding:24px}
.count{font-size:12px;color:var(--sub)}
tr.lowRow{background:#fdf8ee}
tbody tr:hover{background:#f2f7ff}
.fwbar{display:inline-block;vertical-align:middle;width:70px;height:6px;background:#eef1f5;border-radius:3px;margin-left:8px;position:relative;overflow:hidden}
.fwbar i{position:absolute;left:0;top:0;bottom:0;background:#7ba7d7;border-radius:3px}
.fwbar.low i{background:var(--amber)}
.revsn{font-variant-numeric:tabular-nums}
/* 抽屉 */
#mask{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;z-index:50}
#drawer{position:fixed;top:0;right:-560px;width:540px;max-width:96vw;height:100vh;background:#fff;box-shadow:-6px 0 24px rgba(0,0,0,.15);z-index:51;transition:right .22s ease;display:flex;flex-direction:column}
#drawer.open{right:0}
.dHead{padding:16px 20px 12px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:flex-start}
.dHead h3{font-size:18px}.dBody{padding:14px 20px 30px;overflow-y:auto;flex:1}
.dClose{border:none;background:none;font-size:20px;color:var(--sub);cursor:pointer;padding:2px 8px}
.dCard{border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin-bottom:12px}
.dCard h4{font-size:13px;margin-bottom:8px;color:var(--ink)}
.kv{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px 12px;font-size:12px}
.kv .k{color:var(--sub)}.kv .v{font-weight:600;text-align:right}
.dHint{color:var(--sub);font-size:12px;margin:6px 0 10px}
</style></head><body><div class="wrap">
<div class="head"><div><h1>双创指数权重前30 × 公募持仓对比</h1>
<div class="sub">权重源：双创标的.xlsx｜公募：2026Q3主动偏股基金前十大重仓（截至2026-09-30）｜<b>持仓偏低＝组内30只中公募A股配置占比最低的1/3</b>｜点击个股名称查看盈利·估值·ROE详情</div></div>
<div class="tabs"><div class="tab active" data-idx="科创综指">科创综指</div><div class="tab" data-idx="创业板综">创业板综</div></div></div>
<div class="cards">
<div class="card"><div class="v" id="cCount"></div><div class="l">指数成分股总数</div></div>
<div class="card"><div class="v" id="cWeight"></div><div class="l">前30权重合计</div></div>
<div class="card"><div class="v" id="cLow" style="color:var(--amber)"></div><div class="l">持仓偏低（组内后1/3）</div></div>
<div class="card"><div class="v" id="cNoQ3"></div><div class="l">Q3未进公募重仓</div></div>
<div class="card"><div class="v" id="cOut"></div><div class="l">主模型外（公募为补算）</div></div>
</div>
<div class="panel"><div class="panelHead"><div><h2>权重前30个股明细</h2><div class="sub">默认按指数权重降序；公募配置%列的横向条为组内相对水平，越短公募持仓越少</div><span class="count" id="rowCount"></span></div></div>
<div class="toolbar">
<input type="text" id="search" placeholder="搜索代码 / 简称 / 行业">
<label><input type="checkbox" id="onlyLow"> 仅看持仓偏低</label>
<select id="sortBy"><option value="rank">按指数权重↓</option><option value="fundWeightA">按公募配置↑</option><option value="fundCount">按持有基金数↑</option><option value="ytd">按年初以来涨幅↓</option><option value="since924">按924以来涨幅↓</option></select>
<button class="primary" id="exportBtn">导出CSV</button></div>
<div style="overflow:auto;max-height:720px"><table><thead>
<tr class="grp"><th colspan="6">指数权重</th><th colspan="4">2026Q3公募重仓</th><th colspan="2">过往跌幅</th></tr>
<tr><th class="c0">排名</th><th class="c1">代码</th><th class="c2">简称</th><th class="c3">行业</th><th>权重%</th><th>累计权重%</th>
<th>持有基金数</th><th>公募配置%</th><th>流通市值占比%</th><th>超配pct</th>
<th>年初以来%</th><th>924以来%</th></tr></thead><tbody id="tbody"></tbody></table></div></div>
<div class="note"><b>口径与限制</b><br>
① 公募持仓＝2026Q3主动偏股基金<b>前十大重仓</b>口径：持有基金数、公募A股配置占比（Q3重仓合计归一100%）、超配＝配置占比−流通市值占比（由持股总市值/占流通股比反推）；<b>持仓偏低</b>＝该指数权重前30中公募A股配置占比最低的1/3（10只），仅组内相对比较，未进重仓者必属最低一档。<br>
② 公募配置%后的横向条为组内相对水平（前30中最高配置＝满格），直观对比谁持仓少，不构成评分。<br>
③ 过往跌幅为区间涨跌幅（源表"下半年以来"实际自2026/1/1起算，故标年初以来）；仅展示，不参与任何打分。<br>
④ 主模型外的成分股：公募以Q3重仓原始值补算（标注"模型外补算"），无同源基本面数据；创业板综权重合计100.4%为源表舍入。<br>
⑤ 低公募持仓可能反映流动性、解禁、基本面分歧等原因，名单需二次验证；季报披露滞后，Q3之后调仓不可见。</div>
</div>
<div id="mask"></div><div id="drawer"><div class="dHead"><div id="dTitle"></div><button class="dClose" id="dClose">×</button></div><div class="dBody" id="dBody"></div></div>
<script id="pageData" type="application/json">__DATA__</script>
<script id="pageSummary" type="application/json">__SUMMARY__</script>
<script>
const DATA=JSON.parse(document.getElementById('pageData').textContent),SUM=JSON.parse(document.getElementById('pageSummary').textContent);
const $=id=>document.getElementById(id);
let cur='科创综指';
const fmt=(v,d=1)=>v===null||v===undefined||Number.isNaN(v)?'—':Number(v).toFixed(d);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const signed=(v,d=1)=>v===null||v===undefined?'—':(v>0?'+':'')+Number(v).toFixed(d);
const cls=v=>v===null||v===undefined?'':(v>0?'up':v<0?'down':'');
function items(){return DATA[cur]}
function findStock(code){return items().find(x=>x.code===code)}
function filtered(){const q=$('search').value.trim().toLowerCase();return items().filter(x=>{
 if($('onlyLow').checked&&!x.lowHold)return false;
 if(q&&!(x.code.toLowerCase().includes(q)||x.name.toLowerCase().includes(q)||(x.industry||'').toLowerCase().includes(q)))return false;
 return true})}
function renderCards(){const s=SUM[cur];
 $('cCount').textContent=s.total+' 只';$('cWeight').textContent=fmt(s.weightSum)+'%';
 $('cLow').textContent=s.lowHold+' 只';$('cNoQ3').textContent=s.noQ3+' 只';$('cOut').textContent=s.outModel+' 只'}
function renderTable(){const key=$('sortBy').value;const rows=filtered().sort((a,b)=>(a[key]??-1e9)-(b[key]??-1e9));
 $('rowCount').textContent=rows.length+' / 30 只';
 let cum=0;
 $('tbody').innerHTML=rows.map(x=>{cum+=x.weight;return `<tr class="${x.lowHold?'lowRow':''}"><td class="c0">${x.rank}</td><td class="c1">${x.code}</td>
 <td class="c2"><span class="name" onclick="openStock('${x.code}')"><b>${esc(x.name)}</b></span>${x.recentIPO?' <span class="tag amber">次新</span>':''}${x.lowHold?' <span class="tag amber">持仓偏低</span>':''}${x.hasQ3?'':' <span class="tag red">未进重仓</span>'}${x.inModel?'':' <span class="tag gray">模型外补算</span>'}</td>
 <td class="c3">${esc(x.industry)}</td><td><b>${fmt(x.weight,3)}</b></td><td>${fmt(cum,1)}</td>
 <td>${x.hasQ3?fmt(x.fundCount,0):'—'}</td>
 <td><span class="${x.lowHold?'':''}" style="${x.lowHold?'color:var(--amber);font-weight:600':''}">${fmt(x.fundWeightA??0,3)}</span><span class="fwbar ${x.lowHold?'low':''}"><i style="width:${fmt(x.fwBar,0)}%"></i></span></td>
 <td>${fmt(x.marketCapPct)}%</td>
 <td class="${cls(x.overweight)}">${x.overweight===null||x.overweight===undefined?'—':signed(x.overweight,3)}</td>
 <td class="${cls(x.ytd)}">${signed(x.ytd)}</td><td class="${cls(x.since924)}">${signed(x.since924)}</td></tr>`}).join('')||'<tr><td colspan="12" class="empty">无匹配个股</td></tr>'}
function renderAll(){renderCards();renderTable()}
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{document.querySelectorAll('.tab').forEach(z=>z.classList.remove('active'));t.classList.add('active');cur=t.dataset.idx;cum=0;renderAll()});
for(const id of ['onlyLow','sortBy'])$(id).onchange=renderAll;
$('search').oninput=renderTable;
/* ---------- 个股详情抽屉 ---------- */
const YEARS=['2021','2022','2023','2024','2025','26E','27E','28E'];
function barChart(arr){if(!arr||arr.every(v=>v===null||v===undefined))return '<div class="dHint">无数据</div>';
 const vals=arr.map(v=>typeof v==='number'?v:null);const W=470,H=120,P=18;
 const known=vals.filter(v=>v!==null);
 if(!known.length)return '<div class="dHint">无数据</div>';
 const mx=Math.max(...known.map(v=>Math.abs(v)),0.001);
 const bw=(W-2*P)/vals.length;
 const y0=H/2;
 let s=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto">`;
 s+=`<line x1="${P}" y1="${y0}" x2="${W-P}" y2="${y0}" stroke="#e5e9ef"/>`;
 vals.forEach((v,i)=>{if(v===null)return;const h=v/mx*(H/2-P);const x=P+i*bw+bw*0.18;const w=bw*0.64;
  const fill=i===6?'#2f6fed':(v>=0?'#0e9f8a':'#d0453e');
  s+=`<rect x="${x}" y="${v>=0?y0-h:y0}" width="${w}" height="${Math.max(1,Math.abs(h))}" rx="2" fill="${fill}" opacity=".9"><title>${YEARS[i]}: ${fmt(v,3)}</title></rect>`;
  s+=`<text x="${x+w/2}" y="${H-4}" font-size="9.5" fill="#6b7684" text-anchor="middle">${YEARS[i]}</text>`;
  s+=`<text x="${x+w/2}" y="${v>=0?y0-h-3:y0+Math.abs(h)+10}" font-size="9" fill="#6b7684" text-anchor="middle">${fmt(v,i>=6?2:1)}</text>`});
 return s+'</svg>'}
function revText(v){if(v===null||v===undefined)return '—';if(typeof v==='string')return v;return signed(v)+'%'}
function openStock(code){const x=findStock(code);if(!x)return;
 $('dTitle').innerHTML=`<h3>${esc(x.name)}</h3><div class="sub">${x.code}｜${esc(x.industry)}${x.marketCap?'｜市值 '+fmt(x.marketCap,0)+'亿':''}${x.recentIPO?' <span class="tag amber">次新</span>':''}${x.lowHold?' <span class="tag amber">持仓偏低</span>':''}${x.inModel?'':' <span class="tag gray">模型外补算</span>'}</div>`;
 let h='';
 h+=`<div class="dCard"><h4>指数权重与Q3公募持仓</h4><div class="kv">
 <div class="k">指数权重</div><div class="v">${fmt(x.weight,3)}%</div><div class="k">组内权重排名</div><div class="v">第 ${x.rank} 位</div><div class="k">持仓相对水平</div><div class="v ${x.lowHold?'':''}" style="${x.lowHold?'color:var(--amber)':''}">${fmt(x.fwBar,0)}/100</div>
 <div class="k">持有基金数</div><div class="v">${x.hasQ3?fmt(x.fundCount,0):'未进重仓'}</div><div class="k">公募A股配置</div><div class="v">${fmt(x.fundWeightA??0,3)}%</div><div class="k">流通市值占比</div><div class="v">${fmt(x.marketCapPct)}%</div>
 <div class="k">超配(pct)</div><div class="v ${cls(x.overweight)}">${x.overweight==null?'—':signed(x.overweight,3)}</div><div class="k"></div><div class="v"></div><div class="k"></div><div class="v"></div></div></div>`;
 if(x.inModel){
  h+=`<div class="dCard"><h4>盈利（27E EPS，元）</h4>
  <div class="kv" style="grid-template-columns:1fr 1fr 1fr 1fr">
  <div class="k">27E EPS</div><div class="v">${fmt(x.eps27,3)}</div><div class="k">27E同比</div><div class="v ${cls(x.epsGrowth)}">${signed(x.epsGrowth)}%</div>
  <div class="k">最新周上修</div><div class="v">${revText(x.revs?x.revs[0]:null)}</div><div class="k">连续上修</div><div class="v">${x.epsStreak?x.epsStreak+' 周':'—'}</div></div>
  <div class="dHint">近三周上修（近→远）：${(x.revs||[]).map(revText).join(' / ')}</div>${barChart(x.eps)}</div>`;
  h+=`<div class="dCard"><h4>估值（${x.method==='PB'?'PB法':'PE法'}）</h4>
  <div class="kv" style="grid-template-columns:1fr 1fr 1fr 1fr">
  ${x.method==='PB'?`<div class="k">27E PB</div><div class="v">${fmt(x.pb27,2)}</div>`:`<div class="k">27E PE</div><div class="v">${fmt(x.pe27)}</div>`}
  <div class="k">26E PE</div><div class="v">${fmt(x.pe26)}</div><div class="k">28E PE</div><div class="v">${fmt(x.pe28)}</div>
  <div class="k">${x.method}历史分位</div><div class="v">${fmt(x.valPct)}</div><div class="k">利润消化</div><div class="v">${x.digestLabel||'—'}（${fmt(x.digestScore,0)}分）</div></div></div>`;
  h+=`<div class="dCard"><h4>ROE（%）</h4>
  <div class="kv" style="grid-template-columns:1fr 1fr 1fr 1fr">
  <div class="k">27E ROE</div><div class="v">${fmt(x.roe27)}</div><div class="k">27E−26E</div><div class="v ${cls(x.roeDelta)}">${signed(x.roeDelta)}</div>
  <div class="k">较上周修正</div><div class="v">${signed(x.roeRev)}pct</div><div class="k">周期位置</div><div class="v">${esc(x.cycle||'—')}</div></div>${barChart(x.roe)}</div>`;
  h+=`<div class="dCard"><h4>过往涨跌幅（%）</h4><div class="kv" style="grid-template-columns:1fr 1fr 1fr">
  <div class="k">近1月</div><div class="v ${cls(x.month)}">${signed(x.month)}</div><div class="k">年初以来</div><div class="v ${cls(x.ytd)}">${signed(x.ytd)}</div><div class="k">924以来</div><div class="v ${cls(x.since924)}">${signed(x.since924)}</div></div></div>`;
  h+=`<div class="note" style="border:none;padding:0">基本面数据与主评分模型（2026-11-02截面）同口径，仅展示、不参与本页任何打分。</div>`;
 }else{
  h+=`<div class="dCard"><h4>基本面</h4><div class="dHint">该股不在主评分模型样本内（无同源盈利/估值/ROE数据），公募持仓为Q3重仓原始值补算。</div></div>`;
 }
 $('dBody').innerHTML=h;
 $('mask').style.display='block';$('drawer').classList.add('open')}
window.openStock=openStock;
function closeDrawer(){$('mask').style.display='none';$('drawer').classList.remove('open')}
$('dClose').onclick=closeDrawer;$('mask').onclick=closeDrawer;
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDrawer()});
/* ---------- 导出 ---------- */
$('exportBtn').onclick=()=>{const key=$('sortBy').value;const rows=filtered().sort((a,b)=>(a[key]??-1e9)-(b[key]??-1e9));
 const head=['组内权重排名','代码','简称','指数','行业','权重%','持有基金数','公募A股配置占比%','流通市值占比%','超配pct','持仓偏低','27E EPS同比%','近1周上修','前1周上修','前2周上修','连续上修周数','27E ROE%','ROE周期','年初以来%','924以来%','Q3未进重仓','模型外补算'];
 const revS=v=>v===null||v===undefined?'':(typeof v==='string'?v:fmt(v,3));
 const lines=[head,...rows.map(x=>[x.rank,x.code,x.name,cur,x.industry,fmt(x.weight,4),fmt(x.fundCount,0),fmt(x.fundWeightA??0,4),fmt(x.marketCapPct),x.overweight??'',x.lowHold?'是':'否',x.epsGrowth??'',...(x.revs||[null,null,null]).map(revS),x.epsStreak??'',x.roe27??'',x.cycle||'',x.ytd??'',x.since924??'',x.hasQ3?'否':'是',x.inModel?'否':'是'])];
 const csv='﻿'+lines.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
 const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=`双创权重前30公募持仓_${cur}.csv`;a.click()};
renderAll();
</script></body></html>"""

out = TEMPLATE.replace("__DATA__", DATA_JSON).replace("__SUMMARY__", SUMMARY_JSON)
io.open(OUTPUT, "w", encoding="utf-8").write(out)
print("saved:", OUTPUT)
