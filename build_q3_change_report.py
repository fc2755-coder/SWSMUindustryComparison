from pathlib import Path
import json
import html

ROOT = Path(r"C:/Users/Owner/WorkBuddy/2026-10-26-13-16-06")
OUT = Path(r"C:/Users/Owner/Desktop/公募Q3更新_行业配置与个股评分变化.html")
d = json.loads((ROOT / "q3_crowding_comparison.json").read_text(encoding="utf-8"))
before, after = d["before"], d["after"]
bs = {x["code"]: x for x in before["stocks"]}
bi = {(x["level"], x["name"]): x for x in before["industries"]}

def f(v, n=2):
    return "—" if v is None else f"{v:.{n}f}"

def signed(v, n=2):
    return "—" if v is None else f"{v:+.{n}f}"

def esc(s):
    return html.escape(str(s))

def delta(v, n=2):
    cls = "up" if v > 0 else "down" if v < 0 else "flat"
    return f'<span class="{cls}">{signed(v,n)}</span>'

industries = []
for x in after["industries"]:
    if x["level"] != "L1":
        continue
    y = bi[(x["level"], x["name"])]
    industries.append({**x, "oldScore": y["totalScore"], "oldCrowd": y["crowdRaw"], "scoreDelta": round(x["totalScore"]-y["totalScore"],2), "crowdDelta": round(x["crowdRaw"]-y["crowdRaw"],2), "rankDelta": y["rank"]-x["rank"]})

stocks = []
for x in after["stocks"]:
    if x["recentIPO"]:
        continue
    y = bs[x["code"]]
    stocks.append({**x, "oldScore": y["totalScore"], "oldCrowd": y["crowdRaw"], "scoreDelta": round(x["totalScore"]-y["totalScore"],2), "crowdDelta": round(x["crowdRaw"]-y["crowdRaw"],2), "oldHeld": y["hasFundHolding"]})

persistent_up = sorted([x for x in stocks if x["hasFundHolding"] and x["oldHeld"]], key=lambda z:z["scoreDelta"], reverse=True)[:18]
scope_up = sorted([x for x in stocks if not x["hasFundHolding"]], key=lambda z:z["scoreDelta"], reverse=True)[:15]
down = sorted(stocks, key=lambda z:z["scoreDelta"])[:15]
adds = sorted([x for x in stocks if x["totalScore"] >= 60 and x["overweightChange"] is not None], key=lambda z:z["overweightChange"], reverse=True)[:15]
ind_up = sorted(industries, key=lambda z:z["scoreDelta"], reverse=True)
ind_down = sorted(industries, key=lambda z:z["scoreDelta"])
ow_up = sorted(industries, key=lambda z:z["overweightChange"], reverse=True)
ow_down = sorted(industries, key=lambda z:z["overweightChange"])

def industry_rows(rows, all_cols=True):
    out=[]
    for x in rows:
        out.append(f'<tr><td><b>{esc(x["name"])}</b></td><td>{f(x["oldScore"])}</td><td><b>{f(x["totalScore"])}</b></td><td>{delta(x["scoreDelta"])}</td><td>{f(x["oldCrowd"])} → {f(x["crowdRaw"])}</td><td>{f(x["overweight"])}pct</td><td>{delta(x["overweightChange"])}</td></tr>')
    return "".join(out)

def stock_rows(rows, mode="score"):
    out=[]
    for x in rows:
        q3ow = "未进Q3重仓" if not x["hasFundHolding"] else f'{f(x["overweight"],3)}pct'
        extra = delta(x["overweightChange"],3) if mode=="add" else f'{f(x["oldCrowd"])} → {f(x["crowdRaw"])}'
        out.append(f'<tr><td><b>{esc(x["name"])}</b><small>{x["code"]}</small></td><td>{esc(x["industry"])}</td><td>{f(x["oldScore"])}</td><td><b>{f(x["totalScore"])}</b></td><td>{delta(x["scoreDelta"])}</td><td>{extra}</td><td>{q3ow}</td><td>{f(x["valuationScore"])}</td></tr>')
    return "".join(out)

def mini_bar(value, maximum, positive=True):
    w = max(3, abs(value)/maximum*100)
    cls = "barUp" if positive else "barDown"
    return f'<div class="mini"><span>{signed(value)}</span><i class="{cls}" style="width:{w:.1f}%"></i></div>'

up_cards = "".join(f'<div class="rankItem"><b>{esc(x["name"])}</b>{mini_bar(x["scoreDelta"],6.5,True)}</div>' for x in ind_up[:5])
down_cards = "".join(f'<div class="rankItem"><b>{esc(x["name"])}</b>{mini_bar(x["scoreDelta"],6.5,False)}</div>' for x in ind_down[:5])

html_text = f'''<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>公募Q3更新｜行业配置与个股评分变化</title><style>
:root{{--ink:#172b3a;--muted:#6c7d89;--line:#e5eaee;--paper:#f5f7f8;--teal:#087d75;--red:#c84b4b;--green:#16866b;--amber:#ad7616;--blue:#386b9e}}*{{box-sizing:border-box}}body{{margin:0;background:var(--paper);color:var(--ink);font:14px/1.65 "Microsoft YaHei",Arial,sans-serif}}.wrap{{max-width:1440px;margin:auto;padding:32px 38px 50px}}header{{padding:30px 34px;background:linear-gradient(120deg,#123c46,#176d6a);color:white;border-radius:18px;box-shadow:0 12px 32px #173d4718}}h1{{font-size:28px;margin:0 0 8px}}header p{{margin:4px 0;opacity:.84}}.eyebrow{{font-size:12px;letter-spacing:2px;text-transform:uppercase;opacity:.7}}.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:18px 0}}.card,.panel{{background:#fff;border:1px solid var(--line);border-radius:14px;box-shadow:0 3px 14px #18313b06}}.card{{padding:18px 20px}}.card b{{display:block;font-size:24px;line-height:1.2;margin:4px 0}}.card span,.sub{{color:var(--muted);font-size:12px}}.panel{{padding:22px 24px;margin-top:18px}}h2{{font-size:19px;margin:0 0 6px}}h3{{font-size:15px;margin:20px 0 8px}}.lead{{font-size:16px;line-height:1.9;margin:12px 0}}.conclusions{{display:grid;grid-template-columns:1.1fr 1fr 1fr;gap:14px}}.callout{{padding:16px 18px;border-radius:12px;background:#f2f8f7;border-left:4px solid var(--teal)}}.callout.warn{{background:#fff8ed;border-color:var(--amber)}}.callout b{{display:block;margin-bottom:5px}}.twocol{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}.rankItem{{margin:12px 0}}.rankItem>b{{display:block;margin-bottom:4px}}.mini{{display:flex;align-items:center;gap:10px}}.mini span{{width:54px;font-variant-numeric:tabular-nums}}.mini i{{display:block;height:8px;border-radius:8px;max-width:calc(100% - 64px)}}.barUp{{background:var(--red)}}.barDown{{background:var(--green)}}.up{{color:var(--red);font-weight:700}}.down{{color:var(--green);font-weight:700}}.flat{{color:var(--muted)}}.tableWrap{{overflow:auto;border:1px solid var(--line);border-radius:10px;margin-top:12px}}table{{width:100%;border-collapse:collapse;white-space:nowrap;font-size:13px}}th{{position:sticky;top:0;background:#f2f5f6;color:#52636e;text-align:right;padding:10px 12px;font-weight:600}}td{{padding:10px 12px;border-top:1px solid #edf0f2;text-align:right;font-variant-numeric:tabular-nums}}th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){{text-align:left}}tr:hover td{{background:#f8fbfb}}td small{{display:block;color:var(--muted);font-size:10px}}details{{margin-top:14px}}summary{{cursor:pointer;color:var(--blue);font-weight:600}}.note{{background:#fff8e9;border:1px solid #f0dfb9;padding:13px 16px;border-radius:10px;color:#74551b;margin:12px 0}}.tag{{display:inline-block;padding:2px 8px;background:#eaf4f2;color:var(--teal);border-radius:10px;font-size:11px}}footer{{color:var(--muted);font-size:12px;margin-top:22px}}@media(max-width:900px){{.wrap{{padding:16px}}.cards,.conclusions,.twocol{{grid-template-columns:1fr 1fr}}header{{padding:22px}}h1{{font-size:23px}}}}@media(max-width:600px){{.cards,.conclusions,.twocol{{grid-template-columns:1fr}}}}
</style></head><body><div class="wrap"><header><div class="eyebrow">PUBLIC FUND HOLDINGS · Q3 UPDATE</div><h1>公募Q3更新：行业配置与个股评分变化</h1><p>固定价格、盈利、ROE与估值，仅比较公募数据更新带来的评分变化</p><p>市场与盈利预测截面：2026-10-25｜旧基线：2026H1全持仓｜新评分：2026Q3前十大重仓｜同口径环比：2026Q2→Q3</p></header>
<section class="cards"><div class="card"><span>综合分上升最多</span><b>{ind_up[0]["name"]} {signed(ind_up[0]["scoreDelta"])}</b><span>拥挤度 {f(ind_up[0]["oldCrowd"])} → {f(ind_up[0]["crowdRaw"])}</span></div><div class="card"><span>Q2→Q3超配提升最多</span><b>{ow_up[0]["name"]} {signed(ow_up[0]["overweightChange"])}pct</b><span>增配与综合分上升不是一回事</span></div><div class="card"><span>综合分下降最多</span><b>{ind_down[0]["name"]} {signed(ind_down[0]["scoreDelta"])}</b><span>拥挤度 {f(ind_down[0]["oldCrowd"])} → {f(ind_down[0]["crowdRaw"])}</span></div><div class="card"><span>低估值高分股</span><b>134 → 263只</b><span>+129只；受重仓披露范围收窄影响较大</span></div></section>
<section class="panel"><h2>先看结论：把“升分”与“增配”分开</h2><div class="conclusions"><div class="callout"><b>1｜升分主线：去拥挤带来的性价比改善</b>建筑材料、环保、钢铁综合分分别上升 <span class="up">+4.53、+3.13、+2.93</span>。但建材Q2→Q3超配实际下降0.70pct，因此其升分不能解释为Q3资金加仓。</div><div class="callout warn"><b>2｜增配主线：医药最突出，但拥挤惩罚加大</b>医药生物超配环比提升3.23pct，居一级行业首位；综合分却下降1.33分。资金增配是事实，但在本模型中会增加拥挤度、压低安全分。</div><div class="callout"><b>3｜个股更值得看“持续重仓且升分”</b>法拉电子、固德威、中钨高新、伟测科技、扬杰科技等，在H1和Q3均有持仓记录，升分更具可解释性。未进入Q3前十大的股票，不能直接认定被清仓。</div></div></section>
<section class="panel"><h2>一、行业综合分：谁显著上升，谁明显下降</h2><p class="sub">综合分变化＝Q3重仓口径评分－H1全持仓基线评分；其他四维保持不变。数据源：用户提供公募工作簿，2026H1／2026Q3。</p><div class="twocol"><div><h3>综合分上升前五</h3>{up_cards}</div><div><h3>综合分下降前五</h3>{down_cards}</div></div><div class="tableWrap"><table><thead><tr><th>行业</th><th>更新前</th><th>Q3评分</th><th>评分变化</th><th>拥挤度 前→后</th><th>Q3超配</th><th>超配Q2→Q3</th></tr></thead><tbody>{industry_rows(ind_up[:8]+ind_down[:8])}</tbody></table></div><details><summary>展开全部31个一级行业</summary><div class="tableWrap"><table><thead><tr><th>行业</th><th>更新前</th><th>Q3评分</th><th>评分变化</th><th>拥挤度 前→后</th><th>Q3超配</th><th>超配Q2→Q3</th></tr></thead><tbody>{industry_rows(sorted(industries,key=lambda z:z["scoreDelta"],reverse=True))}</tbody></table></div></details></section>
<section class="panel"><h2>二、行业配置：Q2→Q3同口径看到了什么</h2><p class="lead"><b>增配端：</b>医药生物 <span class="up">+3.23pct</span>、汽车 <span class="up">+0.75pct</span>、国防军工 <span class="up">+0.74pct</span>、有色金属 <span class="up">+0.70pct</span>。<b>减配端：</b>电子 <span class="down">−4.53pct</span>、银行 <span class="down">−1.27pct</span>、通信 <span class="down">−1.06pct</span>。</p><p>电子和通信虽然超配回落，Q3拥挤度仍分别为100.0和96.7，仍属于高拥挤方向；这更像“高位降温”，而不是已经变成低配。家电、军工、食品饮料综合分下降较多，主要源于新口径下拥挤度上行，并非盈利或估值突然恶化。</p></section>
<section class="panel"><h2>三、个股选择：持续有持仓记录，且分数显著上升</h2><p class="sub">按综合分升幅排序；剔除上市未满一年标的。该组H1与Q3均有持仓记录，优先于“未进Q3重仓导致升分”的股票。</p><div class="tableWrap"><table><thead><tr><th>股票</th><th>行业</th><th>更新前</th><th>Q3评分</th><th>评分变化</th><th>拥挤度 前→后</th><th>Q3超配</th><th>估值得分</th></tr></thead><tbody>{stock_rows(persistent_up)}</tbody></table></div><p class="lead">若同时要求当前综合分较高，优先复核 <b>伟测科技（688372.SH，76.39分）</b>、<b>法拉电子（600563.SH，69.97分）</b>、<b>扬杰科技（300373.SZ，69.68分）</b>、<b>固德威（688390.SH，68.97分）</b>和<b>中钨高新（000657.SZ，66.67分）</b>。它们共同特征是拥挤度显著下降，而不是盈利、估值维度发生变化。</p></section>
<section class="panel"><h2>四、公募同口径增配、且当前总分不低于60</h2><p class="sub">按Q2→Q3个股超配变化排序。增配不一定带来评分上升，需同时观察当前拥挤度。</p><div class="tableWrap"><table><thead><tr><th>股票</th><th>行业</th><th>更新前</th><th>Q3评分</th><th>评分变化</th><th>超配Q2→Q3(pct)</th><th>Q3超配</th><th>估值得分</th></tr></thead><tbody>{stock_rows(adds,"add")}</tbody></table></div><div class="note">工业富联、海康威视、中国船舶、比亚迪等出现同口径增配且当前分数仍在60以上；但睿创微纳、生益电子、新宙邦等拥挤度已偏高，不宜把“增配”单独理解为性价比提升。</div></section>
<section class="panel"><h2>五、口径敏感型升分：可以看，但不要误读为公募减持</h2><p class="sub">下列股票更新前分数较高拥挤，Q3未进入前十大重仓名单，因此拥挤度和综合分大幅变化。</p><div class="tableWrap"><table><thead><tr><th>股票</th><th>行业</th><th>更新前</th><th>Q3评分</th><th>评分变化</th><th>拥挤度 前→后</th><th>Q3状态</th><th>估值得分</th></tr></thead><tbody>{stock_rows(scope_up)}</tbody></table></div><div class="note"><b>关键限制：</b>季度报告只披露前十大重仓，半年报披露全部持仓。未进Q3前十大，可能仍被基金持有，只是不再位列前十。因此该组只能标记为“拥挤度口径敏感”，不能写成“公募清仓”或“确定减持”。</div></section>
<section class="panel"><h2>六、分数下降较多的个股</h2><p class="sub">按综合分降幅排序。分数下降主要由拥挤安全分下降造成，其他四维未变。</p><div class="tableWrap"><table><thead><tr><th>股票</th><th>行业</th><th>更新前</th><th>Q3评分</th><th>评分变化</th><th>拥挤度 前→后</th><th>Q3超配</th><th>估值得分</th></tr></thead><tbody>{stock_rows(down)}</tbody></table></div></section>
<section class="panel"><h2>使用建议与口径边界</h2><p>① 行业层面先看Q2→Q3超配方向，再看当前拥挤度，最后结合综合分；不要只按升分排名理解调仓。② 个股优先研究“持续有持仓且升分”组，再核查盈利预测与估值；对“未进Q3重仓”组保持谨慎。③ 本报告衡量的是模型评分变化，不是收益预测，也不直接构成买卖名单。</p><p class="sub">来源：用户提供的《主动偏股基金全持仓数据_26H1.xlsx》《重仓持股26Q2.xlsx》《重仓持股26Q3.xlsx》《主动偏股基金数据库_2026Q3.xlsx》。行业配置与超配采用数据库汇总；个股采用重仓明细。所有评分公式与权重保持不变。</p></section>
<footer>独立分析报告｜不修改原《行业比较与个股筛选模型》页面｜生成日期：2026-10-27</footer></div></body></html>'''
OUT.write_text(html_text, encoding="utf-8")
print(OUT)
print("INDUSTRY_UP", [(x["name"],x["scoreDelta"]) for x in ind_up[:3]])
print("STOCK_UP", [(x["name"],x["scoreDelta"]) for x in persistent_up[:5]])
