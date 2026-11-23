# -*- coding: utf-8 -*-
"""行业因子历史有效性回测
数据: panels.pkl (SWSMU周度收益/成交拥挤 + 兴证季度持仓 + SWSMU年度估值盈利)
统计: RankIC均值/t值/ICIR/IC胜率 + 五分位单调性 + 多空t + 分段稳健性
口径: 与主模型一致——估值仅银行/非银金融/房地产用PB，其余PE
"""
import pickle, json
import numpy as np
import pandas as pd

def _spearman(a, b):
    ra = pd.Series(a).rank()
    rb = pd.Series(b).rank()
    return np.corrcoef(ra, rb)[0, 1]

class _stats:
    spearmanr = staticmethod(lambda a, b: (_spearman(a, b), np.nan))

ST = _stats

P = pickle.load(open('panels.pkl', 'rb'))
ret_w = P['ret_w'] / 100.0          # 周收益小数
to_d = P['turnover_pct_d']
alloc = P['alloc_q']                 # 季度配置 %
pe_a, pb_a = P['pe_a'], P['pb_a']
roe_a, npg_a, dy_a = P['roe_a'], P['npg_a'], P['dy_a']
MODEL_PB = {'银行', '非银金融', '房地产'}

RESULTS = {}

def rank_ic(F, R):
    out = {}
    for t in F.index:
        f, r = F.loc[t], R.loc[t]
        m = f.notna() & r.notna()
        if m.sum() >= 10:
            out[t] = ST.spearmanr(f[m], r[m])[0]
    return pd.Series(out)

def ic_stats(ic):
    ic = ic.dropna()
    if len(ic) < 5 or ic.std() == 0:
        return dict(n=len(ic), ic=np.nan, t=np.nan, icir=np.nan, pos=np.nan)
    return dict(n=len(ic), ic=round(ic.mean(), 4), t=round(ic.mean()/ic.std()*np.sqrt(len(ic)), 2),
                icir=round(ic.mean()/ic.std(), 3), pos=round((ic>0).mean(), 3))

def quintile(F, R, nq=5):
    qm = {q: [] for q in range(1, nq+1)}; ls = []
    for t in F.index:
        f, r = F.loc[t], R.loc[t]
        m = f.notna() & r.notna()
        if m.sum() < 10: continue
        lab = pd.qcut(f[m].rank(method='first'), nq, labels=False) + 1
        rr = r[m]
        for q in range(1, nq+1):
            qm[q].append(rr[lab == q].mean())
        ls.append(rr[lab == nq].mean() - rr[lab == 1].mean())
    means = [np.mean(qm[q]) for q in range(1, nq+1)]
    mono = ST.spearmanr(range(1, nq+1), means)[0]
    ls = pd.Series(ls)
    return dict(q_mean=[round(x*100, 3) for x in means], mono=round(mono, 3),
                ls_mean=round(ls.mean()*100, 4), ls_t=round(ls.mean()/ls.std()*np.sqrt(len(ls)), 2), n=len(ls))

def fwd_ret(ret, h, start_offset=0):
    """非重叠前瞻收益: 每h周一个截面"""
    gross = (1 + ret).rolling(h).apply(np.prod, raw=True).shift(-h)
    idx = ret.index[start_offset::h]
    return gross.loc[gross.index.intersection(idx)]

def past_ret(ret, h):
    return (1 + ret).rolling(h).apply(np.prod, raw=True) - 1

# ============ A. 周度价格因子 (2021-01 ~ 2026-10, 297周) ============
print('=' * 70)
print('A. 价格/动量反转因子（周度面板，31行业）')
for h_past, tag in [(1, 'mom1w'), (4, 'mom4w'), (13, 'mom13w'), (52, 'mom52w')]:
    past = past_ret(ret_w, h_past)
    for h_fwd in (1, 4, 13):
        if h_fwd > h_past and h_past < 4:  # 短因子只测短前瞻，减少组合数
            continue
        fwd = fwd_ret(ret_w, h_fwd, start_offset=h_fwd - 1)
        F = past.reindex(fwd.index).dropna(how='all')
        if F.empty: continue
        ic = ic_stats(rank_ic(F, fwd.reindex(F.index)))
        q = quintile(F, fwd.reindex(F.index))
        RESULTS[f'{tag}->fwd{h_fwd}w'] = dict(ic=ic, q=q)
        print(f"{tag:>7} → fwd{h_fwd:>2}w | IC={ic['ic']:+.3f} t={ic['t']:+.2f} 胜率={ic['pos']:.0%} | Q5-Q1={q['ls_mean']:+.3f}%/期 t={q['ls_t']:+.2f} 单调={q['mono']:+.2f} (n={ic['n']})")

# 模型价格分口径: 超跌高分（-past4w, 上涨clip为中性0）
print('\n模型"价格回撤分"口径: factor=-past4w（超跌高分）')
past4 = past_ret(ret_w, 4)
f_rev = (-past4).clip(upper=0)  # 上涨视为中性0，仅负收益计跌幅
for h_fwd in (4, 13):
    fwd = fwd_ret(ret_w, h_fwd, start_offset=h_fwd - 1)
    F = f_rev.reindex(fwd.index).dropna(how='all')
    ic = ic_stats(rank_ic(F, fwd.reindex(F.index)))
    q = quintile(F, fwd.reindex(F.index))
    RESULTS[f'rev4w_clipped->fwd{h_fwd}w'] = dict(ic=ic, q=q)
    print(f"  rev4(clip) → fwd{h_fwd}w | IC={ic['ic']:+.3f} t={ic['t']:+.2f} | Q5-Q1={q['ls_mean']:+.3f}%/期 t={q['ls_t']:+.2f} 单调={q['mono']:+.2f}")

# 分段稳健性（动量4w→fwd4w）
fwd4 = fwd_ret(ret_w, 4, start_offset=3)
F4 = past4.reindex(fwd4.index)
for seg, lo, hi in [('2021-2023', '2021', '2024'), ('2024-2026', '2024', '2027')]:
    mask = (F4.index >= lo) & (F4.index < hi)
    ic = ic_stats(rank_ic(F4[mask], fwd4[mask]))
    RESULTS[f'mom4w->fwd4w_{seg}'] = ic
    print(f"  mom4w→fwd4w {seg}: IC={ic['ic']:+.3f} t={ic['t']:+.2f} (n={ic['n']})")

# ============ B. 交易拥挤因子（日频成交额占比分位, 2024-10~2026-09） ============
print('\n' + '=' * 70)
print('B. 交易拥挤因子（成交额占比分位数, 周五采样）')
to_w = to_d.resample('W-FRI').last().dropna(how='all')
to_w = to_w.reindex(ret_w.index.intersection(to_w.index)).dropna(how='all')
for h_fwd in (1, 4):
    fwd = fwd_ret(ret_w.reindex(to_w.index), h_fwd, start_offset=h_fwd - 1)
    F = to_w.reindex(fwd.index).dropna(how='all')
    ic = ic_stats(rank_ic(F, fwd.reindex(F.index)))
    q = quintile(F, fwd.reindex(F.index))
    RESULTS[f'turnoverpct->fwd{h_fwd}w'] = dict(ic=ic, q=q)
    print(f"  成交分位 → fwd{h_fwd}w | IC={ic['ic']:+.3f} t={ic['t']:+.2f} 胜率={ic['pos']:.0%} | Q5-Q1={q['ls_mean']:+.3f}%/期 t={q['ls_t']:+.2f} 单调={q['mono']:+.2f} (n={ic['n']})")

# ============ C. 季度持仓拥挤因子（兴证配置比例, 2004Q1~2026Q3） ============
print('\n' + '=' * 70)
print('C. 公募持仓拥挤因子（季度配置比例）——收益窗口受周度面板限制: 2021Q1起')
alloc_lv = alloc.copy()
alloc_chg = alloc.diff()
# 2010年以来历史分位（每季度值在2010Q1~当季序列中的分位）——新口径因子2
alloc_pct = pd.DataFrame(index=alloc.index, columns=alloc.columns, dtype=float)
q_idx = list(alloc.index)
for i, q in enumerate(q_idx):
    d = pd.Timestamp(q)
    if d < pd.Timestamp('2010-03-31'):
        continue
    hist = alloc.loc[(alloc.index >= '2010-03-31') & (alloc.index <= d)]
    alloc_pct.loc[q] = (hist.rank(pct=True) * 100).loc[q]

for fname, F_panel in [('配置水平', alloc_lv), ('配置环比', alloc_chg), ('配置2010以来分位', alloc_pct)]:
    for lag_weeks, ver in [(0, '即时(上限)'), (4, '披露后(保守)')]:
        ics = []
        for qt in q_idx:
            f = F_panel.loc[qt]
            start = pd.Timestamp(qt) + pd.Timedelta(weeks=lag_weeks)
            end = pd.Timestamp(qt) + pd.Timedelta(weeks=lag_weeks + 13)
            mask = (ret_w.index > start) & (ret_w.index <= end)
            if mask.sum() < 10:
                continue
            fwd = (1 + ret_w[mask]).prod() - 1
            m = f.notna() & fwd.notna()
            if m.sum() >= 10:
                ics.append(ST.spearmanr(f[m], fwd[m])[0])
        ic = pd.Series(ics)
        st = ic_stats(ic)
        RESULTS[f'alloc_{fname}_{ver}'] = st
        print(f"  {fname:<10} {ver:<8} → fwd13w | IC={st['ic']:+.3f} t={st['t']:+.2f} 胜率={st['pos']:.0%} (n={st['n']})")

# ============ D. 年度基本面因子（2007~2025, 近似收益+真实收益交叉验证） ============
print('\n' + '=' * 70)
print('D. 年度估值/盈利/红利因子（年末→次年收益；收益=估值变动×盈利复利近似）')
cols_y = [c for c in pe_a.columns if c < '2026']  # 只取2006-2025年度列（剔除2026Q1/H1避免列重名）
years = [c[:4] for c in cols_y]
pe = pe_a[cols_y].copy(); pe.columns = years
pb = pb_a[cols_y].copy(); pb.columns = years
roe = roe_a[cols_y].copy(); roe.columns = years
npg = npg_a[cols_y].copy(); npg.columns = years
dy = dy_a[cols_y].copy(); dy.columns = years

# 近似年度收益
ret_a = pd.DataFrame(index=pe.index, columns=years[1:], dtype=float)
for y0, y1 in zip(years[:-1], years[1:]):
    val_ratio = np.where(pe[y1].notna() & (pe[y0] != 0), pe[y1] / pe[y0], np.nan)
    pb_ratio = np.where(pb[y1].notna() & (pb[y0] != 0), pb[y1] / pb[y0], np.nan)
    pe_ret = val_ratio * (1 + npg[y1] / 100) - 1
    pb_ret = pb_ratio * (1 + roe[y1] / 100) - 1
    for ind in pe.index:
        ret_a.loc[ind, y1] = pb_ret[ind] if ind in MODEL_PB else pe_ret[ind]
ret_a = ret_a.astype(float)

# 近似收益已校验失败（盈利转负年份公式崩坏）→ 年度因子改用真实收益（周度面板复利），样本2021-2025年末
print('  近似收益校验（已判定不可信，年度因子改用真实收益）:')
for y in ['2022', '2023', '2024', '2025']:
    mask = (ret_w.index >= f'{y}-01-01') & (ret_w.index < f'{int(y)+1}-01-01')
    if mask.sum() < 40: continue
    real = ((1 + ret_w[mask]).prod() - 1).mean() * 100
    approx = ret_a[y].astype(float).mean() * 100
    corr = ST.spearmanr(ret_a[y].astype(float), (1 + ret_w[mask]).prod() - 1)[0]
    print(f"    {y}: 近似均值{approx:+.1f}% vs 真实均值{real:+.1f}% | 截面Spearman={corr:+.2f} → {'不可信' if abs(corr)<0.3 else '尚可'}")
RESULTS['annual_return_approx_verdict'] = '近似收益截面相关≈0，年度因子全部改用真实收益'

# 真实年度收益: 年末t → 次年（2026年至10-23）
ret_real = {}
for y in ['2021', '2022', '2023', '2024', '2025']:
    y0, y1 = f'{int(y)+1}-01-01', f'{int(y)+1}-12-31'
    mask = (ret_w.index >= y0) & (ret_w.index <= y1)
    if mask.sum() < 30: continue
    ret_real[y] = (1 + ret_w[mask]).prod() - 1
ret_real = pd.DataFrame(ret_real)
print('  真实年度收益样本:', list(ret_real.columns))

def annual_test_real(fname, Fmat, negate=False):
    ics = []
    for y in ret_real.columns:
        if y not in Fmat.columns: continue
        f = Fmat[y].astype(float) * (-1 if negate else 1)
        r = ret_real[y]
        m = f.notna() & r.notna()
        if m.sum() >= 10:
            ics.append(ST.spearmanr(f[m], r[m])[0])
    ic = pd.Series(ics)
    st = ic_stats(ic)
    RESULTS[fname] = dict(st, yearly_ic=[round(v, 3) for v in ic.tolist()])
    print(f"  {fname:<18} → 次年收益 | IC均值={st['ic']:+.3f} 胜率={st['pos']:.0%} (n={st['n']}, 年度频率样本小仅看方向) 逐年IC={[round(v,2) for v in ic.tolist()]}")
    return ic

# 估值: 10年滚动分位（年末），负PE置NaN；模型方向=便宜高分→对收益预期正
val_pct = pd.DataFrame(index=pe.index, columns=years, dtype=float)
for ind in pe.index:
    for i, y in enumerate(years):
        src = pb if ind in MODEL_PB else pe
        v = src.loc[ind, y]
        hist = src.loc[ind].iloc[max(0, i - 10):i + 1].astype(float)
        hist = hist[hist > 0]
        if np.isfinite(v) and v > 0 and len(hist.dropna()) >= 5:
            val_pct.loc[ind, y] = (hist.dropna() < v).mean() * 100
annual_test_real('估值分位(高=贵)', val_pct, negate=True)
annual_test_real('ROE水平', roe)
annual_test_real('ROE变化', roe.diff(axis=1))
annual_test_real('净利增速', npg)
annual_test_real('股息率', dy)
combo = pd.DataFrame(index=pe.index, columns=years, dtype=float)
for y in years:
    v = val_pct[y].astype(float); r_ = roe[y].astype(float)
    combo[y] = (100 - v) * 0.75 + r_.rank(pct=True) * 100 * 0.25
annual_test_real('便宜+高ROE(3:1)', combo)
# 近似收益版本（仅供参考，已被判不可信）保留在注释层面不再输出

# 分段（真实收益样本仅5年，改为逐年IC已在上文列出，不再做半样本拆分）

# ============ E. 轮动强度regime条件（已知结论复验） ============
print('\n' + '=' * 70)
print('E. 动量因子的市场状态条件检验（上周市场涨/跌分组）')
mkt = ret_w.mean(axis=1)
down_wk = mkt.shift(1) < 0
fwd4 = fwd_ret(ret_w, 4, start_offset=3)
past4_al = past4.reindex(fwd4.index)
for name, mask in [('上周下跌后', down_wk.reindex(fwd4.index).fillna(False)), ('上周上涨后', ~(down_wk.reindex(fwd4.index).fillna(False)))]:
    F = past4_al[mask]; R = fwd4[mask]
    F = F.dropna(how='all'); R = R.reindex(F.index)
    ic = ic_stats(rank_ic(F, R))
    RESULTS[f'mom4w_cond_{name}'] = ic
    print(f"  mom4w→fwd4w {name}: IC={ic['ic']:+.3f} t={ic['t']:+.2f} (n={ic['n']})")

with open('backtest_results.json', 'w', encoding='utf-8') as f:
    json.dump(RESULTS, f, ensure_ascii=False, indent=1, default=str)
print('\n已保存 backtest_results.json')
