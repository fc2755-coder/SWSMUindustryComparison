# -*- coding: utf-8 -*-
"""ROE周期拐点回测：底部反弹拐点是否应获得更高分数
检验1（信息含量, 2007-2025年末, n≈19期）：底部反弹拐点 → 次年ROE是否继续改善
检验2（收益方向, 2021-2025年末因子→次年真实收益, n=5期）：拐点组 vs 其他组
"""
import pickle
import numpy as np
import pandas as pd

P = pickle.load(open('panels.pkl', 'rb'))
roe_a = P['roe_a'].copy()
ret_w = P['ret_w'] / 100.0

cols = [c for c in roe_a.columns if c < '2026']  # 2006-2025
roe = roe_a[cols].copy()
years = [c[:4] for c in cols]
roe.columns = years
d_roe = roe.diff(axis=1)  # ΔROE

# 底部反弹拐点: ΔROE_t>0 且 ΔROE_{t-1}<0 且 ROE_t 位于自身近5年后30%分位
def classify(ind, y, i):
    if i < 2:
        return None
    r_t, d_t, d_p = roe.loc[ind, y], d_roe.loc[ind, y], d_roe.loc[ind, years[i-1]]
    if any(pd.isna(v) for v in (r_t, d_t, d_p)):
        return None
    hist5 = roe.loc[ind, years[max(0, i-4):i+1]].astype(float).dropna()
    pos = (hist5 < r_t).mean() if len(hist5) >= 3 else 0.5  # 低位→接近0
    if d_t > 0 and d_p < 0 and pos <= 0.30:
        return '底部反弹'
    if d_t > 0 and d_p > 0:
        return '持续向上'
    if d_t < 0 and pos >= 0.70:
        return '高位回落'
    if d_t < 0:
        return '下行'
    return '其他'

cls = pd.DataFrame(index=roe.index, columns=years, dtype=object)
for ind in roe.index:
    for i, y in enumerate(years):
        cls.loc[ind, y] = classify(ind, y, i)

# ===== 检验1: 信息含量（拐点 → 次年ΔROE） =====
print('=' * 72)
print('检验1: 底部反弹拐点 → 次年ROE改善概率（2008-2025年末因子, n≈18期）')
groups = {}
for i, y in enumerate(years[:-1]):
    ny = years[i + 1]
    for ind in roe.index:
        c = cls.loc[ind, y]
        if c is None or pd.isna(d_roe.loc[ind, ny]):
            continue
        groups.setdefault(c, []).append(d_roe.loc[ind, ny])
base = pd.concat([pd.Series(v) for k, v in groups.items() if k != '底部反弹'])
bt = pd.Series(groups.get('底部反弹', []))
print(f"  底部反弹样本: {len(bt)} 个行业-年")
for k in ('持续向上', '高位回落', '下行', '其他'):
    v = groups.get(k, [])
    if v:
        print(f"  {k:<5} 样本{len(v):4d} | 次年ΔROE均值 {np.mean(v):+6.2f}pct | 改善概率 {np.mean([x > 0 for x in v]):.0%}")
print(f"  {'底部反弹':<5} 样本{len(bt):4d} | 次年ΔROE均值 {bt.mean():+6.2f}pct | 改善概率 {np.mean(bt > 0):.0%}")
print(f"  底部反弹 vs 全部其他: ΔROE差 {bt.mean() - base.mean():+.2f}pct")

# 秩和检验（无scipy, 用置换检验）
obs = bt.mean() - base.mean()
pool = np.concatenate([bt.values, base.values])
rng = np.random.default_rng(7)
perm = [abs(rng.permutation(pool)[:len(bt)].mean() - rng.permutation(pool)[len(bt):].mean()) for _ in range(5000)]
# 简化：随机分组置换
perm = []
for _ in range(5000):
    p = rng.permutation(pool)
    perm.append(abs(p[:len(bt)].mean() - p[len(bt):].mean()))
pval = np.mean([p >= abs(obs) for p in perm])
print(f"  置换检验 p值 = {pval:.4f}（H0: 拐点组次年ΔROE不高于其他）")

# 拐点持续性: 拐点后1~2年 ΔROE 均值
pers1, pers2 = [], []
for i, y in enumerate(years[:-2]):
    ny1, ny2 = years[i + 1], years[i + 2]
    for ind in roe.index:
        if cls.loc[ind, y] == '底部反弹' and not pd.isna(d_roe.loc[ind, ny1]):
            pers1.append(d_roe.loc[ind, ny1])
            if not pd.isna(d_roe.loc[ind, ny2]):
                pers2.append(d_roe.loc[ind, ny2])
print(f"  拐点后第1年 ΔROE均值 {np.mean(pers1):+.2f}pct (n={len(pers1)}) | 第2年 {np.mean(pers2):+.2f}pct (n={len(pers2)})")

# ===== 检验2: 收益方向（2021-2025年末因子 → 次年真实收益） =====
print('\n' + '=' * 72)
print('检验2: 底部反弹拐点 → 次年真实收益（周度面板复利, n=5期）')
rows = []
for y_end in ['2021', '2022', '2023', '2024', '2025']:
    y0 = f'{int(y_end)+1}-01-01'
    mask = (ret_w.index >= y0) & (ret_w.index <= f'{int(y_end)+1}-12-31')
    if y_end == '2025':
        mask = ret_w.index >= '2026-01-01'
    if mask.sum() < 20:
        continue
    fwd = (1 + ret_w[mask]).prod() - 1
    yi = years.index(y_end)
    for ind in roe.index:
        c = cls.loc[ind, y_end]
        if c is None or ind not in fwd.index or pd.isna(fwd[ind]):
            continue
        rows.append((y_end, ind, c, fwd[ind] * 100))
df = pd.DataFrame(rows, columns=['年', '行业', '周期', 'fwd'])
piv = df.groupby('周期')['fwd'].agg(['count', 'mean'])
print(piv.round(2).to_string())
bt2 = df[df['周期'] == '底部反弹']['fwd']
ot2 = df[df['周期'] != '底部反弹']['fwd']
if len(bt2) >= 3:
    print(f"\n  底部反弹次年收益均值 {bt2.mean():+.2f}% (n={len(bt2)}) vs 其他 {ot2.mean():+.2f}% (n={len(ot2)}) → 差 {bt2.mean()-ot2.mean():+.2f}pct")
# 拐点 vs 持续向上（核心问题：拐点是否优于顺趋势）
su = df[df['周期'] == '持续向上']['fwd']
if len(bt2) and len(su):
    print(f"  底部反弹 {bt2.mean():+.2f}% vs 持续向上 {su.mean():+.2f}% → 拐点{'优于' if bt2.mean()>su.mean() else '弱于'}顺趋势")

# ===== 检验3: 模型现有ROE打分 vs 拐点加分制对比 =====
print('\n' + '=' * 72)
print('检验3: 两种ROE打分的收益区分度对比（5年期，IC方向）')
for label, F in [('现制: ROE水平分', roe.rank(axis=1, pct=True) * 100),
                 ('拐点制: 拐点=100/持续向上=80/其他=50', None)]:
    ics = []
    for y_end in ['2021', '2022', '2023', '2024', '2025']:
        mask = (ret_w.index >= f'{int(y_end)+1}-01-01') & (ret_w.index <= f'{int(y_end)+1}-12-31')
        if y_end == '2025':
            mask = ret_w.index >= '2026-01-01'
        if mask.sum() < 20:
            continue
        fwd = (1 + ret_w[mask]).prod() - 1
        if label.startswith('现制'):
            f = F[y_end]
        else:
            f = cls[y_end].map({'底部反弹': 100, '持续向上': 80, '其他': 50, '下行': 30, '高位回落': 20}).astype(float)
        m = f.notna() & fwd.notna()
        f2, r2 = f[m].rank(), fwd[m].rank()
        ics.append(np.corrcoef(f2, r2)[0, 1])
    print(f"  {label}: 逐年IC {[round(v, 2) for v in ics]} 均值 {np.mean(ics):+.3f}")
