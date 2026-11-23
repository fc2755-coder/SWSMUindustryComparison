# -*- coding: utf-8 -*-
"""提取因子回测所需的面板数据 → panels.pkl
面板1 周度收益: SWSMU行业轮动强度sheet 31行业×297周(2021-01-08~2026-10-23)
面板2 季度持仓配置: 兴证Q3数据库一级行业sheet 31行业×~90季(2004Q1~2026Q3)
面板3 日频成交拥挤: SWSMU成交额占比分位数 31L1×481日(2024-10-10~2026-09-28)
面板4 年度基本面: SWSMU 估值-PE/估值-PB/盈利-ROE/盈利-净利润同比/股息率 (2006~2026H1)
"""
import openpyxl, pickle
import pandas as pd
import numpy as np
from pathlib import Path

SWSMU = Path(r'C:/Users/Owner/Desktop/周报/【SWSMU】行业比较结果.xlsx')
XING = Path(r'C:/Users/Owner/Desktop/周报/【兴证策略】主动偏股基金数据库_2026Q3(1).xlsx')
OUT = {}

def _num(v):
    return float(v) if isinstance(v, (int, float)) else np.nan

# ---------- 面板1: 周度收益 ----------
wb = openpyxl.load_workbook(SWSMU, data_only=True, read_only=True)
ws = wb['行业轮动强度']
dates = []
for row in ws.iter_rows(min_row=2, max_row=2, min_col=3, values_only=True):
    dates = [v for v in row if v is not None]
data = {}
bench = []
for row in ws.iter_rows(min_row=3, max_row=34, min_col=1, max_col=2 + len(dates), values_only=True):
    name, code = row[0], row[1]
    vals = [_num(v) for v in row[2:2 + len(dates)]]
    if name == '万得偏股基金收盘价':
        bench = vals
    elif code and str(code).endswith('.SI'):
        data[name] = vals
ret_w = pd.DataFrame(data, index=pd.to_datetime([str(d)[:10] for d in dates])).T.T  # 保持 行=日期 列=行业
ret_w = pd.DataFrame(data, index=pd.to_datetime([str(d)[:10] for d in dates]))
print('周度收益面板:', ret_w.shape, '行(日期):', ret_w.index[:2].tolist(), '->', ret_w.index[-2:].tolist(), '| 列(行业):', len(ret_w.columns))
OUT['ret_w'] = ret_w
OUT['bench_close'] = pd.Series(bench, index=ret_w.index)

# ---------- 面板3: 日频成交拥挤(分位数) ----------
ws = wb['成交额占比分位数']
dd = []
for row in ws.iter_rows(min_row=3, max_row=3, min_col=4, values_only=True):
    dd = [v for v in row if v is not None]
to_data = {}
collect = False
for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=2, max_col=3 + len(dd), values_only=True):
    name = row[0]
    if name == '农林牧渔':
        collect = True
    if not collect:
        continue
    # L1共31个，读到美容护理为止
    vals = [_num(v) for v in row[2:2 + len(dd)]]
    to_data[name] = vals
    if name == '美容护理':
        break
to_pct = pd.DataFrame(to_data, index=pd.to_datetime([str(d)[:10] for d in dd]))
print('成交拥挤面板:', to_pct.shape, to_pct.index[0], '->', to_pct.index[-1])
OUT['turnover_pct_d'] = to_pct

# ---------- 面板4: 年度基本面 ----------
def read_annual(sheet, first_row, last_row_hint='美容护理', val_col0=3, nyears=22):
    ws = wb[sheet]
    dates_row = []
    for row in ws.iter_rows(min_row=3, max_row=3, min_col=val_col0 + 1, max_col=val_col0 + nyears, values_only=True):
        dates_row = list(row)
    out = {}
    started = False
    for row in ws.iter_rows(min_row=first_row, max_row=ws.max_row, min_col=2, max_col=val_col0 + nyears, values_only=True):
        name = row[0]
        if name == '农林牧渔':
            started = True
        if not started:
            continue
        out[name] = [_num(v) for v in row[2:2 + nyears]]
        if name == last_row_hint:
            break
    df = pd.DataFrame(out, index=[str(d)[:10] for d in dates_row]).T
    return df

years = [f'{y}-12-31' for y in range(2006, 2028)][:22]  # 2006..2027占位, 实际按表头
pe_a = read_annual('估值-PE', 4)
pb_a = read_annual('估值-PB', 4)
roe_a = read_annual('盈利-ROE', 4)
np_a = read_annual('盈利-净利润同比', 4)
dy_a = read_annual('股息率', 4)
# 统一索引名
for nm, df in [('pe', pe_a), ('pb', pb_a), ('roe', roe_a), ('npg', np_a), ('dy', dy_a)]:
    df.index = [i.replace('(', '').replace(')', '') for i in df.index]
    print(f'年度{nrm if (nrm:=nm) else ""}面板:', df.shape, '年份:', df.columns[0], '->', df.columns[-1])
    OUT[nm + '_a'] = df

# 适用估值方法（风格行业池 L22-52 = 申万一级31个）
ws = wb['风格行业池']
meth = {}
started = False
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=4, values_only=True):
    name, code, m = row[0], row[1], row[2]
    if name == '农林牧渔':
        started = True
    if started and code and str(code).endswith('.SI'):
        meth[name] = m
        if name == '美容护理':
            break
OUT['method'] = meth
print('估值方法:', len(meth), '个; PB行业:', [k for k, v in meth.items() if v == 'PB'])
wb.close()

# ---------- 面板2: 季度持仓配置 ----------
wb = openpyxl.load_workbook(XING, data_only=True, read_only=True)
ws = wb['一级行业']
# 表头在第1行: 列3(0基)起为季度日期，只收 datetime（跳过"变化"等文字列）
import datetime as _dt
qdates = []
for row in ws.iter_rows(min_row=1, max_row=1, min_col=4, values_only=True):
    for v in row:
        if isinstance(v, (_dt.datetime, _dt.date)):
            qdates.append(str(v)[:10] if isinstance(v, str) else f'{v.year:04d}-{v.month:02d}-{v.day:02d}')
qdata = {}
collect = False
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=3 + len(qdates), values_only=True):
    name = row[0]
    if name == '农林牧渔':
        collect = True
    if not collect:
        continue
    qdata[name] = [_num(v) * 100 for v in row[1:1 + len(qdates)]]
    if name == '美容护理':
        break
alloc_q = pd.DataFrame(qdata, index=pd.to_datetime(qdates))
alloc_q = alloc_q.dropna(how='all').dropna(axis=1, how='all')
print('季度配置面板:', alloc_q.shape, '季度:', alloc_q.index[0], '->', alloc_q.index[-1], '| 行业:', alloc_q.shape[1])
OUT['alloc_q'] = alloc_q
wb.close()

with open('panels.pkl', 'wb') as f:
    pickle.dump(OUT, f)
print('\n已保存 panels.pkl; keys:', list(OUT.keys()))
