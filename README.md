# SWSMU 行业比较与个股筛选模型

申万一级行业与个股性价比评分模型（五因子：价格回撤 5% / 盈利动能 30% / ROE质量 5% / 估值分位 35% / 公募拥挤度 25%），权重经历史因子有效性回测校准。

## 目录结构

- `build_scoring_model.py` — 主模型生成脚本：读取《多维数据比较YYYYMMDD.xlsx》（68列周度EPS/ROE序列）+ 公募重仓持股（26Q3/Q2）+ 仓位周报 JSON，产出单文件评分页面
- `factor_backtest.py` / `extract_factor_panels.py` / `roe_turning_backtest.py` / `build_backtest_report.py` — 因子历史有效性回测（周度293截面/季度23季/年度5截面，Rank IC 框架）
- `build_dualinnov_page.py` — 双创权重前30 × 公募持仓对比页
- `prep_alloc_backtest.py` / `factor_holdings.json` — 配置思路代理回测的持仓构建
- `posweek_*.json` — 公募基金行业仓位周报提取数据（勾稹通过）
- `recent_ipos_*.csv` — 次新剔除清单（iFinD 上市日期，格式 %Y%m%d）
- `docs/` — 交付成品：各截面模型页、回测报告、演讲稿、双创页

## 评分口径要点

- 盈利分内部：27E EPS 位置 20% / 同比增速 15% / 周度上修 65%，连续≥2周上修+3分、≥3周+5分
- 估值：仅银行/非银/地产用PB，其余用PE；个股估值分=便宜程度50%+利润消化50%，PE/PB分位<50%默认硬筛
- 行业拥挤度 = 50%超配（同层级min-max）+ 50%配置比例2010Q1以来历史分位；安全分=100-拥挤
- 次新剔除：上市未满一年（清单来自iFinD）

## 更新流程

换截面只需改 `build_scoring_model.py` 顶部常量（SOURCE / OUTPUT / ASOF / IPO清单 / posweek文件），页面所有日期自动派生；跑前确认源表无 "Fetching..." 残留。

> 评分为相对比较工具，不构成投资建议。
