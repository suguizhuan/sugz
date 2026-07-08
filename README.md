# PARP7 抑制剂资讯监控站

一个关于 **PARP7（TIPARP / ARTD14，TCDD-inducible poly-ADP-ribose polymerase）抑制剂** 的静态资讯监控网站。
PARP7 是一种单-ADP-核糖基转移酶（mono-ART），通过 MAR 化 TBK1 抑制 I 型干扰素、帮助肿瘤逃避免疫；抑制它可重新激活抗肿瘤免疫，
是免疫肿瘤学中机制新颖、竞争格局清晰的小分子靶点。本站汇集其基础生物学、临床管线、竞争情报与商业动态，供研究与竞品分析参考。

## 内容结构

| 页面 | 模块 | 内容 |
| --- | --- | --- |
| `index.html` | 首页 | PARP7 概览、基本信息与模块导航 |
| `structure.html` | 结构与功能 | 基因/蛋白结构、WWE/锌指/催化域、MARylation 机制、底物 |
| `pathway.html` | 信号通路 | STING/RIG-I→TBK1→IRF3 I 型干扰素轴、AHR 反馈环、抑制剂作用机制 |
| `disease.html` | 疾病与治疗 | 3q 扩增鳞癌（肺鳞癌/头颈鳞癌）、生物标志物、单药与免疫联合策略 |
| `landscape.html` | 竞争格局与里程碑 | 在研管线快照、格局解读、关键里程碑时间线 |
| `news.html` | 每日资讯 | ClinicalTrials.gov、PubMed、BioSpace/新闻爬虫的每日自动条目 |
| `companies.html` | 公司 | Ribon Therapeutics 及竞争对手/监控名单 |
| `deals.html` | 交易与投融资 | 授权、并购、融资与里程碑付款监控看板 |
| `patents.html` | 专利申请进展 | WIPO PATENTSCOPE 的 PARP7/TIPARP 专利检索与监控框架 |
| `assets/style.css` | — | 共用样式（深/浅色主题、响应式设计） |
| `assets/script.js` | — | 主题切换、移动端菜单、当前页高亮 |

## 每日自动更新

`每日资讯` 页面由 GitHub Actions 每日（UTC 03:17）自动刷新：

- `scripts/fetch_latest.py`：使用 Python 标准库（无第三方依赖）抓取
  - **PubMed** — NCBI E-utilities，检索 `PARP7`/`TIPARP`（标题/摘要）
  - **ClinicalTrials.gov** — 官方 API v2，检索 `PARP7 / RBN-2397 / TIPARP`
  - **新闻爬虫** — Google News RSS（覆盖 BioSpace 等媒体），检索 `PARP7 / RBN-2397 / TIPARP inhibitor`
- 抓取结果写入 `data/updates.json`，并注入 `news.html` 中 `<!-- LIVE-NEWS:START/END -->` 标记之间的区块。
- 工作流：`.github/workflows/parp7-daily-update.yml`。

手动运行：

```bash
python3 scripts/fetch_latest.py                   # 抓取并更新
python3 scripts/fetch_latest.py --featured-only   # 仅重渲染，不联网抓取
```

## 本地预览

无需任何构建工具，直接用浏览器打开 `index.html`，或启动一个简单的静态服务器：

```bash
python3 -m http.server 8000
# 然后访问 http://localhost:8000/
```

## 免责声明

本站内容整理自公开文献、临床数据库、新闻与专利检索，仅供学术研究与竞品情报参考，
不构成任何医学诊断、治疗、投资或法律建议。
