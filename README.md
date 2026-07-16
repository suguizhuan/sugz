# STAT6 资讯站

一个关于 **STAT6（Signal Transducer and Activator of Transcription 6，信号转导及转录激活因子 6）** 的静态资讯网站。汇集了 STAT6 的基础生物学、信号通路、疾病关联和药物开发等内容，供科普与研究者参考。

## 内容结构

| 页面 | 内容 |
| --- | --- |
| `index.html` | 首页，STAT6 概览与快速导航 |
| `structure.html` | 基因/蛋白结构、结构域组成、激活机制、下游靶基因 |
| `pathway.html` | IL-4/IL-13 → JAK → STAT6 信号通路、Ⅰ/Ⅱ 型受体、正负反馈 |
| `disease.html` | 过敏/哮喘/特应性皮炎、B 细胞淋巴瘤、孤立性纤维瘤、相关药物管线 |
| `research.html` | 关键里程碑、FAQ、代表性文献与外部资源 |
| `assets/style.css` | 共用样式（深/浅色主题、响应式设计） |
| `assets/script.js` | 主题切换、移动端菜单、当前页高亮 |

## 本地预览

无需任何构建工具，直接用浏览器打开 `index.html` 即可，或使用一个简单的静态服务器：

```bash
python3 -m http.server 8000
# 然后访问 http://localhost:8000/
```

## 康百达在研项目信息监控看板

`monitor.html` 是一个数据驱动的**靶点资讯监控看板**：每个在研靶点一栏，每栏下含
**竞争格局 · 临床更新 · 专利更新 · 文章更新 · 其他更新** 五个模块。

| 文件 | 作用 |
| --- | --- |
| `monitor.html` | 监控看板页面（顶部统计、本期更新总览矩阵、最新动态流、按靶点分栏的看板） |
| `assets/monitor.css` | 看板专用样式（复用 `style.css` 的主题变量） |
| `assets/monitor.js` | 从 `data/monitor.json` 渲染看板、矩阵、筛选与搜索 |
| `data/monitor.json` | 结构化数据：靶点、竞品全景、按模块的动态明细（人工策展 + 自动抓取） |
| `scripts/monitor_fetch.py` | 每日自动抓取（ClinicalTrials.gov API v2 / PubMed E-utilities），仅用标准库、幂等合并 |
| `.github/workflows/monitor-daily-update.yml` | 每日定时运行抓取脚本并回写 `data/monitor.json` |

当前监控靶点（7 个）：**KBD1921/BY1921（PARP7）、KBD4409（TLR7/8）、KBD2430（透脑 PARP1）、
KBD2188（GPC3-RDC）、KBD1183（APJ）、KBD1180（c-Kit）、KBD1298/BY1298（DNA-PK）**。
历史数据整理自监控简报（2026.5.15 起），每日增量由 GitHub Actions 自动抓取补充，
自动抓取的条目在页面上标注“自动抓取”，与人工策展内容合并展示。

> 说明：监控简报口径为「在监控时间范围内若某模块有更变则标记；竞争格局仅在竞品进入 IND 及以上阶段时才视为变动」。

### 手动运行抓取

```bash
python3 scripts/monitor_fetch.py   # 拉取增量并写回 data/monitor.json
```

## 免责声明

本站内容整理自公开文献与数据库，仅供学术研究与科普参考，不构成任何医学诊断或治疗建议。
