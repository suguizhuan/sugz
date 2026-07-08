# TLR7/8 拮抗剂资讯站

一个关于 **TLR7/8（Toll 样受体 7 与 8）拮抗剂（抑制剂）** 的静态资讯与监控网站。汇集了 TLR7/8 的基础生物学、信号通路、自身免疫疾病关联与在研拮抗药物管线，并每日自动拉取 PubMed 与 ClinicalTrials.gov 的最新条目，供科普与研究者参考。

与同仓库的 STAT6 资讯站采用完全相同的页面结构、样式系统与每日更新机制。

## 内容结构

| 页面 | 内容 |
| --- | --- |
| `index.html` | 首页，TLR7/8 拮抗剂概览与快速导航 |
| `structure.html` | 受体结构域组成（LRR / Z-loop / TIR）、切割激活、拮抗剂作用机制、TLR7 vs TLR8 |
| `pathway.html` | ssRNA → TLR7/8 → MyD88 → IRAK → NF-κB / IRF7 通路、促炎与干扰素两条分支、拮抗节点 |
| `disease.html` | SLE、皮肤型狼疮、狼疮性肾炎、皮肌炎、干燥综合征、相关药物管线 |
| `research.html` | 每日自动更新的「最新进展」、在研管线对比、关键里程碑、FAQ、文献与外部资源 |
| `deals.html` | 交易与投融资：TLR7/8 拮抗剂靶点的交易、合作、许可与资本动向，及与 STAT6 赛道的对比 |
| `assets/style.css` | 共用样式（深/浅色主题、响应式设计） |
| `assets/script.js` | 主题切换、移动端菜单、当前页高亮 |
| `data/updates.json` | 策展 + 自动抓取的进展数据 |
| `scripts/fetch_latest.py` | 每日抓取 PubMed / ClinicalTrials.gov 并回填 research.html |

## 每日自动更新

`.github/workflows/tlr78-daily-update.yml` 每日 UTC 03:41 运行 `tlr78/scripts/fetch_latest.py`：

1. 以「TLR7/8 + antagonist/inhibitor/lupus」为条件检索 PubMed 近日文献；
2. 检索 ClinicalTrials.gov 上 enpatoran / afimetoran / MHV370 / E-6742 等相关试验；
3. 合并去重后写入 `data/updates.json`，并重新生成 `research.html` 的「最新进展」区块；
4. 若有变化则提交并推送回 `master`。

本地手动刷新：

```bash
python tlr78/scripts/fetch_latest.py
```

## 本地预览

无需任何构建工具，直接用浏览器打开 `tlr78/index.html` 即可，或使用一个简单的静态服务器：

```bash
python3 -m http.server 8000
# 然后访问 http://localhost:8000/tlr78/
```

## 免责声明

本站内容整理自公开文献与数据库，仅供学术研究与科普参考，不构成任何医学诊断或治疗建议。药物开发状态请以各公司公告与监管机构信息为准。
