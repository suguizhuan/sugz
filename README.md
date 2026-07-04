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

## 免责声明

本站内容整理自公开文献与数据库，仅供学术研究与科普参考，不构成任何医学诊断或治疗建议。
