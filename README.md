# 📊 Survey Research — AI 驱动的问卷全流程自动分析工具

> 问卷研究全流程自动化分析 Skill —— 基础统计 + 交叉分析 + 归纳式文本分析 + 多格式报告

---

## ⚡ 安装

### ✨一键安装 （npx）

> ⚠️ **注意**：需要 [Node.js](https://nodejs.org/) >= 18

```bash
npx skills add 2811jh/survey-research
```

```bash
# 指定 Agent
npx skills add 2811jh/survey-research -a claude-code

# 全局安装
npx skills add 2811jh/survey-research -g

# 安装到所有 Agent
npx skills add 2811jh/survey-research --all

# 其他管理命令
npx skills list      # 查看已安装
npx skills check     # 检查更新
npx skills update    # 更新
npx skills remove survey-research  # 卸载
```


### ✨手动安装
> ⚠️ **注意**：`npx skills` 会将 skill 安装到 `~/.agents/skills/` 目录。如果你的 Agent（如 Claude Code）使用的是 `~/.claude/skills/` 目录，请使用下方的手动安装方式。

如果你使用的是 **Claude Code / CodeMaker**，直接通过 Git 安装到原生 skills 目录：

```bash
# macOS / Linux
git clone https://github.com/2811jh/survey-research.git ~/.claude/skills/survey-research

# Windows
git clone https://github.com/2811jh/survey-research.git "%USERPROFILE%\.claude\skills\survey-research"
```

如果你使用的是 **Cursor / Cline / Codex** 等其他 Agent：

```bash
# macOS / Linux
git clone https://github.com/2811jh/survey-research.git ~/.agents/skills/survey-research

# Windows
git clone https://github.com/2811jh/survey-research.git "%USERPROFILE%\.agents\skills\survey-research"
```

更新到最新版本：

```bash
cd ~/.claude/skills/survey-research && git pull
```

---

## � 环境要求

- **Python 3.8+**（用于执行分析脚本）
- **Python 依赖**：`pandas`、`numpy`、`openpyxl`（Word 报告额外需要 `python-docx`）

```bash
pip install pandas numpy openpyxl python-docx
```

---


## 一、这是什么？

这是一个 🤖 **AI 驱动的问卷全流程分析工具**——从基础统计到交叉分析到文本编码，一句话搞定。

我们做用户调研时，问卷回收后要做的事情太多了——频率统计、人群差异对比、开放题归纳、写分析报告、排版 Excel……每一步都是重复劳动，每一步都在消耗你的精力。

😩 **传统流程有多痛？**

> 先在 Excel 里拉频率表，再用 SPSS 跑交叉分析，然后手动逐条读几千条开放题贴标签、分类、统计，最后还要把所有发现整合成一份有洞察的报告。十几道题交叉下来，一天就没了。

✨ **现在只需一句话：**

> 「帮我分析 xxx.xlsx 问卷数据，对比不同性别和职业的差异，看看文本题中玩家的建议，生成报告」

AI 就会自动完成 **全部流程**：

> 📂 加载数据 → 🏷️ 智能识别题型 → 📊 基础统计 → 🔀 交叉分组对比 → 📝 300条文本抽样编码 → 🧩 主题维度聚合 → 📋 撰写洞察报告 → 📁 导出全套 Excel + Markdown

**你只说一句话，工具跑完五个阶段。中间不打断、不追问、不需要你操心任何参数。**

---

## 二、核心优势

### 🚀 1. 全流程自动化 — 10 分钟替代两天活

| 传统方式 | 现在 |
|---------|------|
| 基础统计：手动建频率表 → **2小时** | 一键生成 → **30秒** |
| 交叉分析：SPSS 配置+逐题对比 → **半天** | 自然语言指定分组 → **2分钟** |
| 文本编码：逐条读+贴标签+归类 → **1~2天** | AI 抽样300条归纳式编码 → **5分钟** |
| 整合报告：拼数据+写结论+排版 → **半天** | 自动生成含策略建议的完整报告 → **1分钟** |

以前要两个工作日才能交付的分析报告，现在 **一杯咖啡的时间就出来了**。

### 🧠 2. 自然语言交互 — 零学习成本

不需要 SPSS，不需要写公式，不需要配置参数，不需要学任何新软件。**用中文说人话就行：**

> 「分析这份问卷，按满意度分成满意、一般、不满意三组做交叉分析，顺便看看开放题的建议，最后给我一份 Word 报告」

甚至可以更随意：

> 「帮我看看不同性别玩家有啥区别，文本题也分析一下」

工具会自动理解你的意图、找到正确的列、选择合适的方法，然后一口气跑完。

### 🔬 3. 文本分析用真方法 — 不是关键词云

市面上大多数工具做文本分析，要么生成词云，要么预设几个分类往里塞。**这个工具用的是正经的归纳式编码方法（扎根理论）：**

1. 随机抽样 300 条有效文本
2. **逐条阅读、逐条打标**——不预设任何框架
3. 让主题维度**从数据中自然涌现**
4. 合并同义标签，聚合为 4-8 个核心维度
5. 统计占比，选取代表性原声

**先看数据再分类，不是先分类再塞数据。** 这是研究方法论上的根本区别。

### 📊 4. 专业报告输出 — 拿来即用

一次分析，自动生成 **全套交付物**：

| 产出 | 内容 |
|------|------|
| 📋 **Markdown 报告** | 关键发现 + 交叉洞察 + 文本主题 + 策略建议，结构完整、逻辑清晰 |
| 📁 **基础统计 Excel** | 样本概况 + 各题频率分布，含 DataBar 可视化 |
| 📁 **交叉分析 Excel** | 列百分比表 + 得分分析 + 差异摘要，精美排版 |
| 📁 **文本分析 Excel** | 总结概览 + 逐条标注明细，每条文本都标了属于哪个维度 |

还不够？**一句话转格式**：

> 「请把报告转成 Word」 → .docx  
> 「导出成 Excel」 → .xlsx  
> 「给我纯文本版」 → .txt

Markdown / Word / Excel / TXT，**想要什么格式就什么格式**。

### 🎯 5. 智能识别题型 — 零配置开箱即用

扔进去一个 Excel 或 CSV，工具 **自动识别**：

- ✅ 单选题、多选题（自动展开子选项）
- ✅ 量表题（自动计算满意度均值 / NPS 得分）
- ✅ 文本题（自动归纳编码）
- ✅ 元数据列（自动排除，不污染分析）

**不需要你标注哪道是单选、哪道是多选、哪道是文本。** 全部自动搞定。

---

## 三、一句话总结

> **你负责提问，AI 负责分析。**  
> 从 13,000 条问卷数据到一份完整的分析报告，全程自动化，中间零操作。  
> 把时间花在**读洞察、做决策**上，而不是花在拉表格、贴标签上。


## 📁 项目结构

```
survey-research/
├── SKILL.md              # Skill 主文件（触发入口 + 完整工作流程）
├── scripts/              # Python 自动化脚本
│   ├── load_and_classify.py   # 数据加载与题型分类
│   ├── basic_stats.py         # 基础统计分析
│   ├── crosstab.py            # 交叉分析
│   ├── text_extract.py        # 文本提取与抽样
│   ├── text_export.py         # 文本分析 Excel 导出
│   ├── report_export.py       # 报告格式转换（md→docx/xlsx/txt）
│   ├── _styles.py             # Excel 样式工具
│   └── requirements.txt       # Python 依赖
└── references/           # 方法论参考文档
    ├── 00-index.md            # 索引与导航
    ├── 01-thematic-analysis.md    # 主题分析六步法
    ├── 02-affinity-mapping.md     # 亲和图法
    ├── 03-triangulation.md        # 三角验证
    ├── 04-interview-analysis.md   # 定性分析框架
    ├── 05-survey-interpretation.md # 定量数据解读
    ├── 06-qual-quant-integration.md # 定性定量融合
    ├── 07-persona-development.md   # 用户画像
    └── 08-opportunity-sizing.md    # 机会规模量化
```

---

## 📊 输出文件

| 文件 | 内容 |
|------|------|
| `{文件名}_分析报告.md` | Markdown 综合报告（含关键发现、策略建议） |
| `{文件名}_分析报告.docx` | Word 版报告（按需生成） |
| `{文件名}_基础统计.xlsx` | 各题频率分布详细数据 |
| `{文件名}_交叉分析.xlsx` | 分组差异对比数据 |
| `{文件名}_文本分析.xlsx` | 文本维度总结 + 逐条标注明细 |

---

## 🤝 兼容性

基于 [Agent Skills 规范](https://agentskills.io)，兼容以下 Agent：

- ✅ Claude Code / CodeMaker
- ✅ Cursor
- ✅ Codex
- ✅ Windsurf
- ✅ Cline / Roo Code
- ✅ 以及 [更多 Agent](https://github.com/vercel-labs/skills#supported-agents)

## 📄 License

MIT
