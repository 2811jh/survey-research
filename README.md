# 📊 survey-research

> 问卷研究全流程自动化分析 Skill —— 基础统计 + 交叉分析 + 文本分析 + 综合报告

一个面向 AI 编程助手的 Agent Skill，让 AI 能够自动完成问卷数据的完整分析流程，输出结构化的 Markdown 报告和 Excel 数据文件。

## ⚡ 一键安装

需要 [Node.js](https://nodejs.org/) >= 18

```bash
npx skills add 2811jh/survey-research
```

### 指定安装到特定 Agent

```bash
# 安装到 Claude Code
npx skills add 2811jh/survey-research -a claude-code

# 安装到 Cursor
npx skills add 2811jh/survey-research -a cursor

# 全局安装（所有项目可用）
npx skills add 2811jh/survey-research -g

# 安装到所有已检测到的 Agent
npx skills add 2811jh/survey-research --all
```

### 其他管理命令

```bash
# 查看已安装的 skills
npx skills list

# 检查更新
npx skills check

# 更新到最新版本
npx skills update

# 卸载
npx skills remove survey-research
```

## 🎯 功能特性

| 阶段 | 功能 | 说明 |
|------|------|------|
| **阶段 1** | 数据加载与理解 | 自动识别题型（单选/多选/量表/文本/元数据），智能推荐分组变量 |
| **阶段 2** | 基础统计分析 | 频率分布、Top N 选项、分布形态分析，输出 Excel |
| **阶段 3** | 交叉分析 | 不同人群差异对比，支持选项合并、满意度/NPS 得分计算 |
| **阶段 4** | 归纳式文本分析 | 随机抽样 300 条 → 开放编码 → 维度聚合 → 全量标注，输出 Excel |
| **阶段 5** | 综合报告生成 | Markdown 报告 + 所有 Excel 附件，含策略建议 |

### 🔬 文本分析方法论

采用 **归纳式（自下而上/扎根理论）** 方法：

1. 随机抽样 300 条有效文本（±5.7% 误差，95% 置信度）
2. **逐条开放编码** —— 不预设任何分类框架，标签从数据中自然涌现
3. **标签合并与维度聚合** —— 合并同义标签，聚合为 4-8 个主题维度
4. **全量自动标注** —— 基于发现的维度关键词，对全量数据进行规则化标注
5. 全流程自动执行，报告生成后才询问是否需要全量分析

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

## 🔧 环境要求

- **Python 3.8+**（用于执行分析脚本）
- **Python 依赖**：`pandas`、`numpy`、`openpyxl`

```bash
pip install pandas numpy openpyxl
```

## 💬 使用示例

安装 Skill 后，在你的 AI 编程助手中直接对话：

```
分析 "C:\data\survey_results.xlsx" 问卷数据，
对比不同满意度玩家在其他各题的表现，
同时看看文本题中玩家提到的建议，
最终生成报告
```

AI 将自动：
1. 加载数据并识别题型
2. 生成基础统计 Excel
3. 按满意度分组做交叉分析
4. 抽样 300 条文本做归纳式编码
5. 输出完整报告 + 所有 Excel 附件

## 📊 输出文件

| 文件 | 内容 |
|------|------|
| `{文件名}_分析报告.md` | Markdown 综合报告（含关键发现、策略建议） |
| `{文件名}_基础统计.xlsx` | 各题频率分布详细数据 |
| `{文件名}_交叉分析.xlsx` | 分组差异对比数据 |
| `{文件名}_文本分析.xlsx` | 文本维度总结 + 逐条标注明细 |

## 🤝 兼容性

基于 [Agent Skills 规范](https://agentskills.io)，兼容以下 Agent：

- ✅ Claude Code
- ✅ Cursor
- ✅ Codex
- ✅ Windsurf
- ✅ Cline / Roo Code
- ✅ 以及 [更多 Agent](https://github.com/vercel-labs/skills#supported-agents)

## 📄 License

MIT
