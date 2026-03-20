---
name: survey-research
description: |
  对问卷原始数据进行全流程自动化分析，包括基础统计（频率分布 + 人口学概览）、
  交叉分析（不同人群差异对比）和文本分析（开放题主题归纳）。
  支持两种数据来源：用户提供本地 CSV/Excel 文件，或通过问卷 ID/名称从问卷系统直接下载（支持清洗）后分析。
  输出 Markdown 摘要报告 + Excel 详细数据报告，支持转换为 Word/TXT 格式。
  当用户要求分析问卷数据、生成问卷调研报告、对比不同人群差异、
  分析开放题文本内容时触发。即使用户没有明确说"问卷"，只要涉及
  调研数据分析、用户反馈分析、满意度分析、NPS 分析等场景也应触发此 skill。
  当用户说"下载并分析问卷"、"清洗问卷然后出报告"、"帮我下载问卷90450"、
  "下载国内问卷数据"、"从问卷系统拉数据"等涉及下载+分析的场景时，也应触发。
  确保在用户说"帮我分析这份问卷"、"分析一下不同性别的差异"、
  "请结合文本题分析"、"下载问卷90450然后分析"、"清洗并导出问卷"等类似表达时使用。
---

# 问卷研究分析（Survey Research）

你是用户研究合成方面的专家，能够将原始的定性和定量数据转化为驱动产品决策的结构化洞察。帮助用户研究员从访谈、问卷、可用性测试、支持数据和行为分析中提炼有效信息。你将使用 `scripts/` 目录下的 Python 脚本，
对用户提供的问卷原始数据进行全流程自动化分析，最终输出 Markdown 摘要报告 + Excel 详细数据。

## 脚本路径

所有脚本位于此 skill 的 `scripts/` 目录下。执行时使用绝对路径：
```
{SKILL_DIR}/scripts/load_and_classify.py
{SKILL_DIR}/scripts/basic_stats.py
{SKILL_DIR}/scripts/crosstab.py
{SKILL_DIR}/scripts/text_extract.py
{SKILL_DIR}/scripts/text_export.py
{SKILL_DIR}/scripts/report_export.py
{SKILL_DIR}/scripts/survey_download.py
{SKILL_DIR}/scripts/refresh_cookie.py
```

其中 `{SKILL_DIR}` 是本 skill 所在目录的绝对路径。

## 依赖要求

脚本依赖 `pandas`、`numpy`、`openpyxl`、`requests`。如果用户环境缺失，先执行：
```bash
pip install pandas numpy openpyxl requests
```

---

## 数据来源路由

在开始分析前，首先判断数据从哪里来。根据用户的表达分为两条路径：

### 路径 A：用户已有本地文件（直接分析）

**触发条件**：用户提供了本地文件路径，或说"分析这份数据"、"帮我看看这个 xlsx"等。

→ 直接跳到下方「阶段 1：数据加载与理解」。

### 路径 B：从问卷系统下载后分析

**触发条件**：用户提到"下载问卷"、"先帮我下数据再分析"、"清洗并分析问卷 xxx"、
"从问卷系统拉数据"、"帮我下载 90450 的数据然后分析"、给了问卷 ID 但没给本地路径等。

**执行步骤**：

1. **确定平台**：
   支持国内（`cn`，survey-game.163.com）和国外（`intl`，survey-game.easebar.com）。
   - 用户提到"国内"、"163" → `--platform cn`
   - 用户提到"国外"、"intl"、"easebar" → `--platform intl`
   - 未说明 → 用 `ask_user_question` 让用户选择

2. **读取下载参考文档并执行**：
   根据用户意图读取对应的 reference 文档：

   | 用户意图 | 读取文档 |
   |----------|----------|
   | 下载问卷数据 | `references/09-survey-download.md` |
   | 清洗/筛选数据 | `references/10-survey-clean.md` |
   | 清洗并下载 | 先读 `10-survey-clean.md` 完成确认，再读 `09-survey-download.md` 执行下载 |
   | Cookie 问题 | `references/11-survey-cookie.md` |

   快速参考命令（`--platform` 放在子命令前面）：
   ```bash
   # 搜索问卷
   python {SKILL_DIR}/scripts/survey_download.py --platform cn search --name "关键词"

   # 下载问卷
   python {SKILL_DIR}/scripts/survey_download.py --platform cn download --id 问卷ID --output_dir "输出目录"

   # 清洗预览 → 确认 → 清洗下载
   python {SKILL_DIR}/scripts/survey_download.py --platform cn clean --id 问卷ID --dry-run
   python {SKILL_DIR}/scripts/survey_download.py --platform cn download --id 问卷ID --clean --output_dir "输出目录"
   ```

3. **确定分析用的文件**：
   下载成功后，脚本返回 JSON 包含文件路径。优先使用 **量化数据（quantified_data）** 文件
   进行分析（列名为编码后的 Q1, Q2...，适合统计分析）。

4. **自动进入分析流程**：
   拿到文件路径后，自动进入下方「阶段 1：数据加载与理解」继续执行。
   **不需要用户再手动指定路径**——下载完直接开始分析，一气呵成。

> 💡 **清洗 + 下载 + 分析 可以一句话完成**：用户说"清洗并下载问卷 90450，然后帮我分析"，
> 你应该依次执行：清洗预览 → 用户确认 → 清洗下载 → 数据加载 → 基础统计 → 报告生成，全程不中断。

> ⚠️ **错误处理**——根据 JSON 中的 `status` 字段决定下一步：
> - `"error"` → 将 `message` 翻译为用户友好语言告知原因
> - `"no_match"` → 告知用户未找到，建议换关键词或提供 ID
> - `"multiple_matches"` → 用 `ask_user_question` 展示列表让用户选择
> - Cookie 失效时脚本会自动弹出浏览器让用户登录，登录后自动继续。
>   **严禁询问用户"选择哪种登录方式"或让用户手动复制 Cookie**，详见 `references/11-survey-cookie.md`

---

## 整体工作流程

根据用户请求，按以下 5 个阶段顺序执行。并非所有阶段都必须执行——
交叉分析和文本分析仅在用户明确要求或暗示需要时触发。

### 阶段 1：数据加载与理解

**目标**：理解问卷结构，识别分组变量，与用户确认分析范围。

1. **获取文件路径**：用户提供数据文件路径，或由路径 B（下载流程）自动传入。

2. **加载并分类数据**：
   ```bash
   python {SKILL_DIR}/scripts/load_and_classify.py --file_path "用户文件路径"
   ```
   脚本输出 JSON，包含：
   - `single_choice`：单选题列表
   - `multi_choice`：多选题（根 → 子列映射）
   - `matrix_scale`：矩阵量表题
   - `text`：文本题列表
   - `meta`：元数据列
   - `valid_for_crosstab`：可用于交叉分析的列

3. **识别分组变量**：从单选题中寻找低唯一值（2-10个选项）且含人口学/行为特征关键词的列，
   如性别、年龄、付费等级、会员类型、使用频率等。

4. **向用户确认**（使用 ask_user_question）：
   - 确认分组变量：如"我识别到以下可能的分组变量：Q17.性别、Q18.年龄段。是否正确？"
   - 确认分析范围：是否需要交叉分析、是否需要文本分析
   - 如果用户已在请求中明确指定（如"分析不同性别的差异"），可跳过确认

### 阶段 2：基础统计分析

**目标**：生成样本概况和各题频率统计。

**始终执行**——这是所有分析的基础。

```bash
python {SKILL_DIR}/scripts/basic_stats.py --file_path "用户文件路径"
```

脚本自动：
- 生成 `{文件名}_基础统计.xlsx`（样本概况 + 各题频率统计）
- stdout 输出 JSON 摘要（总样本量、各题 Top3 选项等）

读取 JSON 摘要，记住关键数据，后续写报告时使用。

> 📖 **参考** `references/05-survey-interpretation.md`：不只看 Top3 选项，关注分布形态（双峰/单峰/偏态），双峰分布意味着用户群体存在明显分化，需要在报告中单独说明。

### 阶段 3：交叉分析（按需）

**触发条件**：用户要求对比不同人群差异（如"分析不同性别的差异"、"请重点分析一下不同 xxx"），
或用户要求全面分析且数据中存在明显的分组变量。

→ **读取 `references/12-crosstab-workflow.md` 获取完整执行步骤。**

### 阶段 4：文本分析（按需）

**触发条件**：用户要求分析文本题/开放题（如"请结合文本题分析"、"看看开放题的建议"），
或用户要求全面分析且数据中存在文本题。

→ **读取 `references/13-text-analysis-workflow.md` 获取完整执行步骤。**

### 阶段 5：生成报告

完成分析后根据用户意图选择对应的报告框架：

| 用户表述特征 | 报告类型 | 读取文档 |
|------------|---------|---------|
| 模糊表述（"分析报告"/"出报告"/"全面分析"/"导出报告"等） | 通用综合报告 | → `references/14-report-workflow.md` |
| 提到满意度/NPS/满意度变化/产品健康度/满意度周报 | 满意度专项报告 | → `references/15-satisfaction-report.md` |

> 💡 未来扩展：新增报告类型只需新建 `references/1X-xxx-report.md` 并在此表中加一行即可。

---

## 重要注意事项

1. **脚本输出解析**：所有脚本通过 stdout 输出 JSON，错误信息输出到 stderr。
   执行后读取 stdout 的 JSON 来获取结果数据。

2. **大数据量处理**：如果文本题回答超过 500 条，先抽样分析：
   ```bash
   python text_extract.py --file_path "..." --column "..." --sample_n 300
   ```

3. **Windows 路径**：在 Windows 上执行脚本时，文件路径使用正斜杠或双反斜杠。

4. **错误处理**：如果脚本报错，检查：
   - 文件路径是否正确
   - 依赖是否已安装
   - Excel 文件是否被其他程序占用

5. **交叉分析列名**：`--col_questions` 中的列名必须与数据中完全匹配。
   从 `load_and_classify.py` 的输出中获取准确的列名。

6. **中文编码**：所有脚本使用 UTF-8 编码，JSON 输出 `ensure_ascii=False`。

---

## 分析方法参考文档

在进行报告撰写和洞察提炼时，参考 `references/` 目录下的方法论文档，以提升分析深度和报告质量。

> 📋 查看 `references/00-index.md` 获取完整索引、按阶段查找指南和文件关系图。

| 文件 | 适用场景 |
|------|---------|
| `references/01-thematic-analysis.md` | 文本分析的维度归纳和主题提炼（阶段 4） |
| `references/02-affinity-mapping.md` | 开放题聚类分组的操作规范（阶段 4） |
| `references/03-triangulation.md` | 交叉分析与文本分析互相印证，综合报告的证据写法（阶段 5） |
| `references/04-interview-analysis.md` | 定性文本分析通用框架：强度信号识别、行为vs态度区分、开放题专用技巧（阶段 4） |
| `references/05-survey-interpretation.md` | 定量数据解读原则、常见分析错误规避（阶段 2、3） |
| `references/06-qual-quant-integration.md` | 综合报告中融合定量与定性发现的写法（阶段 5） |
| `references/07-persona-development.md` | 用户分群特征描述，如需输出用户画像 |
| `references/08-opportunity-sizing.md` | 策略建议的机会规模量化与优先级排序（阶段 5） |
| `references/09-survey-download.md` | 从问卷系统下载数据的完整流程（数据来源路由 B） |
| `references/10-survey-clean.md` | 问卷数据清洗规则与操作流程（数据来源路由 B） |
| `references/11-survey-cookie.md` | Cookie 处理与自动刷新（下载遇到认证问题时） |

### 核心调用时机

- **阶段 4 文本分析**：参考 `01`、`02` 进行维度归纳
- **阶段 5 报告撰写**：参考 `03`、`06` 进行多数据源融合表达；参考 `08` 对建议进行量化支撑
- **解读定量数据时**：参考 `05` 规避常见统计错误
- **发现明显用户分群时**：参考 `07` 描述用户群体特征

---

## 预留模块（后续扩展）

以下模块当前未实现，后续版本补充：

- **舆情分析**：结合外部舆情数据进行综合分析
- **竞品参考**：对比竞品的调研数据或公开报告
