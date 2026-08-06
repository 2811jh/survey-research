# 设计文档：问卷时间异动诊断工具 (survey_drift.py)

**日期**：2026-08-06
**状态**：待实施
**目标**：让 survey-research skill 支持对**单份回流问卷**按时间（周/月/天）自动分桶，逐题跑显著性检验，由 LLM 写一句话结论，输出 Excel 异动诊断报告，解决 MC 回流体验反馈滞后、无法及时发现数据异动的问题。

---

## 1. 背景与目标

### 背景
MC（游戏）持续回流问卷每天/每周都在收数据。目前 survey-research 只能对**全样本**做基础统计、交叉分析、文本分析，无法回答"这周相比上周，满意度/NPS/某个选项占比有没有显著变化、哪里出了问题"。人工逐题逐期比对成本高、滞后。

### 与已有 `survey_compare.py`（未实现）的区别
| 维度 | survey_compare（旧设计，未实现） | survey_drift（本设计） |
|------|------|------|
| 数据来源 | 多份独立 CSV（跨赛季/版本） | **单份** CSV 按时间列分桶 |
| 对比结构 | 文件 A vs 文件 B | 相邻期对比 + 全时间线展示 |
| 显著性检验 | 明确排除（仅描述性 Δ） | **核心能力**（自动选检验法 + 双门槛判异动） |
| 结论 | 无 | **LLM 逐题一句话结论**写入 Excel |
| 用途 | 横向复盘 | 及时诊断异动 |

两者是不同链路，**不复用** survey_compare（其脚本本就未实现）。survey_drift 独立新建。

### 目标
- 输入一份含答题时间列的量化 CSV + 用户指定粒度（周/月/天），一句话触发完整分析
- 逐题（满意度量表、NPS、单选、多选）跑**相邻期**显著性检验，双门槛判定"异动"
- 全时间线各期数值以表格呈现，看趋势
- LLM 为每道题写一句话初步结论，写进 Excel
- 输出多 Sheet Excel：指标总览 / 逐题异动明细 / 异动汇总 / 方法与样本

---

## 2. 新增/改动文件清单

| 文件路径 | 类型 | 说明 |
|---------|------|------|
| `scripts/survey_drift.py` | 新建 | 核心脚本，两子命令 `analyze` / `export` |
| `references/18-drift-workflow.md` | 新建 | 时间异动诊断工作流 reference |
| `SKILL.md` | 改动 | 新增"阶段 6：时间异动诊断（按需）"触发条件 + 跳转；后续操作提示补充相关项 |

---

## 3. 数据流与编排（一句话触发）

从用户视角是一句话，Agent 自动串联三步、中途不打断（对齐现有"下载+分析一气呵成"与"text_extract→agent→export"范式）：

```
量化CSV ──[1] analyze──> drift_findings.json  (+ 分桶统计中间数据)
                              │
                    [2] Agent(LLM) 逐题读 findings 写一句话结论
                              │  → conclusions.json
                              │
                         ──[3] export──> 最终 Excel（数据 + AI结论 + 异动诊断）
```

- **[1] analyze**：复用 `load_and_classify.py` 的题型分类逻辑（导入或子进程调用）；用时间列分桶；逐题相邻期检验；把每题的各期数值、相邻期 Δ、检验方法、p 值、是否异动标记写入 `drift_findings.json`。同时把逐题各期明细缓存进同一 JSON（供 export 直接出表，避免重算）。
- **[2] Agent 写结论**：Agent 读 `drift_findings.json`，**逐题**写一句话结论（结论只做定性判断，数字由脚本已算好）。异动题重点写"哪个选项/指标、朝哪个方向、幅度多大、是否显著"；无异动题写"本期无显著变化"。结果写入 `conclusions.json`（`{题目: 结论文本}` 映射）。
- **[3] export**：读 `drift_findings.json` + `conclusions.json`，把 AI 结论注入对应列，生成最终 Excel。

> 💡 结论文本"交给 LLM"= 交给编排本流程的 Agent 本体来写，脚本不调用外部 LLM API（无 API key 依赖，与现有 text 分析同构）。

---

## 4. CLI 接口

### 4.1 `analyze` 子命令

```bash
python survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --granularity week \
  [--time_col "结束答题时间"] \
  [--nps_col "Q51.您有多大可能将...推荐给朋友？"] \
  [--satisfaction_cols "Q1.整体满意度" "Q13.赛季满意度"] \
  [--min_n 30] \
  [--findings_out "drift_findings.json"]
```

| 参数 | 必填 | 默认 | 说明 |
|------|------|------|------|
| `--file_path` | ✅ | — | 量化原始 CSV（列名为编码后 Q1/Q2…） |
| `--granularity` | ✅ | — | `week` / `month` / `day`，分桶粒度（用户指定） |
| `--time_col` | ❌ | `结束答题时间` | 分桶用的时间列 |
| `--nps_col` | ❌ | 自动识别 | NPS 题列名；不传则按关键词（"推荐"+0~10 量表）自动识别 |
| `--satisfaction_cols` | ❌ | 自动识别 | 满意度量表题列名，可多个；不传则按关键词（"满意度"+量表）自动识别 |
| `--min_n` | ❌ | `30` | 桶内最小有效样本量，低于此不判异动 |
| `--findings_out` | ❌ | `{file同目录}/drift_findings.json` | findings 输出路径 |

**stdout JSON**：
```json
{
  "status": "success",
  "granularity": "week",
  "buckets": ["第30周（7.21-7.27）", "第31周（7.28-8.3）", "第32周（8.4-8.10）"],
  "bucket_sizes": {"第30周（7.21-7.27）": 812, "第31周（7.28-8.3）": 790, "第32周（8.4-8.10）": 120},
  "low_n_buckets": ["第32周（8.4-8.10）"],
  "questions_total": 45,
  "questions_with_drift": 6,
  "findings_out": "…/drift_findings.json",
  "nps_col": "Q51…",
  "satisfaction_cols": ["Q1…"]
}
```

若 `--time_col` 不存在、或识别不到 NPS/满意度题 → `status:"need_input"` + `message`，由 Agent 用 `ask_user_question` 追问用户后重跑。

### 4.2 `export` 子命令

```bash
python survey_drift.py export \
  --findings "drift_findings.json" \
  --conclusions "conclusions.json" \
  [--output_path "回流异动诊断_按周_{timestamp}.xlsx"]
```

| 参数 | 必填 | 说明 |
|------|------|------|
| `--findings` | ✅ | analyze 产出的 drift_findings.json |
| `--conclusions` | ❌ | Agent 写的结论 JSON；不传则 AI 结论列留空 |
| `--output_path` | ❌ | 默认与 findings 同目录，含粒度与时间戳 |

**stdout JSON**：`{"status":"success","output_path":"…","sheets":[...]}`

---

## 5. drift_findings.json 结构

```json
{
  "granularity": "week",
  "buckets": ["第30周（7.21-7.27）", "第31周（7.28-8.3）", "第32周（8.4-8.10）"],
  "bucket_sizes": {"第30周（7.21-7.27）": 812, "…": 790},
  "low_n_buckets": ["第32周（8.4-8.10）"],
  "metrics": [
    {
      "name": "整体满意度均分",
      "type": "satisfaction_mean",
      "source_col": "Q1…",
      "by_bucket": {"第30周…": 3.58, "第31周…": 3.43, "第32周…": 3.60},
      "adjacent": [
        {"from": "第31周…", "to": "第30周…", "delta": 0.15, "test": "t_test",
         "p": 0.012, "significant": true, "drift": true, "direction": "up"}
      ]
    },
    {
      "name": "NPS",
      "type": "nps",
      "by_bucket": {"第30周…": 12.3, "第31周…": 9.8},
      "adjacent": [{"…": "…", "test": "two_prop_z", "…": "…"}]
    }
  ],
  "questions": [
    {
      "question": "Q7.您对本次活动的评价（单选）",
      "type": "single_choice",
      "options": ["非常满意", "比较满意", "一般", "不满意"],
      "by_bucket": {
        "第30周…": {"非常满意": 0.22, "比较满意": 0.35, "…": "…"},
        "第31周…": {"非常满意": 0.18, "…": "…"}
      },
      "overall_test": {"test": "chi_square", "p": 0.03, "significant": true},
      "adjacent_option_tests": [
        {"option": "非常满意", "from": "第31周…", "to": "第30周…",
         "delta_pp": 4.0, "test": "two_prop_z", "p": 0.02,
         "significant": true, "drift": true, "direction": "up"}
      ],
      "drift": true,
      "low_n": false
    }
  ]
}
```

- `metrics`：满意度均分、NPS 等总览指标。
- `questions`：所有单选/多选/量表题逐题明细。
- `drift`（题级）：该题任一相邻期任一选项/指标触发双门槛即为 true。
- `low_n`：该题参与对比的某桶样本不足，结论仅供参考。

---

## 6. 统计方法（自动按题型选）

| 指标/题型 | `type` | 对比量 | 检验方法 | 备注 |
|-----------|--------|--------|----------|------|
| 满意度/量表题 | `satisfaction_mean` | 相邻期均分 | 两样本 t 检验 | 桶内 n<30 或分布严重偏态 → Mann-Whitney U |
| NPS | `nps` | 推荐者%−贬损者% | 两比例 z 检验 | 0~10：9-10 推荐者、0-6 贬损者 |
| 单选题 | `single_choice` | 整体分布 + 各选项占比 | 整体卡方 + 单选项两比例 z 检验 | 卡方看整体是否变，z 检验定位到具体选项 |
| 多选题 | `multi_choice` | 各选项勾选率 | 每选项两比例 z 检验 | 每选项独立二分类（勾/未勾） |

### 判定"异动"（双门槛，硬性）
一个相邻期对比被标 `drift=true` 需同时满足：
1. **统计显著**：`p < 0.05`
2. **实际差异达标**：占比 `|Δ| ≥ 5pp`，或均分 `|Δ| ≥ 0.1`（NPS 用 pp 门槛）

### 样本量守卫
- 参与对比的任一桶有效样本 `n < min_n`（默认 30）→ 该对比不判异动，`low_n=true`，结论标注"样本不足，仅供参考"。
- `analyze` 的 stdout 与 findings 都列出 `low_n_buckets`，方便 Agent 在结论中提示。

### 依赖
新增 `scipy`（`ttest_ind`、`mannwhitneyu`、`chi2_contingency`、`proportions_ztest` 可用 statsmodels 或自实现两比例 z）。优先 `scipy`；两比例 z 检验可用 scipy 正态近似自实现，避免引入 statsmodels。更新 `requirements.txt` 增加 `scipy`。

---

## 7. Excel 结构（4 Sheet）

沿用 `_styles.py` 主题与 `survey_compare` 设计里的趋势色规范。

### Sheet 1：📊 指标总览
满意度均分、NPS（及不满意率等衍生指标）按时间线全期展示，末列给最新期 vs 上期趋势。

| 指标 | 第30周 | 第31周 | 第32周(最新) | 最新vs上期 | 是否显著 |
|------|--------|--------|--------------|-----------|---------|
| 样本量 | 812 | 790 | 120⚠ | — | — |
| 整体满意度均分 | 3.58 | 3.43 | 3.60 | ▲ +0.17 | ✅ p=0.01 |
| NPS | 12.3 | 9.8 | 11.0 | ▲ +1.2pp | ✗ 不显著 |

- 最新桶样本不足打 ⚠ 并在趋势列标"样本不足"。

### Sheet 2：📈 逐题异动明细
每道题占多行（每个选项一行），题名首行显示、其余合并单元格。

| 题目 | 选项 | 第30周 | 第31周 | 第32周 | 最新Δ(pp) | p | 异动 | AI结论 |
|------|------|--------|--------|--------|-----------|---|------|--------|
| Q7 活动评价 | 非常满意 | 22% | 18% | 24% | +6pp | 0.02 | ✅ | 首行合并显示整题一句话结论 |
| | 比较满意 | 35% | 32% | 33% | +1pp | 0.6 | | |

- **AI结论列**按题合并单元格，一题一句话（来自 conclusions.json）。
- 异动行（双门槛命中）：正向浅绿 `E2EFDA`、负向浅红 `FCE4EC` 高亮。
- 全时间线各期占比全列出（看趋势），Δ 与 p 只针对最新相邻期。

### Sheet 3：⚠️ 异动汇总（诊断视图）
只列 `drift=true` 的题目/指标，供快速定位"哪里出问题"。

| 题目/指标 | 变化选项 | 方向 | 幅度 | 显著性 | AI结论 |
|-----------|---------|------|------|--------|--------|
| 整体满意度 | 均分 | ▲ | +0.17 | p=0.01 | … |
| Q7 活动评价 | 非常满意 | ▲ | +6pp | p=0.02 | … |

- 无任何异动时该 Sheet 显示"本期各指标/题目均无显著异动"。

### Sheet 4：ℹ️ 方法与样本
分桶方式、各桶样本量分布、检验方法映射表、双门槛阈值、样本不足桶清单、免责说明。

---

## 8. 集成到 skill

### 8.1 `references/18-drift-workflow.md`（新建）
含：触发条件、两步 CLI 命令、granularity 选择、findings/conclusions JSON 格式、Agent 写结论的规范（只写定性、异动题重点写方向幅度、low_n 提示）、Excel 结构说明、错误/need_input 处理、后续操作提示。

### 8.2 `SKILL.md` 更新
在"整体工作流程"新增：

```
### 阶段 6：时间异动诊断（按需）

**触发条件**：用户有单份含时间列的回流问卷数据，想按周/月/天自动对比、
诊断满意度/NPS/单选/多选的异动（如"按周诊断这份回流数据的变化"、
"逐题对比各月满意度和NPS有没有显著变化"）。

→ **读取 `references/18-drift-workflow.md` 获取完整执行步骤。**
```

同时把 `survey_drift.py` 加入"脚本路径"清单，"后续操作提示"分析方面补一条"做时间异动诊断（按周/月/天对比满意度/NPS/单选/多选并定位显著变化）"。

---

## 9. 触发词
"按周/月/天分析回流"、"诊断回流数据异动"、"这周相比上周满意度有没有变化"、
"逐题对比各周/各月/各天的满意度和NPS"、"回流数据有没有异常波动"、
"自动对比各期问卷指标"、"哪个指标这期掉了/涨了、显著吗"。

---

## 10. 本次范围外（后续阶段）
- **文本"新增反馈"检测**：按时间桶分别跑文本分析、检测新增/激增主题（下一阶段，工作量大）
- 图表可视化 / HTML 报告（v1 只出 Excel）
- 多重比较 FDR 校正（v1 用双门槛降噪，暂不做 BH 校正）
- 自动检测并下载最新回流数据（由用户传路径）
- 相邻期以外的对比结构（固定基准期 / 全序列趋势检验）

---

## 11. 边界与单元划分（便于分块实现与测试）
| 单元 | 职责 | 依赖 | 可独立测试点 |
|------|------|------|-------------|
| 分桶器 `bucketize()` | 按 granularity 把时间列切成有序桶 | enrich_columns 的 week/month_label | 给定日期序列 → 正确桶标签与顺序 |
| 题型取数 `collect_by_bucket()` | 复用 load_and_classify，逐题算各桶占比/均分 | load_and_classify | 单选/多选/量表分别产出正确结构 |
| 检验器 `run_tests()` | 按 type 选检验法、算 p、双门槛判 drift | scipy | 已知数据 → 已知 p 与 drift 标记 |
| findings 组装 | 汇总为 drift_findings.json | 上三者 | schema 完整、字段齐全 |
| export 出表 | findings + conclusions → Excel 4 Sheet | openpyxl, _styles | Sheet 数、合并单元格、AI结论列到位 |

每个单元职责单一、接口清晰（JSON/DataFrame），可独立理解与测试。
