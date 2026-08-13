# 异动分析通用化设计

> 日期：2026-08-12
> 范围：将阶段 6 异动诊断从「回流问卷专用」泛化为通用问卷时序对比能力，并打通下载/清洗 → 异动诊断的衔接。

---

## 1. 目标

- **泛化**：去除「回流」硬语义，使异动分析能力可用于满意度跟踪、NPS 月度对比、活动前后对比、版本/渠道对比等多场景。
- **衔接连贯**：下载/清洗产物可直接作为异动分析输入，文档明确指引，无歧义断点。
- **向后兼容**：所有现有命令、MC 月度复用 `value_labels.json` 机制、Excel 4 Sheet 结构均保持不变。

## 2. 范围内改动

| 层 | 文件 | 改动类型 |
|----|------|---------|
| SKILL.md | 阶段 6 标题、description 触发语、路由 B 末尾衔接说明 | 编辑 |
| references/18-drift-workflow.md | 全篇中性化 + 入口衔接 + 粒度扩展 + `--bucket_col` 分桶模式 + 时间列自动检测 | 大改 |
| references/09-survey-download.md | 文末加「后续操作」 | 小改 |
| references/10-survey-clean.md | 文末加「后续操作」 | 小改 |
| scripts/survey_drift.py | 输出名中性化 + 时间列自动检测 + `--bucket_col` 分桶 + `--granularity quarter` + `--custom_ranges` | 中改 |

## 3. 范围外（YAGNI）

- 不做多列交叉分桶（分层检验复杂度过高）
- 不强制 Markdown 报告（用户已选 Excel only）
- 不动阶段 5 报告流程
- 不改下载/清洗脚本本身

---

## 4. 分桶模式双形态

### 模式 A：时间分桶（默认，向后兼容）

```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --granularity week \
  --findings_out "{数据目录}/drift_findings.json"
```

支持的粒度：

| 粒度 | label 格式 | 适用 |
|------|-----------|------|
| `week` | 第N周（M.D-M.D） | 短期监控（现状） |
| `month` | YY年M月 | 月度对比（现状） |
| `day` | YYYY-MM-DD | 日粒度（现状） |
| `quarter`（新）| YY年QX | 季度满意度跟踪 |
| `custom_ranges`（新）| 用户传区间标签 | 活动/版本节点对比 |

`custom_ranges` 示例：
```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "..." \
  --granularity custom_ranges \
  --custom_ranges '[["双11前","2026-10-01","2026-11-10"],["双11期","2026-11-11","2026-11-13"],["双11后","2026-11-14","2026-11-30"]]' \
  --time_col "结束答题时间" \
  --findings_out "..."
```

### 模式 B：列分桶（新增，非时间维度对比）

```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --bucket_col "Q35.用户版本号" \
  --findings_out "..."
```

- `--bucket_col` 指定任意离散列（版本号、活动批次、渠道、用户分层等）
- 桶 = 该列的唯一值，按出现顺序排列
- 用户可用 `--bucket_order` 显式指定桶顺序，如 `["版本A","版本B","版本C"]`
- 模式 B 与模式 A 互斥：传 `--bucket_col` 时忽略 `--granularity` 和 `--time_col`

典型用途：版本 A vs 版本 B、活动前 vs 活动后、渠道对比。

---

## 5. 时间列自动检测（模式 A）

优先级：

1. `--time_col` 显式指定 → 用之
2. 默认列「结束答题时间」存在 → 用之
3. 自动扫描列名含「时间/日期/date/time/提交/答题」关键词的列 → 命中置信度最高的
4. 都没命中 → `need_input` 返回，AI 用 `ask_user_question` 让用户从所有 object/datetime 列里选或指定

### 现状硬编码

`survey_drift.py:1134` 当前：
```python
pa.add_argument("--time_col", default="结束答题时间")
```

改为：
```python
pa.add_argument("--time_col", default=None)  # 默认 None 触发自动检测
```

检测逻辑在 `analyze` 主流程内新增 `detect_time_col(df)` 函数。

---

## 6. 衔接设计（方案 A：轻衔接，只改文档）

### 6.1 现状断点

| # | 位置 | 现状 | 断点 |
|---|------|------|------|
| 1 | SKILL.md 路由 B 末尾 | 「下载成功后自动进入阶段 1」 | 只提阶段 1，不提阶段 6 |
| 2 | SKILL.md 阶段 6 触发条件 | 列触发语 | 没说可承接下载/清洗产物 |
| 3 | 09-survey-download.md | 文末只讲大文件处理 | 没后续操作指引 |
| 4 | 10-survey-clean.md | 文末只讲智能识别逻辑 | 没后续操作指引 |
| 5 | 18-drift-workflow.md 前置 | 「输入为含时间列的量化原始 CSV」 | 没说 CSV 从哪来 |

### 6.2 衔接改动

**改动 1：SKILL.md 阶段 6 前置段加入口说明**

```markdown
数据可来自三条路径：
- 本地 CSV/Excel（路径 A）
- `survey_download.py download` 下载的 `files.quantified_data` 文件路径（路径 B）
- `survey_download.py clean` 清洗后下载的 `files.quantified_data` 文件路径（路径 B+清洗）

下载/清洗产物的 `quantified_data` 路径可直接传给 `survey_drift.py analyze --file_path`。
```

**改动 2：SKILL.md 路由 B 流程末尾加分支提示**

```markdown
下载完成后默认进入阶段 1；如用户明确要求「按周/月对比」「异动诊断」，可直接跳到阶段 6，
用 `quantified_data` 路径作为输入。
```

**改动 3：09-survey-download.md / 10-survey-clean.md 文末各加「后续操作」小节**

```markdown
## 后续操作

- 做异动诊断（按周/月/天/季度/列分桶对比各题异动）→ 读取 `18-drift-workflow.md`，
  用下载的 `files.quantified_data` 文件路径作为 `survey_drift.py analyze --file_path` 输入
- 做基础统计 / 交叉分析 / 文本分析 → 回到 SKILL.md 主流程
```

**改动 4：18-drift-workflow.md 前置改为明确三入口**

```markdown
## 前置

- 输入为**含时间列的量化原始 CSV**（列名为编码后 Q1/Q2…，时间列默认 `结束答题时间`）
- 数据来源三选一：
  - **本地文件**：用户直接给路径（路径 A）
  - **下载产物**：`survey_download.py download` 下载的 `files.quantified_data` 文件路径（路径 B）
  - **清洗后产物**：`survey_download.py clean` 清洗后下载的 `files.quantified_data` 文件路径（路径 B+清洗）
- 粒度（周/月/天/季度/自定义区间）或分桶列由用户指定；未指定时用 `ask_user_question` 让用户选
- 下载产物默认带 `结束答题时间` 列，与异动诊断默认时间列对齐，无需额外配置
- `value_labels.json` 放在数据同目录即可被自动加载，适合 MC 月度等场景复用
```

### 6.3 衔接规则（写入 18-drift-workflow.md）

1. **承接下载产物**：09 返回 JSON 的 `files.quantified_data` 路径直接传给 `--file_path`
2. **承接清洗产物**：10+09 流程产出的 `quantified_data` 同样可用，分桶基于清洗后数据
3. **承接本地文件**：路径 A 直传
4. **时间列承接**：下载的量化数据默认带 `结束答题时间` 列，与异动诊断默认时间列对齐
5. **value_labels.json 同目录复用**：映射文件放在下载输出目录，异动诊断自动加载

---

## 7. 文案中性化清单

| 位置 | 现状 | 改后 |
|------|------|------|
| `survey_drift.py:7` 注释 | 「单份回流问卷按 周/月/天 分桶」 | 「按 周/月/天/季度/自定义区间/任意列 分桶」 |
| `survey_drift.py:1068` 输出名 | `回流异动诊断_{label}_{时间戳}.xlsx` | `问卷异动诊断_{label}_{时间戳}.xlsx` |
| `survey_drift.py:1134` time_col | 默认 `结束答题时间` | 默认 `None`，触发自动检测 |
| SKILL.md 阶段 6 标题 | 「时间异动诊断」 | 「异动诊断」 |
| SKILL.md description 触发语 | 偏回流 | 覆盖满意度/NPS/活动对比/版本对比/渠道对比等 |
| 18-drift-workflow.md 标题 | 「时间异动诊断工作流程（阶段 6）」 | 「问卷异动诊断工作流程（阶段 6）」 |
| 18-drift-workflow.md 触发条件 | 偏回流 | 覆盖通用时序对比场景 |

### SKILL.md description 新增触发语

```yaml
当用户说"按周/月/天诊断这份回流数据的变化"、"逐题对比各周/各月的满意度和 NPS 有没有显著变化"、
  "回流数据有没有异常波动"、"哪个指标这期掉了/涨了、显著吗"、"异动分析"、"回流报告"、
  "满意度月度跟踪"、"NPS 月度对比"、"活动前后对比"、"版本 A vs 版本 B 异动"、"渠道对比"等
  涉及问卷时序对比、异动诊断的场景时，也应触发。
```

（注：「回流报告」保留作为兼容旧习惯的触发语，但不作为主语义。）

---

## 8. Excel 4 Sheet 结构

**保持不变**，只调整标题文案：

| Sheet | 现状标题 | 改后标题 |
|-------|---------|---------|
| 📊 指标总览 | 不变 | 不变 |
| 📈 逐题异动明细 | 不变 | 不变 |
| ⚠️ 异动汇总 | 不变 | 不变 |
| ℹ️ 方法与样本 | 不变 | 不变 |

**新增 Sheet 元信息**：在「方法与样本」Sheet 顶部加一行标注当前模式：
- 时间分桶模式：`分桶方式：时间分桶，粒度=周/月/天/季度/自定义区间`
- 列分桶模式：`分桶方式：列分桶，分桶列=Q35.用户版本号`

---

## 9. `drift_findings.json` 结构扩展

新增字段：

```json
{
  "granularity": "week" | "month" | "day" | "quarter" | "custom_ranges",
  "bucket_mode": "time" | "column",
  "bucket_col": "Q35.用户版本号",     // 仅 column 模式存在
  "custom_ranges": [...],           // 仅 custom_ranges 粒度存在
  "time_col": "结束答题时间",          // time 模式存在，自动检测得到的实际列名
  "time_col_source": "default" | "explicit" | "auto_detect" | "user_specified",
  "buckets": [...],
  "bucket_sizes": {...},
  "low_n_buckets": [...],
  "metrics": [...],
  "questions": [...],
  "nps_col": "...",
  "satisfaction_cols": [...]
}
```

`bucket_mode` 字段让 export 子命令知道当前是时间分桶还是列分桶，便于文案生成。

---

## 10. 向后兼容性验证清单

- [ ] 现有命令 `survey_drift.py analyze --file_path xxx --granularity week` 照常工作
- [ ] 默认时间列 `结束答题时间` 仍被自动检测命中（走 `auto_detect` 路径但结果一致）
- [ ] MC 月度复用的 `value_labels.json` 同目录加载机制保留
- [ ] Excel 4 Sheet 结构不变
- [ ] `--summary-scope latest/all` 行为不变
- [ ] 五点量表题自动纳入 metrics 的逻辑不变
- [ ] 多选基数逻辑（以"答过此题的人数"为基数）不变
- [ ] 双门槛判定（p<0.05 且 Δ≥5pp 或 Δ≥0.1）不变
- [ ] 样本守卫（n<30 不判异动）不变

---

## 11. 测试要点

### 11.1 时间分桶模式回归

- 现有回流数据按周/月/天分桶，输出与改动前一致（除文件名中性化）
- 默认时间列自动检测命中 `结束答题时间`
- 缺时间列时返回 `need_input` + `time_col_missing`

### 11.2 新增粒度

- `--granularity quarter`：按季度分桶，label 形如 `26年Q1`
- `--granularity custom_ranges`：按用户传的区间分桶，label 用用户给的标签

### 11.3 列分桶模式

- `--bucket_col "Q35.用户版本号"`：按版本号分桶
- `--bucket_order` 显式指定顺序
- 模式互斥：传 `--bucket_col` 时 `--granularity` 和 `--time_col` 被忽略

### 11.4 衔接

- 09 下载 JSON 的 `files.quantified_data` 路径直接传给 `--file_path`，能正常分析
- 10 清洗后下载的 `quantified_data` 同样可用
- `value_labels.json` 放在下载目录被自动加载

### 11.5 文案

- 输出文件名为 `问卷异动诊断_xxx.xlsx`
- Excel 4 Sheet 结构不变
- 「方法与样本」Sheet 顶部标注分桶模式

---

## 12. 实施顺序

1. **scripts/survey_drift.py**：输出名中性化、time_col 默认改 None、新增 `detect_time_col`、新增 `--bucket_col`/`--bucket_order`、新增 `--granularity quarter`/`custom_ranges`、findings JSON 扩展字段
2. **references/18-drift-workflow.md**：全篇中性化、前置改三入口、新增模式 B 文档、新增粒度文档
3. **references/09-survey-download.md / 10-survey-clean.md**：文末加「后续操作」
4. **SKILL.md**：阶段 6 标题中性化、前置加三入口说明、路由 B 末尾加分支提示、description 扩展触发语
5. **回归测试**：用现有回流数据验证向后兼容

---

## 13. 风险

| 风险 | 缓解 |
|------|------|
| 时间列自动检测误命中（如问卷里有「答题时长」列） | 关键词清单严格区分：「时间/日期/date/time/提交/答题」+ 类型校验（必须是 datetime 或可解析为时间的字符串）|
| 列分桶模式桶数过多（如版本号有 50 个） | 加 `--max_buckets` 默认 20，超出时返回 `need_input` 让用户聚合或筛选 |
| custom_ranges 日期格式错误 | 脚本校验日期格式，错误时返回明确报错 |
| 旧文档用户用「回流报告」触发语找不到 | description 保留「回流报告」作为兼容触发语 |
