# 时间异动诊断工作流程（阶段 6）

> 📌 **何时读取本文档**：用户有单份含时间列的回流问卷数据，想按周/月/天自动对比、诊断满意度/NPS/单选/多选的异动时，由 SKILL.md 阶段 6 指引跳转至此。

## 触发条件

- "按周/月/天诊断这份回流数据的变化"
- "逐题对比各周/各月/各天的满意度和 NPS 有没有显著变化"
- "回流数据有没有异常波动 / 哪个指标这期掉了/涨了、显著吗"
- MC 等持续回流问卷，想及时发现数据异动

## 前置

- 输入为**含时间列的量化原始 CSV**（列名为编码后 Q1/Q2…，时间列默认 `结束答题时间`）。
- 粒度（周/月/天）由用户指定；未指定时用 `ask_user_question` 让用户三选一。

## 三步编排（一句话触发，中途不停）

### Step A：运行 analyze，分桶 + 逐题检验

```bash
python {SKILL_DIR}/scripts/survey_drift.py analyze \
  --file_path "量化数据.csv" \
  --granularity week \
  --findings_out "{数据目录}/drift_findings.json"
```

可选参数：`--time_col`（默认 `结束答题时间`）、`--nps_col`、`--satisfaction_cols`（不传则按关键词自动识别）、`--min_n`（默认 30）。

读 stdout JSON：
- `status=success` → 记下 `findings_out`、`buckets`、`low_n_buckets`、`questions_with_drift`，进入 Step B。
- `status=need_input` → 按 `reason` 处理：`time_col_missing` 用 `ask_user_question` 让用户确认时间列后加 `--time_col` 重跑；`no_metric` 让用户指定 `--nps_col` / `--satisfaction_cols` 重跑。
- 若 `buckets` 少于 2 个 → 提示"当前时间跨度不足以分期，请扩大数据范围或改用更细的粒度"。

### Step B：逐题写一句话结论（你来做）

读 `drift_findings.json`，为 `metrics` 和 `questions` 里**每一项**写一句话结论，输出到 `conclusions.json`（`{题目或指标名: 结论文本}` 映射；键用 question 的 `question` 字段、metric 的 `source_col` 字段）。

结论规范：
- **只写定性判断**，数字（占比/均分/p）脚本已算好，不要手写数字避免矛盾；可引用方向和量级。
- **异动题**（`drift=true`）：写"哪个选项/指标 + 方向（升/降）+ 幅度（pp 或分）+ 是否显著"，如"『非常满意』占比环比下降约 6pp（显著），需排查最新版本体验"。
- **无异动题**：写"本期无显著变化"。
- **样本不足**（`low_n=true`）：结论末尾加"（样本不足，仅供参考）"。

### Step C：运行 export，生成 Excel

```bash
python {SKILL_DIR}/scripts/survey_drift.py export \
  --findings "{数据目录}/drift_findings.json" \
  --conclusions "{数据目录}/conclusions.json"
```

默认输出 `回流异动诊断_{按周/按月/按天}_{时间戳}.xlsx`，保存在 findings 同目录；可用 `--output_path` 指定。

**`--value-labels`（编码→标签映射，可选）**：人口题等以数字编码存储的题目（如性别=1/2/3），可提供一个 JSON 把编码替换成可读标签。
- 格式：`{"Q33.请问您的性别是？": {"1": "男", "2": "女", "3": "不愿意透露"}, ...}`（键=题目 `question`，值=编码→标签）。
- **缺省自动探测**：不传 `--value-labels` 时，自动加载 findings **同目录下的 `value_labels.json`**（若存在）。MC 回流研究每月复用同一份映射，放在数据目录即可"一次配置、后续自动生效"。
- 效果：明细表选项列显示标签而非编码；有映射的题目按**编码数字升序**排列（保持问卷逻辑顺序，不按占比重排）。
- **人口题（性别/年龄/职业）额外补一版归一化**：若题目含「不愿意透露」选项，会在样本量行下方追加一个「剔除「不愿意透露」后归一化」子区块——对其余选项的占比在剔除后重新归一（各期 + 整体），并给出剔除「不愿意透露」后的样本量。

**`--summary-scope`（异动汇总范围，默认 `latest`）**：
- `latest`：⚠️ 异动汇总只列**最新相邻期**（最新一桶 vs 上一桶）的异动。适合"本期有没有掉"的日常监控。
- `all`：⚠️ 异动汇总列出**全时间线任意相邻期**的历史异动，并新增「时段」列标注异动发生在哪一桶（如 `第28周（7.6-7.12）`）。适合多桶（如按周 9 桶）回看拐点、定位异动集中在哪一期。
- 桶数 ≥ 3 且想回溯历史拐点时建议加 `--summary-scope all`：
  ```bash
  python {SKILL_DIR}/scripts/survey_drift.py export --findings "..." --conclusions "..." --summary-scope all
  ```

## drift_findings.json 结构

- 顶层：`granularity`、`time_col`、`buckets`（旧→新有序）、`bucket_sizes`、`low_n_buckets`、`metrics`、`questions`、`nps_col`、`satisfaction_cols`。
- `metrics[]`：满意度均分（`type=satisfaction_mean`）、NPS（`type=nps`）。含 `by_bucket` 各期值、`adjacent` 相邻期检验（`delta`/`delta_pp`、`test`、`p`、`significant`、`drift`、`low_n`、`direction`）。**凡取值为 1~5 的五点量表单选题都会自动纳入 `metrics` 做均分显著性检验**（与 `satisfaction_cols` 关键词识别结果合并去重；`findings.satisfaction_cols` 记录最终纳入的全部均分题）。
- `questions[]`：单选（含 `overall_test` 卡方）、多选（`overall_test=null`）。含 `by_bucket` 各期各选项占比、`adjacent_option_tests`（逐选项相邻期两比例 z）、题级 `drift`、`low_n`。**多选题占比/样本量以"答过此题(至少勾选一项)的人数"为基数**（与交叉分析一致，逻辑门控题不计未触达者；如 Q25 基数≈5.6万而非全样本 29万）。
  - `question`：主键（单选=完整列名；多选=`Q\d+.` 根前缀）。**结论 conclusions.json 的键用 `question`**。
  - `question_label`：Excel 展示用完整题干（多选从子列还原冒号前部分，单选同 `question`）。写结论时可参考它理解题意，但键仍用 `question`。

conclusions.json 示例：
```json
{
  "Q7.活动评价（单选）": "『满意』占比环比下降约 6pp（显著），需排查最新活动体验。",
  "Q1.整体满意度": "均分环比小幅回升，无显著变化。"
}
```

## Excel 4 Sheet

| Sheet | 内容 |
|-------|------|
| 📊 指标总览 | 满意度均分、NPS 全时间线各期一列 + 最新期 vs 上期趋势标 ▲▼ + 是否显著。**所有五点量表题（选项1~5）自动纳入并做相邻期均分显著性检验**（t 检验/Mann-Whitney），无需手动指定 `--satisfaction_cols` |
| 📈 逐题异动明细 | 每题×各期占比（DataBar）+ **C 列「整体」基线**（全样本各期按样本量加权的整体占比，供各期对照）+ **单选/多选选项按整体占比降序排**（五点量表题、人口题「性别/年龄/职业」保持原顺序；有编码→标签映射的题按编码升序）+ **逐周环比热力标注**（某周相对前一周显著变化则着色：琥珀底加粗=大幅异动、红/绿字=一般显著升降、灰字=无变化）+ 异动周列 + **每题末行「样本量」**（整体 N + 各期有效 n）+ **五点量表题再加「加权满意度」行**（Σ分值×人数/总样本量 = 1~5 均分，整体 + 各期）+ **人口题（性别/年龄/职业）再加「剔除「不愿意透露」后归一化」子区块**（剩余选项占比重新归一 + 剔除后样本量）+ **AI 结论列** |
| ⚠️ 异动汇总 | 被判异动的题目/指标：时段/变化项/方向/幅度/显著性/AI 结论。默认只列最新相邻期；`--summary-scope all` 列全时间线历史异动并含「时段」列 |
| ℹ️ 方法与样本 | 分桶方式、各桶样本量、检验方法、双门槛阈值、样本不足清单、免责 |

## 统计方法与判定

- 满意度均分 → t 检验（小样本/偏态转 Mann-Whitney U）；NPS/占比 → 两比例 z 检验；单选整体 → 卡方。
- **判异动双门槛**：`p<0.05` 且（占比 Δ≥±5pp 或 均分 Δ≥±0.1）。
- **样本守卫**：桶内有效样本 `n<30` 不判异动，标"样本不足，仅供参考"。

## 错误处理

| 情况 | 处理 |
|------|------|
| `need_input`（缺时间列） | 追问用户时间列，加 `--time_col` 重跑 |
| `need_input`（识别不到 NPS/满意度） | 让用户指定 `--nps_col` / `--satisfaction_cols` 重跑 |
| 桶数 < 2 | 提示时间跨度不足，扩大数据范围或换更细粒度 |

## 后续操作提示

分析完成后可提示用户：换粒度重跑（周↔月↔天）、补充文本"新增反馈"检测（下一阶段能力）、下载最新回流数据再诊断。
