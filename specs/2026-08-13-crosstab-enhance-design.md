# 交叉分析优化设计

> 日期：2026-08-13
> 范围：交叉分析工具的显著性检验、自动识别分组维度、文件命名规范化、得分分析全量表题+样本量行，并确保多分组并列展开默认可用。

---

## 1. 目标

- **需求 1（保障）**：确保现有「多分组并列展开」功能默认成功调用——`--col_questions` 传多个列时，各分组维度并列展开，各自带总计列。
- **需求 2**：新增显著性检验，对比对象为**该分组维度的总计列**（如男/女对比性别总计，不是全样本总计），双门槛（p<0.05 且 |Δ|≥5pp）才标记显著，Excel 着色区分方向。
- **需求 4**：新增自动识别分组维度——`--col_questions '["auto"]'` 时脚本关键词匹配人口学题返回候选清单，AI+LLM 语义判断后用 `ask_user_question` 让用户多选。
- **需求 5**：输出文件命名规范化——`{文件名}_交叉分析_按{简称1}_{简称2}.xlsx`。
- **需求 7**：得分分析自动识别全部 1-5 分量表题（关键词 ∪ 五点量表自动检测），单 Sheet 多量表题，每均分行下紧接样本量行（对齐异动分析「📊 指标总览」）。

## 2. 范围内改动

| 层 | 文件 | 改动类型 |
|----|------|---------|
| 脚本 | `scripts/crosstab.py` | 核心：显著性检验、得分分析改造、文件命名、auto 模式 |
| 脚本 | `scripts/load_and_classify.py` | 新增 `identify_demographic_cols` 人口学题识别 |
| 文档 | `references/12-crosstab-workflow.md` | 更新多分组/自动识别/显著性说明 |
| 文档 | `SKILL.md` | 阶段 3 描述补自动识别触发语 |
| 测试 | `tests/test_crosstab_generalize.py` | 新建，覆盖新能力 |

## 3. 范围外（YAGNI）

- 不做多列交叉分组（笛卡尔积），只做并列展开（已有，保障即可）
- 不做分组间两两对比，只做 vs 该分组维度的总计列
- 不加块状结构（用户明确不要）
- 不动 `run_crosstab` 的频数计算核心逻辑

---

## 4. 需求 1：确保多分组并列展开默认可用（保障）

### 4.1 现状核实

`run_crosstab`（`crosstab.py:187-392`）的 `col_questions` 处理逻辑（L246-311）已支持多列：
- 遍历 `valid_cols`，每个列变量生成各分组值 + 该维度总计列
- 列标签格式：`"{列名}\n{分组值}"`，总计列：`"{列名}\n总计"`
- 各分组维度并列展开，各自独立总计

### 4.2 保障措施

1. **文档明确化**：`12-crosstab-workflow.md` 补充多分组示例，明确「传多个列即并列展开，各自带总计」。
2. **回归测试**：新增测试验证多分组场景（性别+年龄+职业）输出结构正确——每个分组维度都有总计列，列展开顺序正确。
3. **stdout JSON 补字段**：在 `_generate_output_json` 返回里加 `col_dimensions` 字段，列出各分组维度及其总计列标签，便于 AI 和下游理解结构。

```json
{
  "col_dimensions": [
    {"question": "Q33.请问您的性别是？", "values": ["男", "女"], "total_label": "Q33.请问您的性别是？\n总计"},
    {"question": "Q34.请问您的年龄是？", "values": ["<18", "18-24", ...], "total_label": "Q34.请问您的年龄是？\n总计"}
  ]
}
```

---

## 5. 需求 2：显著性检验 vs 分组维度总计

### 5.1 对比对象

**该分组维度的总计列**，不是全样本总计。例：
- 性别维度下：男 vs 性别总计、女 vs 性别总计
- 年龄维度下：18-24 vs 年龄总计、25-30 vs 年龄总计
- 每个分组维度独立对比自己的总计

### 5.2 检验方法

复用异动分析 `two_prop_z`（两比例 z 检验）：
- 对每个分组维度的每个分组值，每个选项占比 vs 该维度总计列对应选项占比
- 返回 `(z, p_value)`
- 双门槛判定：`p < 0.05` 且 `|Δ占比| ≥ 5pp` → 显著

### 5.3 新增函数

```python
def calc_significance(ct_result: dict) -> dict:
    """对每个分组维度的各分组值 vs 该维度总计列，逐选项做两比例 z 检验。
    
    Returns:
        {分组维度列名: {分组值: {选项: {p, delta_pp, significant, direction}}}}
        direction: "up" (分组 > 总计) / "down" (分组 < 总计)
    """
```

逻辑：
1. 从 `ct_result["freq_df"]` 提取每个分组维度的列（通过列标签的 `\n` 分割识别归属）
2. 每个维度找到自己的总计列（列标签含 `\n总计`）
3. 对该维度下每个非总计分组值列，逐行（选项）：
   - `p_group = freq_df[group_col][option] / col_totals[group_col]`
   - `p_total = freq_df[total_col][option] / col_totals[total_col]`
   - `z, p = two_prop_z(freq_df[group_col][option], col_totals[group_col], freq_df[total_col][option], col_totals[total_col])`
   - `delta_pp = (p_group - p_total) * 100`
   - `significant = (p < 0.05) and (abs(delta_pp) >= 5)`
   - `direction = "up" if delta_pp > 0 else "down"`

### 5.4 Excel 着色规则（替换现有 `_apply_diff_heatmap`）

现有 heatmap 逻辑（跨分组 max-min）**替换**为 vs 总计列的显著性着色：

| 条件 | 背景色 | 字体色 | 标记 |
|------|--------|--------|------|
| 显著 + 分组 > 总计 | amber-100 `FEF3C7` | green-800 `1E7D32` 粗体 | ↑ |
| 显著 + 分组 < 总计 | amber-100 `FEF3C7` | red-700 `C0392B` 粗体 | ↓ |
| 非显著 | 无底色（zebra） | slate-600 `475569` | 无 |

总计列单元格不着色（基准列，用 indigo-100 `E0E7FF` 底色标识，同索引列风格）。

### 5.5 `diff_summary` 改造

从「跨分组 max-min」改为「vs 总计列 max delta」：
```python
# 旧：找跨分组的最大差异选项
# 新：对每题，找出与总计列差异最大（且显著）的选项+方向
diff_summary[question] = {
    "max_diff_option": "选项X",
    "max_delta_pp": 8.5,
    "direction": "up",  # 该分组值高于总计
    "significant": True,
    "group": "Q33.性别\n男",  # 哪个分组值偏离总计最多
}
```

### 5.6 stdout JSON 扩展

```json
{
  "significant_matrix": {
    "Q33.请问您的性别是？": {
      "男": {"满意": {"p": 0.003, "delta_pp": 6.2, "significant": true, "direction": "up"}, ...},
      "女": {"满意": {"p": 0.12, "delta_pp": 2.1, "significant": false, "direction": "up"}, ...}
    }
  }
}
```

---

## 6. 需求 4：自动识别分组维度

### 6.1 `load_and_classify.py` 新增

```python
DEMOGRAPHIC_KEYWORDS = ["性别", "年龄", "职业", "付费", "充值", "会员", "渠道", "地区", "城市", "设备"]


def identify_demographic_cols(df, classification):
    """关键词匹配人口学题。返回候选清单（按列顺序）。"""
    candidates = []
    for col in classification.get("single_choice", []):
        if any(kw in col for kw in DEMOGRAPHIC_KEYWORDS):
            candidates.append(col)
    return candidates
```

### 6.2 `crosstab.py` 新增 `--col_questions '["auto"]'`

在 `run_crosstab_pipeline` 里检测：
```python
if col_questions == ["auto"]:
    from load_and_classify import identify_demographic_cols
    candidates = identify_demographic_cols(df, classification)
    if not candidates:
        return {"status": "need_input", "reason": "no_demographic", 
                "message": "未识别到人口学题，请用 --col_questions 指定分组列"}
    return {"status": "need_input", "reason": "col_candidates", 
            "candidates": candidates,
            "message": "识别到以下候选分组维度，请选择"}
```

AI 拿到候选后做 LLM 语义判断（考虑题型/选项数/业务场景），用 `ask_user_question` 让用户多选，再把选中的列名传回 `--col_questions` 重跑。

### 6.3 AI 判断逻辑（写入 12-crosstab-workflow.md）

脚本只做关键词匹配返回候选，**LLM 语义判断由 AI 在 SKILL.md 层做**：
- 候选清单 + 各列的选项数 + 选项文本（从 `load_and_classify` 的 single_choice 推断）
- AI 结合业务场景判断哪些列适合做分组维度
- 用 `ask_user_question`（multiSelect: true）让用户确认

---

## 7. 需求 5：文件命名规范化

### 7.1 新增函数

```python
def _short_col_label(col):
    """从列名提取简短 label：Q33.请问您的性别是？ → 性别；Q35.请问您的职业是？ → 职业。"""
    s = str(col)
    # 去 Q\d+. 前缀
    m = re.match(r"Q\d+\.\s*(.+)", s)
    s = m.group(1) if m else s
    # 提取核心关键词（性别/年龄/职业/付费/会员/渠道/地区/城市/设备）
    for kw in ["性别", "年龄", "职业", "付费", "充值", "会员", "渠道", "地区", "城市", "设备"]:
        if kw in s:
            return kw
    # 无关键词匹配，截断到 8 字
    return s[:8]


def default_output_filename(col_questions, file_path):
    """多分组用 _按{简称1}_{简称2}_{简称3}"""
    short_names = [_short_col_label(c) for c in col_questions]
    base = os.path.splitext(os.path.basename(file_path))[0]
    return f"{base}_交叉分析_按{'_'.join(short_names)}.xlsx"
```

### 7.2 示例

- 单分组：`survey_91044_..._交叉分析_按性别.xlsx`
- 多分组：`survey_91044_..._交叉分析_按性别_年龄_职业.xlsx`
- 非人口学列：`survey_91044_..._交叉分析_按付费等级.xlsx`（关键词匹配）

---

## 8. 需求 7：得分分析全量表题+样本量行

（详细内容见第 10 节，本节为占位避免编号断裂）

---

## 9. 可视化增强（差异可视化全做）

### 9.1 新增第 4 个 Sheet：📊 差异热力图

**目的**：一眼看出哪个分组在哪个选项上偏离总计最多，全盘可视化差异格局。

**结构**：
- 行 = 题目 × 选项（排除总计行）
- 列 = 各分组值（排除总计列，按分组维度分块）
- 单元格内容 = `+8.5pp` / `-6.2pp` / `—`（非显著），格式 `+0.0pp`/`-0.0pp`，size 10 居中

**渐变着色规则**（按 |delta_pp| 深浅，方向决定色相）：

| delta_pp 区间 | 背景色 | 字体色 |
|--------------|--------|--------|
| ≥ +20pp | green-400 `66BB6A` | 白色 |
| +15 ~ +20pp | green-200 `A5D6A7` | green-900 `1B5E20` |
| +10 ~ +15pp | green-100 `C8E6C9` | green-800 `2E7D32` |
| +5 ~ +10pp | green-50 `E8F5E9` | green-700 `388E3C` |
| ≤ -20pp | red-400 `EF5350` | 白色 |
| -20 ~ -15pp | red-200 `EF9A9A` | red-900 `B71C1C` |
| -15 ~ -10pp | red-100 `FFCDD2` | red-800 `C62828` |
| -10 ~ -5pp | red-50 `FFEBEE` | red-700 `D32F2F` |
| \|delta\| < 5pp 或 p≥0.05 | slate-50 `F8FAFC` | slate-400 `94A3B8` + `—` |

**分块视觉**：
- 各分组维度之间用 indigo-100 `E0E7FF` 空列隔开（列宽 2）
- 分组维度标题行（slate-800 `1E293B` 底 + 白字 + size 10 bold）跨列合并标「性别」「年龄」「职业」
- 表头行：分组值（男/女/<18/...），slate-700 `334155` 底 + 白字
- 冻结 `C3`（前两列题名/选项 + 表头两行：维度行+分组值行）

**新增函数**：
```python
def calc_heatmap_data(ct_result, significance_matrix):
    """生成热力图数据。返回 DataFrame，index=题目×选项，columns=各分组值，值=delta_pp。"""

def _format_heatmap_sheet(ws, heatmap_df, col_dimensions):
    """格式化热力图 Sheet：渐变着色 + 分块 + 维度标题行。"""
```

### 9.2 列百分比 Sheet 强化着色

在现有显著性着色（需求 2 第 5.4 节）基础上叠加 **DataBar**：
- 显著且 up：单元格 `FEF3C7` 底 + green-800 `1E7D32` 字 ↑，叠加 indigo `4F46E5` DataBar（条长 = delta_pp/20，封顶 20pp）
- 显著且 down：单元格 `FEF3C7` 底 + red-700 `C0392B` 字 ↓，叠加 indigo DataBar
- 非显著：zebra + slate-600 `475569` 字，无 DataBar
- 总计列：indigo-100 `E0E7FF` 底（基准列标识），无 DataBar，indigo-900 `312E81` 字

**改造**：`_apply_diff_heatmap` → `_apply_significance_heatmap`，在着色基础上加 DataBar 规则。

### 9.3 得分分析 Sheet 趋势条

每个**均分单元格**叠加横向 DataBar（1-5 分刻度，indigo `4F46E5`，min=1 max=5）：
- 均分值居中显示（size 11 bold indigo-700 `4338CA`），DataBar 作背景
- 样本量行：size 9 slate-400 `94A3B8`，`n=#,##0` 格式，无 DataBar
- 总计列均分：indigo-100 `E0E7FF` 底（基准），DataBar 同样显示

**改造**：`_format_score_sheet_v2` 识别均分行（index 含「得分」不含「样本量」）加 DataBar，样本量行（index 含「样本量」）不加。

### 9.4 最终 Excel 结构（4 Sheet）

| Sheet | 内容 | tabColor | 新增视觉 |
|-------|------|----------|---------|
| 交叉分析 | 频数表 | slate-800 `1E293B` | 基础 Slate+Indigo 风格（已做） |
| 列百分比 | 列百分比 + 显著性着色 + DataBar | indigo-600 `4F46E5` | **DataBar 叠加显著性着色** |
| 得分分析 | 均分+样本量行 + 趋势条 | indigo-900 `312E81` | **均分 DataBar + 样本量行** |
| 📊 差异热力图 | delta_pp 渐变热力 | red-700 `C0392B` | **全新 Sheet，渐变热力图** |

---

## 10. 需求 7：得分分析全量表题+样本量行

### 10.1 自动识别 1-5 分量表题

复用异动分析 `survey_drift.py` 的 `_five_point_scale_series` 逻辑（在 crosstab.py 本地实现，避免跨脚本依赖）：

```python
def _five_point_scale_series(series):
    """判断某列是否五点量表（取值均为 1~5 的整数编码）。"""
    non_null = series.dropna()
    if len(non_null) == 0:
        return False
    vals = set()
    for v in non_null.unique():
        try:
            fv = float(v)
        except (ValueError, TypeError):
            return False
        iv = int(fv)
        if fv != iv:
            return False
        vals.add(iv)
    return bool(vals) and vals.issubset({1, 2, 3, 4, 5}) and max(vals) == 5 and len(vals) >= 4
```

### 10.2 改造 `auto_detect_score_questions`

```python
def auto_detect_score_questions(df, ct_result):
    """自动识别可计算得分的题目：关键词（满意/NPS/推荐）∪ 五点量表自动检测。"""
    scoreable = []
    valid_rows = ct_result["valid_rows_map"]
    for q_name in valid_rows:
        if valid_rows[q_name] != "single":
            continue
        # 关键词识别
        if _detect_score_type(q_name, df) is not None:
            scoreable.append(q_name)
            continue
        # 五点量表自动检测
        if q_name in df.columns and _five_point_scale_series(df[q_name]):
            scoreable.append(q_name)
    return scoreable
```

### 10.3 得分 Sheet 结构（单 Sheet 多量表题，对齐异动分析「📊 指标总览」+ 趋势条）

**改造 `calc_scores`**：在现有得分行基础上，为每个量表题**新增样本量行**。

输出 DataFrame 结构：
```
                                    | 性别(总计) | 性别(男) | 性别(女) | 年龄(总计) | ...
Q1.满意度 - 满意度得分(加权均值)    |   4.20    |   4.15   |   4.25   |   4.20    |
Q1.满意度 - 样本量                  |   1500    |   700    |   800    |   1500    |
Q13.整体印象 - 满意度得分(加权均值) |   3.85    |   3.80   |   3.90   |   3.85    |
Q13.整体印象 - 样本量               |   1480    |   690    |   790    |   1480    |
Q4.继续游玩意愿 - 满意度得分(加权均值) | 4.10   |   4.05   |   4.15   |   4.10    |
Q4.继续游玩意愿 - 样本量             | 1495    |   695    |   800    |   1495    |
```

每量表题两行：
- **均分行**：size 11，bold，indigo-700 `4338CA`，格式 `0.00`，居中，叠加 indigo DataBar（1-5 刻度，见 9.3 节）
- **样本量行**：size 9，slate-400 `94A3B8`，格式 `n=#,##0`，居中，行高 18，无 DataBar

样本量 = 该分组下该题的有效作答数（非 NaN）。

### 10.4 `_format_score_sheet_v2` 更新

- 识别均分行 vs 样本量行（通过 index 含「样本量」判断）
- 均分行：size 11 bold indigo-700，行高 22，加 indigo DataBar（min=1 max=5）
- 样本量行：size 9 slate-400，行高 18，无 DataBar
- 总计列均分：indigo-100 底（基准标识）
- 冻结首列 + 首两行（量表题跨多行时表头可见）

---

## 11. 向后兼容性

- 现有 `--col_questions '["Q33.性别"]'` 单列照常工作
- 现有 `--col_questions '["Q33.性别", "Q34.年龄"]'` 多列照常工作（已支持，需求 1 保障）
- 现有 `--calc_scores auto` 行为扩展（自动识别量表题更多，但原有关键词识别的题仍纳入）
- 现有 stdout JSON 字段保留，新增 `col_dimensions`、`significant_matrix` 字段
- 现有 `--col_questions` 不传 `auto` 时行为不变
- **新增第 4 个 Sheet（差异热力图）**：现有调用者需知道 Sheet 数从 2-3 变为 3-4

---

## 12. 测试要点

### 12.1 多分组并列展开回归（需求 1）
- `--col_questions` 传 3 个列，验证每个分组维度都有总计列
- 列展开顺序：分组1各值 + 分组1总计 + 分组2各值 + 分组2总计 + ...
- `col_dimensions` 字段正确列出各维度

### 10.2 显著性检验（需求 2）
- 构造已知差异数据：某分组某选项占比 vs 总计列占比 差 10pp，验证 significant=True
- 构造无差异数据：各分组占比接近总计，验证 significant=False
- 着色规则：up→green+amber，down→red+amber，非显著→无底色
- `diff_summary` 改为 vs 总计列 max delta

### 10.3 自动识别分组维度（需求 4）
- 含性别/年龄/职业列的数据，`--col_questions '["auto"]'` 返回候选清单
- 无人口学题时返回 `no_demographic`

### 10.4 文件命名（需求 5）
- 单分组：`_交叉分析_按性别.xlsx`
- 多分组：`_交叉分析_按性别_年龄_职业.xlsx`

### 10.5 得分分析全量表题+样本量（需求 7）
- 五点量表题（非满意度/NPS 关键词）被自动识别纳入
- 每均分行下有样本量行
- 样本量行格式 `#,##0`，均分行格式 `0.00`
- 样本量 = 该分组该题有效作答数

---

## 13. 实施顺序

1. **load_and_classify.py**：新增 `identify_demographic_cols` + `DEMOGRAPHIC_KEYWORDS`
2. **crosstab.py 核心逻辑**：
   - 新增 `_short_col_label` + `default_output_filename`
   - 新增 `_five_point_scale_series`（本地实现）
   - 改造 `auto_detect_score_questions`（关键词 ∪ 五点量表）
   - 改造 `calc_scores`（每量表题加样本量行）
   - 新增 `calc_significance`（vs 分组维度总计，两比例 z）
   - 改造 `get_crosstab_summary`（diff_summary 改为 vs 总计列）
   - 改造 `_generate_output_json`（加 `col_dimensions` + `significant_matrix`）
   - `run_crosstab_pipeline` 加 `--col_questions '["auto"]'` 处理
3. **crosstab.py 可视化**：
   - 新增 `calc_heatmap_data` + `_format_heatmap_sheet`（第 4 Sheet 差异热力图，渐变着色+分块）
   - `_apply_diff_heatmap` → `_apply_significance_heatmap`（vs 总计列着色 + DataBar 叠加）
   - `_format_score_sheet_v2` 更新（均分行+样本量行区分样式 + 均分 DataBar）
4. **文档**：`12-crosstab-workflow.md` + `SKILL.md`
5. **测试**：`tests/test_crosstab_generalize.py`

---

## 14. 风险

| 风险 | 缓解 |
|------|------|
| 多分组列展开时列数过多（性别2+年龄6+职业7=15+总计3=18列） | 文档提示用户列数控制，Excel 横向滚动可接受 |
| 显著性检验对多选题不适用（多选各子选项非互斥） | 仅对单选题做显著性检验，多选题只算占比不标显著 |
| 五点量表自动检测误判（如某题恰好取值 1-5 但非量表） | `len(vals) >= 4` 门槛（至少出现 4 个不同值）+ 关键词优先 |
| `--col_questions '["auto"]'` 与现有 `"all"` 语义混淆 | auto 只用于 col_questions，row_questions 的 all 语义不变 |
| 样本量行插入后 Score Sheet 行数翻倍 | 可接受，对齐异动分析指标总览的设计 |
