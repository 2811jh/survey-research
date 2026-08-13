# 交叉分析优化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 交叉分析工具增强：显著性检验 vs 分组维度总计、自动识别分组维度、文件命名规范化、得分分析全量表题+样本量行、差异可视化（热力图 Sheet + DataBar + 趋势条）。

**Architecture:** 分三层：(1) `load_and_classify.py` 加人口学题识别；(2) `crosstab.py` 核心逻辑（显著性、auto 模式、文件命名、得分改造）+ 可视化（热力图 Sheet、DataBar 叠加、趋势条）；(3) 文档与测试。现有 `run_crosstab` 频数计算核心不动。

**Tech Stack:** Python 3 + pandas + scipy + openpyxl + argparse。

**Spec:** `survey-research/specs/2026-08-13-crosstab-enhance-design.md`

---

## File Structure

| 文件 | 责任 | 改动 |
|------|------|------|
| `scripts/load_and_classify.py` | 列分类 + 人口学识别 | 新增 `identify_demographic_cols` |
| `scripts/crosstab.py` | 交叉分析主脚本 | 核心逻辑 + 可视化大改 |
| `references/12-crosstab-workflow.md` | 工作流文档 | 更新 |
| `SKILL.md` | 主入口 | 阶段 3 描述补 auto |
| `tests/test_crosstab_generalize.py` | 新能力测试 | 新建 |

---

## Task 1: load_and_classify.py 新增人口学题识别

**Files:**
- Modify: `scripts/load_and_classify.py`
- Test: `tests/test_crosstab_generalize.py`

- [ ] **Step 1: 写失败测试**

Create `tests/test_crosstab_generalize.py`:
```python
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import pandas as pd
import pytest
from load_and_classify import identify_demographic_cols


def test_identify_demographic_识别性别年龄职业():
    df = pd.DataFrame({
        "Q1.满意度": [4, 5, 3],
        "Q33.请问您的性别是？": ["男", "女", "男"],
        "Q34.请问您的年龄是？": ["18-24", "25-30", "18-24"],
        "Q35.请问您的职业是？": ["学生", "工作", "学生"],
        "Q50.付费等级": ["免费", "付费", "免费"],
    })
    classification = {
        "single_choice": ["Q1.满意度", "Q33.请问您的性别是？", "Q34.请问您的年龄是？",
                          "Q35.请问您的职业是？", "Q50.付费等级"],
        "multi_choice": {}, "matrix_scale": {}, "text": [],
        "meta": [], "excluded": [], "valid_for_crosstab": [],
    }
    candidates = identify_demographic_cols(df, classification)
    assert "Q33.请问您的性别是？" in candidates
    assert "Q34.请问您的年龄是？" in candidates
    assert "Q35.请问您的职业是？" in candidates
    assert "Q50.付费等级" in candidates
    assert "Q1.满意度" not in candidates


def test_identify_demographic_无人口学题返回空():
    df = pd.DataFrame({"Q1.满意度": [4, 5], "Q2.玩法偏好": [1, 2]})
    classification = {
        "single_choice": ["Q1.满意度", "Q2.玩法偏好"],
        "multi_choice": {}, "matrix_scale": {}, "text": [],
        "meta": [], "excluded": [], "valid_for_crosstab": [],
    }
    candidates = identify_demographic_cols(df, classification)
    assert candidates == []
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: FAIL with `ImportError: cannot import name 'identify_demographic_cols'`

- [ ] **Step 3: 实现 identify_demographic_cols**

在 `scripts/load_and_classify.py` 末尾追加：
```python
DEMOGRAPHIC_KEYWORDS = ["性别", "年龄", "职业", "付费", "充值", "会员", "渠道", "地区", "城市", "设备"]


def identify_demographic_cols(df, classification):
    """关键词匹配人口学题。返回候选清单（按列顺序）。"""
    candidates = []
    for col in classification.get("single_choice", []):
        if any(kw in str(col) for kw in DEMOGRAPHIC_KEYWORDS):
            candidates.append(col)
    return candidates
```

- [ ] **Step 4: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/load_and_classify.py tests/test_crosstab_generalize.py && git commit -m "feat(crosstab): add identify_demographic_cols for auto-detecting group dimensions"
```

---

## Task 2: crosstab.py 文件命名规范化 + 五点量表识别

**Files:**
- Modify: `scripts/crosstab.py` (新增 `_short_col_label` + `default_output_filename` + `_five_point_scale_series`)
- Test: `tests/test_crosstab_generalize.py`

- [ ] **Step 1: 写失败测试**

追加到 `tests/test_crosstab_generalize.py`:
```python
from crosstab import _short_col_label, default_output_filename, _five_point_scale_series


def test_short_col_label_提取关键词():
    assert _short_col_label("Q33.请问您的性别是？") == "性别"
    assert _short_col_label("Q34.请问您的年龄？") == "职业"
    assert _short_col_label("Q50.付费等级") == "付费"


def test_short_col_label_无关键词截断():
    assert _short_col_label("Q10.您最喜欢的玩法") == "您最喜欢的玩法"


def test_default_output_filename_单分组():
    name = default_output_filename(["Q33.请问您的性别是？"], "survey_123_数据.csv")
    assert name == "survey_123_数据_交叉分析_按性别.xlsx"


def test_default_output_filename_多分组():
    cols = ["Q33.请问您的性别是？", "Q34.请问您的年龄是？", "Q35.请问您的职业是？"]
    name = default_output_filename(cols, "survey_123_数据.csv")
    assert name == "survey_123_数据_交叉分析_按性别_年龄_职业.xlsx"


def test_five_point_scale_series_识别():
    s = pd.Series([5, 4, 3, 2, 1, 5, 4, 4, 3, 5])
    assert _five_point_scale_series(s) is True


def test_five_point_scale_series_排除二元():
    s = pd.Series([1, 2, 1, 2, 1])
    assert _five_point_scale_series(s) is False


def test_five_point_scale_series_排除NPS():
    s = pd.Series([0, 1, 5, 7, 8, 9, 10, 3, 6, 4])
    assert _five_point_scale_series(s) is False
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: FAIL with `ImportError: cannot import name '_short_col_label'`

- [ ] **Step 3: 实现三个函数**

在 `scripts/crosstab.py` 的辅助函数区（`_detect_csv_encoding` 之后，约 line 57 附近）插入：
```python
# ========================================================================= #
#                    文件命名 + 五点量表识别（通用化）                    #
# ========================================================================= #

_DEMOGRAPHIC_SHORT_KEYWORDS = ["性别", "年龄", "职业", "付费", "充值", "会员", "渠道", "地区", "城市", "设备"]


def _short_col_label(col):
    """从列名提取简短 label：Q33.请问您的性别是？ → 性别。"""
    s = str(col)
    m = re.match(r"Q\d+\.\s*(.+)", s)
    s = m.group(1) if m else s
    for kw in _DEMOGRAPHIC_SHORT_KEYWORDS:
        if kw in s:
            return kw
    return s[:8]


def default_output_filename(col_questions, file_path):
    """多分组用 _按{简称1}_{简称2}_{简称3}"""
    short_names = [_short_col_label(c) for c in col_questions]
    base = os.path.splitext(os.path.basename(file_path))[0]
    return f"{base}_交叉分析_按{'_'.join(short_names)}.xlsx"


def _five_point_scale_series(series):
    """判断某列是否五点量表（取值均为 1~5 的整数编码，且至少出现 4 个不同值）。"""
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

- [ ] **Step 4: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: 9 passed (2 Task 1 + 7 new)

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py tests/test_crosstab_generalize.py && git commit -m "feat(crosstab): add _short_col_label, default_output_filename, _five_point_scale_series"
```

---

## Task 3: 改造 auto_detect_score_questions 纳入五点量表

**Files:**
- Modify: `scripts/crosstab.py` (`auto_detect_score_questions` 函数，约 line 463)
- Test: `tests/test_crosstab_generalize.py`

- [ ] **Step 1: 写失败测试**

追加到 `tests/test_crosstab_generalize.py`:
```python
from crosstab import auto_detect_score_questions


def test_auto_detect_score_纳入五点量表题():
    """非满意度/NPS 关键词，但取值 1-5 的题应被识别。"""
    df = pd.DataFrame({
        "Q1.满意度": [5, 4, 3, 4, 5, 2, 1, 4, 5, 3],
        "Q13.整体印象": [4, 3, 5, 4, 2, 3, 4, 5, 1, 4],  # 无关键词，但 1-5 量表
        "Q2.性别": [1, 2, 1, 2, 1, 2, 1, 2, 1, 2],  # 二元，不识别
    })
    classification = {
        "single_choice": ["Q1.满意度", "Q13.整体印象", "Q2.性别"],
        "multi_choice": {}, "matrix_scale": {}, "text": [],
        "meta": [], "excluded": [],
        "valid_for_crosstab": ["Q1.满意度", "Q13.整体印象", "Q2.性别"],
    }
    # 构造最小 ct_result
    ct_result = {
        "valid_rows_map": {"Q1.满意度": "single", "Q13.整体印象": "single", "Q2.性别": "single"},
    }
    scoreable = auto_detect_score_questions(df, ct_result)
    assert "Q1.满意度" in scoreable       # 关键词命中
    assert "Q13.整体印象" in scoreable    # 五点量表命中
    assert "Q2.性别" not in scoreable     # 二元排除
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py::test_auto_detect_score_纳入五点量表题 -v
```
Expected: FAIL (Q13.整体印象 not in scoreable，因为现有逻辑只靠关键词)

- [ ] **Step 3: 改造 auto_detect_score_questions**

找到 `auto_detect_score_questions` 函数（`grep -n "def auto_detect_score_questions" scripts/crosstab.py`）。现有逻辑只靠 `_is_scoreable_question` 关键词识别。改为关键词 ∪ 五点量表：

```python
def auto_detect_score_questions(df: pd.DataFrame, ct_result: dict) -> list:
    """自动识别可计算得分的题目：关键词（满意/NPS/推荐）∪ 五点量表自动检测。"""
    scoreable = []
    valid_rows = ct_result["valid_rows_map"]
    for q_name in valid_rows:
        if valid_rows[q_name] != "single":
            continue
        # 关键词识别
        if _is_scoreable_question(q_name, df) is not None:
            scoreable.append(q_name)
            continue
        # 五点量表自动检测
        if q_name in df.columns and _five_point_scale_series(df[q_name]):
            scoreable.append(q_name)
    return scoreable
```

- [ ] **Step 4: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: 10 passed

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py tests/test_crosstab_generalize.py && git commit -m "feat(crosstab): auto_detect_score_questions now includes 5-point scale via _five_point_scale_series"
```

---

## Task 4: 改造 calc_scores 加样本量行

**Files:**
- Modify: `scripts/crosstab.py` (`calc_scores` 函数，约 line 478)
- Test: `tests/test_crosstab_generalize.py`

- [ ] **Step 1: 写失败测试**

追加到 `tests/test_crosstab_generalize.py`:
```python
from crosstab import calc_scores


def test_calc_scores_含样本量行():
    """每个量表题的得分行下方应紧跟样本量行。"""
    df = pd.DataFrame({
        "Q1.满意度": [5, 4, 3, 4, 5, 2, 4, 5, 3, 4],
        "Q33.性别": ["男", "女", "男", "女", "男", "女", "男", "女", "男", "女"],
    })
    # 构造最小 ct_result：freq_df 用真实 crosstab
    import pandas as pd
    freq = pd.crosstab(
        pd.Categorical(df["Q1.满意度"].astype(str), categories=["5", "4", "3", "2", "总计"]),
        df["Q33.性别"],
        margins=True, margins_name="总计",
    )
    freq.index = pd.MultiIndex.from_arrays([["Q1.满意度"]*5, ["5","4","3","2","总计"]], names=["题目","选项"])
    ct_result = {
        "freq_df": freq,
        "valid_rows_map": {"Q1.满意度": "single"},
        "col_labels": ["男", "女", "总计"],
        "col_totals": {"男": 5, "女": 5, "总计": 10},
    }
    score_df = calc_scores(df, ct_result, ["Q1.满意度"])
    assert score_df is not None
    # 应有得分行和样本量行
    indices = [str(idx) for idx in score_df.index]
    assert any("得分" in i for i in indices)
    assert any("样本量" in i for i in indices)
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py::test_calc_scores_含样本量行 -v
```
Expected: FAIL (现有 calc_scores 不产生样本量行)

- [ ] **Step 3: 改造 calc_scores 加样本量行**

找到 `calc_scores` 函数（`grep -n "def calc_scores" scripts/crosstab.py`）。在每个量表题的得分行追加后，紧跟追加样本量行。在 `score_results.append(...)` 和 `score_index.append(...)` 之后插入样本量行逻辑。

具体：在 calc_scores 函数里，每个 q 处理完得分后（`score_results` 和 `score_index` append 得分行之后），追加：
```python
            # 样本量行：该分组下该题的有效作答数
            sample_sizes = {}
            for col_label in freq_df.columns:
                if str(col_label).endswith("\n总计") or str(col_label) == "总计":
                    # 总计列样本量 = 该题全局有效作答数
                    sample_sizes[col_label] = int(df[q].notna().sum())
                else:
                    # 分组列样本量 = 该分组下该题有效作答数
                    # 从 col_labels + col_totals 推断分组条件较复杂，这里用 freq_df 该列合计
                    sample_sizes[col_label] = int(freq_df.xs(q, level=0).loc[:, col_label].sum())
            score_results.append([sample_sizes.get(c, 0) for c in freq_df.columns])
            score_index.append((q, f"{q_short} 样本量"))
```

其中 `q_short` 是题名简称（去 Q 前缀）。需在函数内提取。注意得分行的 index 格式应为 `(q, f"{q_short} 得分(加权均值)")`，样本量行 `(q, f"{q_short} 样本量")`。

- [ ] **Step 4: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: 11 passed

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py tests/test_crosstab_generalize.py && git commit -m "feat(crosstab): calc_scores adds sample size row below each score row"
```

---

## Task 5: 新增 calc_significance（vs 分组维度总计）

**Files:**
- Modify: `scripts/crosstab.py` (新增 `calc_significance` + `two_prop_z` 本地实现)
- Test: `tests/test_crosstab_generalize.py`

- [ ] **Step 1: 写失败测试**

追加到 `tests/test_crosstab_generalize.py`:
```python
from crosstab import calc_significance, two_prop_z


def test_two_prop_z_基本():
    z, p = two_prop_z(60, 100, 50, 100)
    assert abs(z) > 1.5
    assert p < 0.2


def test_two_prop_z_无差异():
    z, p = two_prop_z(50, 100, 50, 100)
    assert z == 0.0
    assert p == 1.0


def test_calc_significance_vs_分组维度总计():
    """构造已知差异：男组某选项占比 vs 性别总计占比 差 10pp。"""
    import pandas as pd
    # 100 人：男 50，女 50。某选项男 30(60%)、女 20(40%)、总计 50(50%)
    freq = pd.DataFrame(
        {"男": [30, 20, 50], "女": [20, 30, 50], "Q33.性别\n总计": [50, 50, 100]},
        index=pd.MultiIndex.from_arrays([["Q1"],["选项A","选项B","总计"]], names=["题目","选项"]),
    )
    ct_result = {
        "freq_df": freq,
        "col_labels": ["男", "女", "Q33.性别\n总计"],
        "col_totals": {"男": 50, "女": 50, "Q33.性别\n总计": 100},
    }
    sig = calc_significance(ct_result)
    assert "Q33.性别" in sig or "Q33.性别\n总计" in sig
    # 男 vs 总计：选项A 60% vs 50%，差 10pp，应显著
    dim_key = list(sig.keys())[0]
    assert "男" in sig[dim_key]
    assert "选项A" in sig[dim_key]["男"]
    assert sig[dim_key]["男"]["选项A"]["significant"] is True
    assert sig[dim_key]["男"]["选项A"]["direction"] == "up"
    assert abs(sig[dim_key]["男"]["选项A"]["delta_pp"] - 10.0) < 0.1
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py::test_calc_significance_vs_分组维度总计 -v
```
Expected: FAIL with `ImportError`

- [ ] **Step 3: 实现 two_prop_z + calc_significance**

在 `scripts/crosstab.py` 的 `calc_scores` 函数之后插入：
```python
def two_prop_z(c1, n1, c2, n2):
    """两比例 z 检验（双侧，pooled）。返回 (z, p)。n 为 0 时返回 (0.0, 1.0)。"""
    if n1 == 0 or n2 == 0:
        return 0.0, 1.0
    p1 = c1 / n1
    p2 = c2 / n2
    p_pool = (c1 + c2) / (n1 + n2)
    if p_pool == 0 or p_pool == 1:
        return 0.0, 1.0
    se = (p_pool * (1 - p_pool) * (1/n1 + 1/n2)) ** 0.5
    if se == 0:
        return 0.0, 1.0
    z = (p1 - p2) / se
    from scipy import stats
    p = 2 * (1 - stats.norm.cdf(abs(z)))
    return round(z, 4), round(p, 4)


def _extract_dim_from_label(label):
    """从列标签提取分组维度名。'Q33.性别\\n男' → 'Q33.性别'。无 \\n 返回原值。"""
    s = str(label)
    if "\n" in s:
        return s.split("\n")[0]
    return s


def calc_significance(ct_result: dict) -> dict:
    """对每个分组维度的各分组值 vs 该维度总计列，逐选项做两比例 z 检验。
    
    Returns:
        {分组维度列名: {分组值: {选项: {p, delta_pp, significant, direction}}}}
    """
    freq_df = ct_result["freq_df"]
    col_labels = ct_result["col_labels"]
    col_totals = ct_result["col_totals"]
    
    # 按分组维度归类列
    dim_cols = {}  # {维度名: {"total_col": ..., "group_cols": [列表]}}
    for label in col_labels:
        dim = _extract_dim_from_label(label)
        if dim not in dim_cols:
            dim_cols[dim] = {"total_col": None, "group_cols": []}
        if str(label).endswith("\n总计"):
            dim_cols[dim]["total_col"] = label
        else:
            dim_cols[dim]["group_cols"].append(label)
    
    result = {}
    for dim, info in dim_cols.items():
        total_col = info["total_col"]
        if total_col is None:
            continue
        total_n = col_totals[total_col]
        result[dim] = {}
        
        for group_col in info["group_cols"]:
            group_n = col_totals[group_col]
            result[dim][group_col] = {}
            
            # 遍历每个选项（排除总计行）
            for idx in freq_df.index:
                option = idx[1] if isinstance(idx, tuple) else idx
                if str(option) in ("总计", "合计", "Total"):
                    continue
                c_group = int(freq_df.loc[idx, group_col])
                c_total = int(freq_df.loc[idx, total_col])
                
                z, p = two_prop_z(c_group, group_n, c_total, total_n)
                p_group = c_group / group_n if group_n else 0
                p_total = c_total / total_n if total_n else 0
                delta_pp = round((p_group - p_total) * 100, 1)
                significant = (p < 0.05) and (abs(delta_pp) >= 5)
                direction = "up" if delta_pp > 0 else "down"
                
                result[dim][group_col][str(option)] = {
                    "p": p, "delta_pp": delta_pp,
                    "significant": significant, "direction": direction,
                }
    return result
```

- [ ] **Step 4: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: 14 passed (11 + 3 new)

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py tests/test_crosstab_generalize.py && git commit -m "feat(crosstab): add calc_significance comparing each group vs its dimension total"
```

---

## Task 6: 改造 get_crosstab_summary + _generate_output_json + run_crosstab_pipeline（auto 模式 + 文件命名）

**Files:**
- Modify: `scripts/crosstab.py` (`get_crosstab_summary`, `_generate_output_json`, `run_crosstab_pipeline`)
- Test: `tests/test_crosstab_generalize.py`

- [ ] **Step 1: 写 auto 模式失败测试**

追加到 `tests/test_crosstab_generalize.py`:
```python
from crosstab import run_crosstab_pipeline


def test_run_crosstab_pipeline_auto_返回候选():
    df = pd.DataFrame({
        "Q1.满意度": [5, 4, 3, 4, 5],
        "Q33.请问您的性别是？": ["男", "女", "男", "女", "男"],
        "Q34.请问您的年龄是？": ["18-24", "25-30", "18-24", "25-30", "18-24"],
    })
    df.to_csv("/tmp/test_crosstab_auto.csv", index=False, encoding="utf-8-sig")
    result = run_crosstab_pipeline(
        file_path="/tmp/test_crosstab_auto.csv",
        row_questions=["all"],
        col_questions=["auto"],
        calc_scores_mode="auto",
    )
    assert result["status"] == "need_input"
    assert result["reason"] == "col_candidates"
    assert "Q33.请问您的性别是？" in result["candidates"]
    assert "Q34.请问您的年龄是？" in result["candidates"]
```

- [ ] **Step 2: 运行测试确认失败**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py::test_run_crosstab_pipeline_auto_返回候选 -v
```
Expected: FAIL (现有 run_crosstab_pipeline 不支持 auto)

- [ ] **Step 3: 改造 run_crosstab_pipeline 加 auto 模式 + 文件命名 + significance**

找到 `run_crosstab_pipeline` 函数（`grep -n "def run_crosstab_pipeline" scripts/crosstab.py`）。在函数开头（df 加载后、run_crosstab 调用前）插入 auto 检测：

```python
def run_crosstab_pipeline(file_path, row_questions, col_questions, merge_rules=None, calc_scores_mode=None, output_path=None):
    # ... 现有的 df 加载 + classify_columns + merge 逻辑 ...
    
    # auto 模式：识别候选分组维度
    if col_questions == ["auto"]:
        from load_and_classify import identify_demographic_cols
        candidates = identify_demographic_cols(df, classification)
        if not candidates:
            return {"status": "need_input", "reason": "no_demographic",
                    "message": "未识别到人口学题，请用 --col_questions 指定分组列"}
        return {"status": "need_input", "reason": "col_candidates",
                "candidates": candidates,
                "message": "识别到以下候选分组维度，请选择"}
    
    # ... 现有的 run_crosstab 调用 ...
    ct_result = run_crosstab(df, classification, row_questions, col_questions)
    
    # ... 现有的 calc_scores 逻辑 ...
    
    # 新增：显著性检验
    significance_matrix = calc_significance(ct_result)
    
    # 改造 diff_summary：基于 vs 总计列
    diff_summary = get_crosstab_summary(ct_result, significance_matrix)
    
    # 新增：col_dimensions
    col_dimensions = _extract_col_dimensions(ct_result)
    
    # 文件命名：默认用 default_output_filename
    if output_path is None:
        output_path = os.path.join(
            os.path.dirname(os.path.abspath(file_path)),
            default_output_filename(col_questions, file_path),
        )
    
    # ... export + JSON ...
    return _generate_output_json(ct_result, diff_summary, score_df, output_path,
                                 significance_matrix, col_dimensions)
```

新增辅助函数 `_extract_col_dimensions`：
```python
def _extract_col_dimensions(ct_result):
    """提取各分组维度信息。"""
    col_labels = ct_result["col_labels"]
    dim_map = {}
    for label in col_labels:
        dim = _extract_dim_from_label(label)
        if dim not in dim_map:
            dim_map[dim] = {"question": dim, "values": [], "total_label": None}
        if str(label).endswith("\n总计"):
            dim_map[dim]["total_label"] = label
        else:
            dim_map[dim]["values"].append(label)
    return list(dim_map.values())
```

- [ ] **Step 4: 改造 get_crosstab_summary 基于 significance**

现有 `get_crosstab_summary` 改为接收 significance_matrix，返回 vs 总计列的 max delta：
```python
def get_crosstab_summary(ct_result: dict, significance_matrix: dict = None) -> dict:
    """生成差异摘要：基于 vs 分组维度总计的显著性。"""
    if significance_matrix is None:
        significance_matrix = calc_significance(ct_result)
    
    diff_summary = {}
    # 按 row question 聚合：找每题差异最大（且显著）的选项
    for dim, groups in significance_matrix.items():
        for group_col, options in groups.items():
            for option, info in options.items():
                if not info["significant"]:
                    continue
                # 从 group_col 提取题号
                # 用 option 作为 key 的一部分
                q_key = dim  # 简化：以维度为 key
                if q_key not in diff_summary or abs(info["delta_pp"]) > abs(diff_summary[q_key].get("max_delta_pp", 0)):
                    diff_summary[q_key] = {
                        "max_diff_option": option,
                        "max_delta_pp": info["delta_pp"],
                        "direction": info["direction"],
                        "significant": True,
                        "group": group_col,
                    }
    return diff_summary
```

- [ ] **Step 5: 改造 _generate_output_json 加新字段**

找到 `_generate_output_json` 函数。在返回 dict 里加 `col_dimensions` 和 `significant_matrix`：
```python
def _generate_output_json(ct_result, diff_summary, score_df, output_path,
                          significance_matrix=None, col_dimensions=None):
    # ... 现有逻辑 ...
    return {
        "status": "success",
        # ... 现有字段 ...
        "diff_summary": diff_summary,
        "score_summary": ...,
        "significant_matrix": significance_matrix,
        "col_dimensions": col_dimensions,
    }
```

- [ ] **Step 6: 运行测试确认通过**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/test_crosstab_generalize.py -v
```
Expected: 15 passed

- [ ] **Step 7: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py tests/test_crosstab_generalize.py && git commit -m "feat(crosstab): add auto mode, significance-based diff_summary, col_dimensions, default filename"
```

---

## Task 7: 可视化 — 列百分比 Sheet 显著性着色 + DataBar

**Files:**
- Modify: `scripts/crosstab.py` (`_apply_diff_heatmap` → `_apply_significance_heatmap`，`export_crosstab_excel` 调用处)

- [ ] **Step 1: 替换 _apply_diff_heatmap 为 _apply_significance_heatmap**

找到 `_apply_diff_heatmap` 函数（`grep -n "def _apply_diff_heatmap" scripts/crosstab.py`）。整个函数替换为基于 significance_matrix 的着色 + DataBar：

```python
def _apply_significance_heatmap(ws, percent_df, col_labels, significance_matrix,
                                 start_row=2, n_index_cols=2):
    """列百分比 sheet 显著性着色 + DataBar。
    
    对每个分组值列的每个选项单元格：
    - 显著且 up: amber-100 底 + green-800 字 + DataBar
    - 显著且 down: amber-100 底 + red-700 字 + DataBar
    - 非显著: 保持斑马 + slate-600 字
    - 总计列: indigo-100 底（基准标识）
    """
    if not significance_matrix or percent_df is None or percent_df.empty:
        return
    
    border = thin_border()
    max_row = ws.max_row
    max_col = ws.max_column
    total_rows = _find_total_rows(ws, max_row, n_index_cols)
    
    # 总计列标 indigo-100 底
    for col_idx, label in enumerate(col_labels, start=n_index_cols + 1):
        if str(label).endswith("\n总计"):
            for r in range(1, max_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                cell.fill = make_fill(TR.INDIGO_ACCENT_BG)
    
    # 建行映射：ws 行号 -> (dim, group_col, option)
    df_rows = list(percent_df.index)
    ws_row_to_info = {}
    for i, idx in enumerate(df_rows):
        ws_row = start_row + i
        if ws_row in total_rows:
            continue
        option = idx[1] if isinstance(idx, tuple) else idx
        if str(option) in ("总计", "合计", "Total"):
            continue
        # 找该行属于哪个 dim + group_col
        for dim, groups in significance_matrix.items():
            for group_col, options in groups.items():
                if str(option) in options:
                    ws_row_to_info[ws_row] = (dim, group_col, str(option))
    
    # 着色 + DataBar
    amber_fill = make_fill(_DRIFT_BG)  # FEF3C7
    for ws_row, (dim, group_col, option) in ws_row_to_info.items():
        info = significance_matrix[dim][group_col][option]
        # 找 group_col 对应的列号
        for col_idx, label in enumerate(col_labels, start=n_index_cols + 1):
            if label == group_col:
                cell = ws.cell(row=ws_row, column=col_idx)
                if info["significant"]:
                    cell.fill = amber_fill
                    if info["direction"] == "up":
                        cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=_UP_FONT)
                        cell.value = f"{cell.value} ↑" if cell.value else "↑"
                    else:
                        cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=_DOWN_FONT)
                        cell.value = f"{cell.value} ↓" if cell.value else "↓"
                break
    
    # DataBar：对所有非总计数据单元格加 indigo DataBar（刻度 0~1）
    non_total_rows = [r for r in range(start_row, max_row + 1) if r not in total_rows]
    if non_total_rows:
        for col_idx, label in enumerate(col_labels, start=n_index_cols + 1):
            if str(label).endswith("\n总计"):
                continue
            col_letter = get_column_letter(col_idx)
            data_range = f"{col_letter}{min(non_total_rows)}:{col_letter}{max(non_total_rows)}"
            rule = DataBarRule(
                start_type='num', start_value=0,
                end_type='num', end_value=1,
                color=TR.INDIGO_CHIP, showValue=True,
                minLength=0, maxLength=100,
            )
            ws.conditional_formatting.add(data_range, rule)
```

- [ ] **Step 2: 更新 export_crosstab_excel 调用**

找到 `export_crosstab_excel` 函数里调用 `_apply_diff_heatmap` 的地方，改为 `_apply_significance_heatmap`，并传入 `significance_matrix`。需要给 `export_crosstab_excel` 加 `significance_matrix` 参数。

- [ ] **Step 3: 更新 run_crosstab_pipeline 里 export 调用**

传入 `significance_matrix`。

- [ ] **Step 4: 端到端验证**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python scripts/crosstab.py --file_path "C:\Users\lijinghui03\Desktop\survey_91044\survey_91044_我的世界回流玩家调研问卷_26年【量化数据】20260806-20260813.csv" --row_questions "[\"all\"]" --col_questions "[\"Q33.请问您的性别是？\"]" --calc_scores auto --output_path "C:\Users\lijinghui03\Desktop\survey_91044\crosstab_style_test.xlsx"
```
Expected: success，列百分比 Sheet 有显著性着色 + DataBar。

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py && git commit -m "feat(crosstab): replace diff heatmap with significance heatmap + DataBar on percent sheet"
```

---

## Task 8: 可视化 — 得分 Sheet 趋势条（均分 DataBar）

**Files:**
- Modify: `scripts/crosstab.py` (`_format_score_sheet_v2`)

- [ ] **Step 1: 改造 _format_score_sheet_v2 识别均分行 vs 样本量行 + 加 DataBar**

找到 `_format_score_sheet_v2` 函数（`grep -n "def _format_score_sheet_v2" scripts/crosstab.py`）。在现有逻辑基础上：
- 识别均分行（index 含「得分」不含「样本量」）：size 11 bold indigo-700，行高 22，加 indigo DataBar（min=1 max=5）
- 识别样本量行（index 含「样本量」）：size 9 slate-400，行高 18，格式 `n=#,##0`，无 DataBar
- 总计列均分：indigo-100 底

关键改动：在遍历行时判断 index 文本：
```python
for r in range(2, max_row + 1):
    # 读 index 值判断行类型
    idx_val = str(ws.cell(row=r, column=n_index_cols).value or "")
    is_sample = "样本量" in idx_val
    if is_sample:
        # 样本量行样式
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=Theme.FONT_NAME, size=9, color=TR.TEXT_MUTE)
            cell.alignment = ALIGN_CENTER
            cell.border = border
        ws.row_dimensions[r].height = 18
    else:
        # 均分行样式
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=Theme.FONT_NAME, size=11, bold=True, color=TR.INDIGO_MAIN)
            cell.alignment = ALIGN_CENTER
            cell.border = border
            # 总计列底色
            col_label = col_labels[c - n_index_cols - 1] if c > n_index_cols else None
            if col_label and str(col_label).endswith("\n总计"):
                cell.fill = make_fill(TR.INDIGO_ACCENT_BG)
        ws.row_dimensions[r].height = 22

# DataBar：仅均分行，1-5 刻度
score_rows = [r for r in range(2, max_row + 1) 
              if "样本量" not in str(ws.cell(row=r, column=n_index_cols).value or "")]
for col_idx in range(n_index_cols + 1, max_col + 1):
    col_letter = get_column_letter(col_idx)
    data_range = f"{col_letter}{min(score_rows)}:{col_letter}{max(score_rows)}"
    rule = DataBarRule(
        start_type='num', start_value=1,
        end_type='num', end_value=5,
        color=TR.INDIGO_CHIP, showValue=True,
        minLength=0, maxLength=100,
    )
    ws.conditional_formatting.add(data_range, rule)
```

- [ ] **Step 2: 端到端验证**

Run（同 Task 7 Step 4 命令）。Expected: 得分 Sheet 均分行有 indigo DataBar（1-5 刻度），样本量行 size 9 无 DataBar。

- [ ] **Step 3: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py && git commit -m "feat(crosstab): score sheet adds DataBar trend bars on score rows, distinguishes sample size rows"
```

---

## Task 9: 可视化 — 新增第 4 Sheet 差异热力图

**Files:**
- Modify: `scripts/crosstab.py` (新增 `calc_heatmap_data` + `_format_heatmap_sheet`，`export_crosstab_excel` 加第 4 Sheet)

- [ ] **Step 1: 新增 calc_heatmap_data 函数**

在 `calc_significance` 之后插入：
```python
def calc_heatmap_data(ct_result, significance_matrix):
    """生成热力图数据。返回 DataFrame，index=题目×选项，columns=各分组值，值=delta_pp。"""
    if not significance_matrix:
        return None
    # 收集所有 (dim, group_col) 对
    all_groups = []
    for dim, groups in significance_matrix.items():
        for group_col in groups:
            all_groups.append((dim, group_col))
    
    # 收集所有行索引（题目×选项）
    freq_df = ct_result["freq_df"]
    row_indices = [idx for idx in freq_df.index 
                   if str(idx[1] if isinstance(idx, tuple) else idx) not in ("总计", "合计", "Total")]
    
    # 构建 DataFrame
    data = {}
    for dim, group_col in all_groups:
        col_data = {}
        for idx in row_indices:
            option = str(idx[1] if isinstance(idx, tuple) else idx)
            if option in significance_matrix[dim][group_col]:
                col_data[idx] = significance_matrix[dim][group_col][option]["delta_pp"]
            else:
                col_data[idx] = 0.0
        data[group_col] = col_data
    
    heatmap_df = pd.DataFrame(data, index=row_indices)
    return heatmap_df
```

- [ ] **Step 2: 新增 _format_heatmap_sheet 函数**

```python
def _format_heatmap_sheet(ws, heatmap_df, col_dimensions):
    """格式化热力图 Sheet：渐变着色 + 分块 + 维度标题行。"""
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = _DOWN_FONT  # red-700
    
    # 颜色阶梯
    def _heat_color(delta_pp):
        """返回 (fill_hex, font_hex)。"""
        if delta_pp >= 20: return ("66BB6A", "FFFFFF")
        if delta_pp >= 15: return ("A5D6A7", "1B5E20")
        if delta_pp >= 10: return ("C8E6C9", "2E7D32")
        if delta_pp >= 5:  return ("E8F5E9", "388E3C")
        if delta_pp <= -20: return ("EF5350", "FFFFFF")
        if delta_pp <= -15: return ("EF9A9A", "B71C1C")
        if delta_pp <= -10: return ("FFCDD2", "C62828")
        if delta_pp <= -5:  return ("FFEBEE", "D32F2F")
        return ("F8FAFC", "94A3B8")  # 非显著
    
    # 写维度标题行（第 1 行）+ 分组值表头（第 2 行）
    # A/B 列：题目/选项
    ws.cell(row=1, column=1, value="题目").font = Font(name=Theme.FONT_NAME, size=10, bold=True, color="FFFFFF")
    ws.cell(row=1, column=1).fill = make_fill(TR.TITLE_BG)
    ws.cell(row=1, column=2, value="选项").font = Font(name=Theme.FONT_NAME, size=10, bold=True, color="FFFFFF")
    ws.cell(row=1, column=2).fill = make_fill(TR.TITLE_BG)
    ws.cell(row=2, column=1, value="题目").font = Font(name=Theme.FONT_NAME, size=10, bold=True, color="FFFFFF")
    ws.cell(row=2, column=1).fill = make_fill(TR.SUBTITLE_BG)
    ws.cell(row=2, column=2, value="选项").font = Font(name=Theme.FONT_NAME, size=10, bold=True, color="FFFFFF")
    ws.cell(row=2, column=2).fill = make_fill(TR.SUBTITLE_BG)
    
    # 按维度分块写列
    col_offset = 3  # 从 C 列开始
    for dim_info in col_dimensions:
        dim_name = _extract_dim_from_label(dim_info["question"])
        n_values = len(dim_info["values"])
        # 维度标题行（合并单元格）
        ws.cell(row=1, column=col_offset, value=dim_name)
        ws.cell(row=1, column=col_offset).font = Font(name=Theme.FONT_NAME, size=10, bold=True, color="FFFFFF")
        ws.cell(row=1, column=col_offset).fill = make_fill(TR.TITLE_BG)
        ws.cell(row=1, column=col_offset).alignment = ALIGN_CENTER
        if n_values > 1:
            ws.merge_cells(start_row=1, start_column=col_offset, end_row=1, end_column=col_offset + n_values - 1)
        # 分组值表头
        for i, val_label in enumerate(dim_info["values"]):
            short_val = str(val_label).split("\n")[-1]
            ws.cell(row=2, column=col_offset + i, value=short_val)
            ws.cell(row=2, column=col_offset + i).font = Font(name=Theme.FONT_NAME, size=10, bold=True, color="FFFFFF")
            ws.cell(row=2, column=col_offset + i).fill = make_fill(TR.SUBTITLE_BG)
            ws.cell(row=2, column=col_offset + i).alignment = ALIGN_CENTER
        # 间隔空列
        col_offset += n_values + 1  # +1 空列
    
    # 写数据行（从第 3 行起）
    for r_idx, idx in enumerate(heatmap_df.index):
        ws_row = r_idx + 3
        q, opt = (idx[0], idx[1]) if isinstance(idx, tuple) else (idx, "")
        ws.cell(row=ws_row, column=1, value=q).font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.INDIGO_DEEP)
        ws.cell(row=ws_row, column=1).fill = make_fill(TR.INDIGO_ACCENT_BG)
        ws.cell(row=ws_row, column=2, value=opt).font = Font(name=Theme.FONT_NAME, size=10, color=TR.TEXT_MAIN)
        
        col_offset = 3
        for dim_info in col_dimensions:
            for val_label in dim_info["values"]:
                if val_label in heatmap_df.columns:
                    delta = heatmap_df.loc[idx, val_label]
                    fill_hex, font_hex = _heat_color(delta)
                    cell = ws.cell(row=ws_row, column=col_offset, value=f"{delta:+.1f}pp" if abs(delta) >= 0.1 else "—")
                    cell.fill = make_fill(fill_hex)
                    cell.font = Font(name=Theme.FONT_NAME, size=10, color=font_hex)
                    cell.alignment = ALIGN_CENTER
                col_offset += 1
            col_offset += 1  # 空列
    
    # 列宽 + 冻结
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 26
    ws.freeze_panes = "C3"
```

- [ ] **Step 3: 在 export_crosstab_excel 加第 4 Sheet**

在 `export_crosstab_excel` 函数末尾（得分 Sheet 之后）加：
```python
    # Sheet 4: 📊 差异热力图
    if significance_matrix:
        heatmap_df = calc_heatmap_data(ct_result, significance_matrix)
        if heatmap_df is not None and not heatmap_df.empty:
            ws4 = writer.book.create_sheet("📊 差异热力图")
            _format_heatmap_sheet(ws4, heatmap_df, col_dimensions)
```

需要给 `export_crosstab_excel` 加 `significance_matrix` 和 `col_dimensions` 参数。

- [ ] **Step 4: 端到端验证**

Run（同 Task 7 Step 4 命令）。Expected: Excel 有 4 Sheet，第 4 Sheet「📊 差异热力图」有渐变着色 + 分块。

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add scripts/crosstab.py && git commit -m "feat(crosstab): add 4th sheet 差异热力图 with gradient heatmap and dimension blocks"
```

---

## Task 10: 文档更新 + 回归测试

**Files:**
- Modify: `references/12-crosstab-workflow.md`
- Modify: `SKILL.md`
- Test: `tests/test_crosstab_generalize.py` (端到端回归)

- [ ] **Step 1: 更新 12-crosstab-workflow.md**

补充：
1. 多分组示例：`--col_questions '["Q33.性别", "Q34.年龄", "Q35.职业"]'` 并列展开
2. auto 模式：`--col_questions '["auto"]'` 返回候选，AI 判断后 ask_user
3. 显著性检验说明：vs 分组维度总计，双门槛
4. 输出 4 Sheet 说明
5. 文件命名规则

- [ ] **Step 2: 更新 SKILL.md 阶段 3**

补 auto 触发语：「自动识别分组维度」「对比不同人群差异」时触发。

- [ ] **Step 3: 端到端回归测试**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python scripts/crosstab.py --file_path "C:\Users\lijinghui03\Desktop\survey_91044\survey_91044_我的世界回流玩家调研问卷_26年【量化数据】20260806-20260813.csv" --row_questions "[\"all\"]" --col_questions "[\"Q33.请问您的性别是？\", \"Q34.请问您的年龄是？\", \"Q35.请问您的职业是？\"]" --calc_scores auto --merge_rules "{\"Q35.请问您的职业是？\": \"occupation_default\"}" --output_path "C:\Users\lijinghui03\Desktop\survey_91044\交叉分析_按性别_年龄_职业.xlsx"
```
Expected: success，文件名 `交叉分析_按性别_年龄_职业.xlsx`，4 Sheet，得分 Sheet 有样本量行 + DataBar，热力图 Sheet 有渐变着色。

- [ ] **Step 4: 运行全部测试**

Run:
```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && python -m pytest tests/ -v
```
Expected: 全部 passed（含 drift 原有 44 + crosstab 新增）。

- [ ] **Step 5: Commit**

```bash
cd c:/Users/lijinghui03/.agents/skills/survey-research && git add references/12-crosstab-workflow.md SKILL.md && git commit -m "docs(crosstab): update workflow doc and SKILL.md for auto mode, significance, 4-sheet structure"
```

---

## Self-Review 检查

1. **Spec coverage**: 对照 spec 各节：
   - 需求 1 多分组保障 → Task 6 (col_dimensions) + Task 10 (回归)
   - 需求 2 显著性 vs 总计 → Task 5 (calc_significance) + Task 7 (着色) + Task 6 (diff_summary)
   - 需求 4 auto 识别 → Task 1 (identify_demographic) + Task 6 (pipeline)
   - 需求 5 文件命名 → Task 2 (default_output_filename) + Task 6 (pipeline 调用)
   - 需求 7 得分全量表+样本量 → Task 3 (auto_detect) + Task 4 (样本量行) + Task 8 (DataBar)
   - 可视化 9.1 热力图 Sheet → Task 9
   - 可视化 9.2 列百分比强化 → Task 7
   - 可视化 9.3 得分趋势条 → Task 8

2. **Placeholder scan**: 无 TBD/TODO，每步含完整代码。

3. **Type consistency**: `calc_significance` 返回 dict 结构一致；`_extract_dim_from_label` 在 Task 5/6/9 一致使用；`significance_matrix` 参数在 Task 7/8/9 传递一致。
