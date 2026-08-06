# 问卷时间异动诊断工具 (survey_drift) 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 survey-research skill 新增 `survey_drift.py`，对单份回流问卷按周/月/天分桶，逐题跑显著性检验、双门槛判异动，Agent 写一句话结论，输出 4-Sheet Excel。

**Architecture:** 单脚本双子命令：`analyze`（分桶+检验→`drift_findings.json`）与 `export`（findings+conclusions→Excel）。纯计算函数（分桶/两比例z/NPS/双门槛）先行 TDD，再组装 CLI 与 Excel 导出。复用 `load_and_classify.py` 的题型分类与 `_styles.py` 的样式主题。

**Tech Stack:** Python 3.10+，pandas、numpy、scipy（新增）、openpyxl、pytest。

> ⚠️ **本仓库当前非 git 仓库**（无 `.git`）。计划中的“Commit”步骤统一改为“Checkpoint：跑全量测试确认通过”。若后续 `git init`，可恢复为真实提交。所有路径基于 skill 根目录 `survey-research/`。

> 参考 spec：`specs/2026-08-06-survey-drift-design.md`

---

## 文件结构

| 文件 | 责任 |
|------|------|
| `scripts/survey_drift.py`（新建） | 核心：加载/分桶/取数/检验/组装 findings/导出 Excel/CLI |
| `scripts/requirements.txt`（改动） | 新增 `scipy>=1.10.0` |
| `tests/test_survey_drift.py`（新建） | 纯函数与端到端 smoke 测试 |
| `references/18-drift-workflow.md`（新建） | 工作流 reference |
| `SKILL.md`（改动） | 阶段6触发条件 + 脚本清单 + 后续操作提示 |

**贯穿全计划的函数签名契约**（各任务必须一致）：
- `_detect_csv_encoding(filepath, sample_size=8192) -> str`
- `load_df(file_path) -> pd.DataFrame`
- `week_label(dt) -> str` / `month_label(dt) -> str` / `day_label(dt) -> str`
- `bucketize(dt_series, granularity) -> (label_series: pd.Series, ordered_labels: list[str])`
- `two_prop_z(c1, n1, c2, n2) -> (z: float, p: float)`
- `compare_means(a_values, b_values) -> dict`  → `{"test", "p", "mean_a", "mean_b", "delta"}`
- `compute_nps(series) -> dict` → `{"nps", "promoter", "detractor", "n"}`
- `evaluate_drift(delta, p, kind) -> bool`  （`kind` ∈ `{"pp","mean"}`）
- `single_choice_props(df, col, label_series, ordered) -> (by_bucket: dict, sizes: dict)`
- `multi_choice_rates(df, subcols, root, label_series, ordered) -> (by_bucket: dict, sizes: dict)`
- `scale_means(df, col, label_series, ordered) -> (by_bucket: dict, sizes: dict)`
- `default_output_filename(granularity) -> str`
- `build_findings(...) -> dict`
- `export_excel(findings, conclusions, output_path) -> dict`

---

## Task 1: 脚手架 + 两比例 z 检验

**Files:**
- Create: `scripts/survey_drift.py`
- Modify: `scripts/requirements.txt`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 更新依赖**

在 `scripts/requirements.txt` 的 `numpy>=1.23.0` 下一行加：
```
scipy>=1.10.0
```

- [ ] **Step 2: 写失败测试（two_prop_z）**

创建 `tests/test_survey_drift.py`：
```python
import importlib.util
from pathlib import Path
import math

MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "survey_drift.py"
spec = importlib.util.spec_from_file_location("survey_drift", MODULE_PATH)
survey_drift = importlib.util.module_from_spec(spec)
spec.loader.exec_module(survey_drift)


def test_two_prop_z_no_diff_gives_high_p():
    z, p = survey_drift.two_prop_z(50, 100, 50, 100)
    assert abs(z) < 1e-9
    assert p > 0.99


def test_two_prop_z_big_diff_is_significant():
    z, p = survey_drift.two_prop_z(80, 100, 40, 100)
    assert p < 0.01
    assert z > 0
```

- [ ] **Step 3: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -q`
Expected: FAIL（`survey_drift.py` 不存在 / 无 `two_prop_z`）

- [ ] **Step 4: 写最小实现**

创建 `scripts/survey_drift.py`：
```python
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 时间异动诊断 (survey_drift)
==========================================

单份回流问卷按 周/月/天 分桶，逐题相邻期显著性检验，双门槛判异动，
Agent 写一句话结论，导出 4-Sheet Excel。

子命令:
    analyze  分桶 + 检验 → drift_findings.json
    export   findings + conclusions → Excel
"""

import argparse
import json
import os
import sys
from datetime import timedelta

import numpy as np
import pandas as pd
from scipy import stats


# ===================== 统计基元 ===================== #

def two_prop_z(c1, n1, c2, n2):
    """两比例 z 检验（双侧，pooled）。返回 (z, p)。n 为 0 时返回 (0.0, 1.0)。"""
    if n1 == 0 or n2 == 0:
        return 0.0, 1.0
    p1, p2 = c1 / n1, c2 / n2
    p_pool = (c1 + c2) / (n1 + n2)
    denom = p_pool * (1 - p_pool) * (1 / n1 + 1 / n2)
    if denom <= 0:
        return 0.0, 1.0
    z = (p1 - p2) / (denom ** 0.5)
    p = 2 * stats.norm.sf(abs(z))
    return float(z), float(p)
```

- [ ] **Step 5: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -q`
Expected: PASS（2 passed）

- [ ] **Step 6: Checkpoint** — 全量测试通过即可进入下一任务。

---

## Task 2: 分桶与周/月/日标签

**Files:**
- Modify: `scripts/survey_drift.py`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加到 `tests/test_survey_drift.py`：
```python
import pandas as pd


def test_labels():
    dt = pd.Timestamp("2026-04-06")  # 周一, ISO 第15周
    assert survey_drift.week_label(dt) == "第15周（4.6-4.12）"
    assert survey_drift.month_label(dt) == "26年4月"
    assert survey_drift.day_label(dt) == "2026-04-06"


def test_bucketize_orders_chronologically():
    s = pd.to_datetime(pd.Series([
        "2026-04-06", "2026-04-13", "2026-04-06", "2026-04-20"
    ]))
    labels, ordered = survey_drift.bucketize(s, "week")
    assert ordered == ["第15周（4.6-4.12）", "第16周（4.13-4.19）", "第17周（4.20-4.26）"]
    assert list(labels) == [
        "第15周（4.6-4.12）", "第16周（4.13-4.19）",
        "第15周（4.6-4.12）", "第17周（4.20-4.26）",
    ]
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -k "labels or bucketize" -q`
Expected: FAIL（无 `week_label` 等）

- [ ] **Step 3: 写实现**

在 `scripts/survey_drift.py` 的“统计基元”区块前插入标签与分桶函数：
```python
# ===================== 时间标签与分桶 ===================== #

def _fmt_md(d):
    return f"{d.month}.{d.day}"


def week_label(dt):
    iso = dt.isocalendar()
    monday = dt - timedelta(days=iso[2] - 1)
    sunday = monday + timedelta(days=6)
    return f"第{iso[1]}周（{_fmt_md(monday)}-{_fmt_md(sunday)}）"


def month_label(dt):
    return f"{str(dt.year)[2:]}年{dt.month}月"


def day_label(dt):
    return f"{dt.year}-{dt.month:02d}-{dt.day:02d}"


_LABELERS = {"week": week_label, "month": month_label, "day": day_label}


def bucketize(dt_series, granularity):
    """返回 (label_series, ordered_labels)。ordered_labels 按桶内最早时间升序（旧→新）。"""
    labeler = _LABELERS[granularity]
    labels = dt_series.apply(lambda d: labeler(d) if pd.notna(d) else None)
    valid = pd.DataFrame({"label": labels, "dt": dt_series}).dropna(subset=["label"])
    order = valid.groupby("label")["dt"].min().sort_values()
    return labels, list(order.index)
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -k "labels or bucketize" -q`
Expected: PASS

- [ ] **Step 5: Checkpoint** — `python -m pytest tests/test_survey_drift.py -q` 全绿。

---

## Task 3: NPS 计算 + 双门槛判定

**Files:**
- Modify: `scripts/survey_drift.py`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加：
```python
def test_compute_nps():
    # 5 推荐者(9-10), 3 贬损者(0-6), 2 中立(7-8)
    s = pd.Series([10, 10, 9, 9, 9, 7, 8, 0, 3, 6])
    r = survey_drift.compute_nps(s)
    assert r["n"] == 10
    assert r["promoter"] == 5
    assert r["detractor"] == 3
    assert round(r["nps"], 1) == 20.0  # (5-3)/10 * 100


def test_evaluate_drift_double_threshold():
    # 显著且实际差异达标 → True
    assert survey_drift.evaluate_drift(6.0, 0.02, "pp") is True
    # 显著但差异不达标（<5pp）→ False
    assert survey_drift.evaluate_drift(3.0, 0.01, "pp") is False
    # 达标但不显著 → False
    assert survey_drift.evaluate_drift(8.0, 0.20, "pp") is False
    # 均分门槛 0.1
    assert survey_drift.evaluate_drift(0.15, 0.03, "mean") is True
    assert survey_drift.evaluate_drift(0.05, 0.03, "mean") is False
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -k "nps or drift" -q`
Expected: FAIL

- [ ] **Step 3: 写实现**

在“统计基元”区块追加：
```python
def compute_nps(series):
    """0~10 推荐题 → NPS。9-10 推荐者, 0-6 贬损者, 7-8 中立。"""
    vals = pd.to_numeric(series, errors="coerce").dropna()
    n = len(vals)
    if n == 0:
        return {"nps": 0.0, "promoter": 0, "detractor": 0, "n": 0}
    promoter = int((vals >= 9).sum())
    detractor = int((vals <= 6).sum())
    nps = (promoter - detractor) / n * 100
    return {"nps": float(nps), "promoter": promoter, "detractor": detractor, "n": n}


def compare_means(a_values, b_values):
    """相邻期均分比较。任一样本 n<30 用 Mann-Whitney U，否则 t 检验。
    返回 {test, p, mean_a, mean_b, delta}（delta = a - b，a 为较新期）。"""
    a = pd.to_numeric(pd.Series(a_values), errors="coerce").dropna()
    b = pd.to_numeric(pd.Series(b_values), errors="coerce").dropna()
    mean_a = float(a.mean()) if len(a) else 0.0
    mean_b = float(b.mean()) if len(b) else 0.0
    delta = mean_a - mean_b
    if len(a) < 2 or len(b) < 2:
        return {"test": "insufficient", "p": 1.0, "mean_a": mean_a, "mean_b": mean_b, "delta": delta}
    if len(a) < 30 or len(b) < 30:
        _, p = stats.mannwhitneyu(a, b, alternative="two-sided")
        test = "mann_whitney_u"
    else:
        _, p = stats.ttest_ind(a, b, equal_var=False)
        test = "t_test"
    return {"test": test, "p": float(p), "mean_a": mean_a, "mean_b": mean_b, "delta": delta}


def evaluate_drift(delta, p, kind):
    """双门槛：p<0.05 且 实际差异达标（pp≥5 / mean≥0.1）。"""
    if p >= 0.05:
        return False
    if kind == "pp":
        return abs(delta) >= 5.0
    if kind == "mean":
        return abs(delta) >= 0.1
    return False
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -k "nps or drift" -q`
Expected: PASS

- [ ] **Step 5: Checkpoint** — 全量测试全绿。

---

## Task 4: 逐桶取数（单选/多选/量表）

**Files:**
- Modify: `scripts/survey_drift.py`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加：
```python
def _demo_df():
    return pd.DataFrame({
        "结束答题时间": pd.to_datetime([
            "2026-04-06", "2026-04-06", "2026-04-13", "2026-04-13"]),
        "Q1.整体满意度": [5, 4, 3, 3],
        "Q7.活动评价": ["满意", "满意", "一般", "满意"],
    })


def test_single_choice_props():
    df = _demo_df()
    labels, ordered = survey_drift.bucketize(df["结束答题时间"], "week")
    by_bucket, sizes = survey_drift.single_choice_props(df, "Q7.活动评价", labels, ordered)
    b0, b1 = ordered
    assert sizes[b0] == 2 and sizes[b1] == 2
    assert round(by_bucket[b0]["满意"], 3) == 1.0
    assert round(by_bucket[b1]["满意"], 3) == 0.5


def test_scale_means():
    df = _demo_df()
    labels, ordered = survey_drift.bucketize(df["结束答题时间"], "week")
    by_bucket, sizes = survey_drift.scale_means(df, "Q1.整体满意度", labels, ordered)
    b0, b1 = ordered
    assert round(by_bucket[b0], 2) == 4.5
    assert round(by_bucket[b1], 2) == 3.0


def test_multi_choice_rates():
    df = pd.DataFrame({
        "结束答题时间": pd.to_datetime(["2026-04-06", "2026-04-06"]),
        "Q9.喜欢的模式:生存": ["生存", None],
        "Q9.喜欢的模式:创造": ["创造", "创造"],
    })
    labels, ordered = survey_drift.bucketize(df["结束答题时间"], "week")
    subcols = ["Q9.喜欢的模式:生存", "Q9.喜欢的模式:创造"]
    by_bucket, sizes = survey_drift.multi_choice_rates(df, subcols, "Q9.", labels, ordered)
    b0 = ordered[0]
    assert round(by_bucket[b0]["生存"], 2) == 0.5
    assert round(by_bucket[b0]["创造"], 2) == 1.0
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -k "props or means or rates" -q`
Expected: FAIL

- [ ] **Step 3: 写实现**

在 `scripts/survey_drift.py` 追加取数区块：
```python
# ===================== 逐桶取数 ===================== #

def _bucket_mask(label_series, bucket):
    return label_series == bucket


def single_choice_props(df, col, label_series, ordered):
    """单选题各桶各选项占比。返回 (by_bucket={bucket:{option:prop}}, sizes={bucket:n})。"""
    by_bucket, sizes = {}, {}
    for b in ordered:
        sub = df.loc[_bucket_mask(label_series, b), col].dropna()
        sub = sub[sub.astype(str).str.strip() != ""]
        n = len(sub)
        sizes[b] = n
        if n == 0:
            by_bucket[b] = {}
            continue
        by_bucket[b] = (sub.astype(str).value_counts(normalize=True)).to_dict()
    return by_bucket, sizes


def _is_selected(v):
    if pd.isna(v):
        return False
    s = str(v).strip()
    return s not in ("", "0", "nan", "NaN", "None", "否", "未选择")


def multi_choice_rates(df, subcols, root, label_series, ordered):
    """多选题各桶各选项勾选率。选项名取子列名冒号后部分（无冒号则去 root 前缀）。"""
    def opt_name(sc):
        s = str(sc)
        for sep in (":", "："):
            if sep in s:
                return s.split(sep, 1)[1].strip()
        return s[len(root):].strip() if s.startswith(root) else s

    by_bucket, sizes = {}, {}
    for b in ordered:
        mask = _bucket_mask(label_series, b)
        block = df.loc[mask, subcols]
        n = len(block)
        sizes[b] = n
        rates = {}
        for sc in subcols:
            if "输入文本" in str(sc):
                continue
            sel = block[sc].apply(_is_selected).sum()
            rates[opt_name(sc)] = (sel / n) if n else 0.0
        by_bucket[b] = rates
    return by_bucket, sizes


def scale_means(df, col, label_series, ordered):
    """量表题各桶均分。返回 (by_bucket={bucket:mean}, sizes={bucket:n})。"""
    by_bucket, sizes = {}, {}
    for b in ordered:
        vals = pd.to_numeric(df.loc[_bucket_mask(label_series, b), col], errors="coerce").dropna()
        sizes[b] = len(vals)
        by_bucket[b] = float(vals.mean()) if len(vals) else 0.0
    return by_bucket, sizes
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -k "props or means or rates" -q`
Expected: PASS

- [ ] **Step 5: Checkpoint** — 全量测试全绿。

---

## Task 5: 相邻期检验组装 + 指标识别 + load_df

**Files:**
- Modify: `scripts/survey_drift.py`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加：
```python
def test_adjacent_prop_tests_flags_drift():
    # b0 满意=100%(n=60), b1 满意=50%(n=60) → 相邻期 z 检验显著且 >5pp
    by_bucket = {"b0": {"满意": 1.0}, "b1": {"满意": 0.5}}
    sizes = {"b0": 60, "b1": 60}
    ordered = ["b0", "b1"]  # 旧→新
    res = survey_drift.adjacent_prop_tests(by_bucket, sizes, ordered, min_n=30)
    # 只有一对相邻：newer=b1, older=b0
    row = [r for r in res if r["option"] == "满意"][0]
    assert row["from"] == "b0" and row["to"] == "b1"
    assert row["significant"] is True
    assert row["drift"] is True
    assert row["direction"] == "down"


def test_identify_metric_cols():
    single = ["Q1.请问您对本赛季的满意度如何？", "Q51.您有多大可能将本游戏推荐给朋友？", "Q3.性别"]
    nps, sat = survey_drift.identify_metric_cols(single)
    assert nps == "Q51.您有多大可能将本游戏推荐给朋友？"
    assert "Q1.请问您对本赛季的满意度如何？" in sat
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -k "adjacent or identify" -q`
Expected: FAIL

- [ ] **Step 3: 写实现**

追加：
```python
# ===================== 相邻期检验 ===================== #

def _adjacent_pairs(ordered):
    """返回 [(older, newer), ...]，按 ordered(旧→新) 相邻配对。"""
    return [(ordered[i - 1], ordered[i]) for i in range(1, len(ordered))]


def adjacent_prop_tests(by_bucket, sizes, ordered, min_n=30):
    """对每个选项、每对相邻期跑两比例 z 检验。返回 list[dict]。"""
    options = set()
    for b in ordered:
        options.update(by_bucket.get(b, {}).keys())
    results = []
    for older, newer in _adjacent_pairs(ordered):
        n_old, n_new = sizes.get(older, 0), sizes.get(newer, 0)
        low_n = n_old < min_n or n_new < min_n
        for opt in sorted(options):
            p_old = by_bucket.get(older, {}).get(opt, 0.0)
            p_new = by_bucket.get(newer, {}).get(opt, 0.0)
            c_old, c_new = round(p_old * n_old), round(p_new * n_new)
            z, pval = two_prop_z(c_new, n_new, c_old, n_old)
            delta_pp = (p_new - p_old) * 100
            drift = (not low_n) and evaluate_drift(delta_pp, pval, "pp")
            results.append({
                "option": opt, "from": older, "to": newer,
                "delta_pp": round(delta_pp, 1), "test": "two_prop_z",
                "p": round(pval, 4), "significant": pval < 0.05,
                "drift": drift, "low_n": low_n,
                "direction": "up" if delta_pp > 0 else ("down" if delta_pp < 0 else "flat"),
            })
    return results


def adjacent_mean_tests(df, col, label_series, ordered, min_n=30):
    """对量表题每对相邻期跑均分检验。返回 list[dict]。"""
    results = []
    for older, newer in _adjacent_pairs(ordered):
        a = df.loc[label_series == newer, col]
        b = df.loc[label_series == older, col]
        cmp = compare_means(a, b)
        low_n = len(pd.to_numeric(a, errors="coerce").dropna()) < min_n or \
                len(pd.to_numeric(b, errors="coerce").dropna()) < min_n
        drift = (not low_n) and evaluate_drift(cmp["delta"], cmp["p"], "mean")
        results.append({
            "from": older, "to": newer, "delta": round(cmp["delta"], 3),
            "test": cmp["test"], "p": round(cmp["p"], 4),
            "significant": cmp["p"] < 0.05, "drift": drift, "low_n": low_n,
            "direction": "up" if cmp["delta"] > 0 else ("down" if cmp["delta"] < 0 else "flat"),
        })
    return results


# ===================== 指标题识别 ===================== #

_NPS_KEYS = ["推荐给", "推荐给身边", "有多大可能将", "推荐给朋友", "净推荐"]
_SAT_KEYS = ["满意度", "满意程度", "整体满意", "整体体验"]


def identify_metric_cols(single_choice_cols):
    """从单选/量表列名按关键词识别 NPS 题与满意度题。返回 (nps_col|None, [sat_cols])。"""
    nps_col = None
    sat_cols = []
    for c in single_choice_cols:
        name = str(c)
        if nps_col is None and any(k in name for k in _NPS_KEYS):
            nps_col = c
            continue
        if any(k in name for k in _SAT_KEYS):
            sat_cols.append(c)
    return nps_col, sat_cols


# ===================== 数据加载 ===================== #

def _detect_csv_encoding(filepath, sample_size=8192):
    with open(filepath, "rb") as f:
        raw = f.read(sample_size)
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    try:
        raw.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError:
        return "gbk"


def load_df(file_path):
    ext = file_path.rsplit(".", 1)[-1].lower()
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(file_path)
    else:
        df = pd.read_csv(file_path, encoding=_detect_csv_encoding(file_path), low_memory=False)
    df.columns = [str(c).strip() for c in df.columns]
    return df
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -k "adjacent or identify" -q`
Expected: PASS

- [ ] **Step 5: Checkpoint** — 全量测试全绿。

---

## Task 6: build_findings 组装 + analyze CLI

**Files:**
- Modify: `scripts/survey_drift.py`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加：
```python
import subprocess, sys, json, os


def test_build_findings_structure(tmp_path):
    df = pd.DataFrame({
        "结束答题时间": pd.to_datetime(
            ["2026-04-06"] * 40 + ["2026-04-13"] * 40),
        "Q1.整体满意度": [5] * 40 + [3] * 40,
        "Q7.活动评价（单选）": (["满意"] * 40) + (["满意"] * 20 + ["一般"] * 20),
    })
    classification = {
        "single_choice": ["Q1.整体满意度", "Q7.活动评价（单选）"],
        "multi_choice": {}, "matrix_scale": {}, "text": [], "meta": ["结束答题时间"],
    }
    findings = survey_drift.build_findings(
        df, classification, granularity="week", time_col="结束答题时间",
        nps_col=None, satisfaction_cols=["Q1.整体满意度"], min_n=30)
    assert findings["granularity"] == "week"
    assert len(findings["buckets"]) == 2
    q_names = [q["question"] for q in findings["questions"]]
    assert "Q7.活动评价（单选）" in q_names
    assert any(m["type"] == "satisfaction_mean" for m in findings["metrics"])


def test_analyze_cli_end_to_end(tmp_path):
    csv = tmp_path / "demo.csv"
    pd.DataFrame({
        "结束答题时间": (["2026-04-06 10:00:00"] * 40 + ["2026-04-13 10:00:00"] * 40),
        "Q1.整体满意度": [5] * 40 + [3] * 40,
        "Q51.您有多大可能将本游戏推荐给朋友？": [10] * 40 + [5] * 40,
    }).to_csv(csv, index=False, encoding="utf-8-sig")
    out = tmp_path / "findings.json"
    r = subprocess.run(
        [sys.executable, str(MODULE_PATH), "analyze",
         "--file_path", str(csv), "--granularity", "week",
         "--findings_out", str(out)],
        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    payload = json.loads(r.stdout)
    assert payload["status"] == "success"
    assert out.exists()
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["granularity"] == "week"
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -k "build_findings or analyze_cli" -q`
Expected: FAIL

- [ ] **Step 3: 写实现**

追加 `build_findings` 与 `analyze` + CLI（先只写 build_findings 和 analyze，`export` 在 Task 7）：
```python
# ===================== findings 组装 ===================== #

def build_findings(df, classification, granularity, time_col,
                   nps_col, satisfaction_cols, min_n=30):
    if time_col not in df.columns:
        raise KeyError(f"时间列不存在：{time_col}")
    dt = pd.to_datetime(df[time_col], errors="coerce")
    labels, ordered = bucketize(dt, granularity)
    sizes_all = {b: int((labels == b).sum()) for b in ordered}
    low_n_buckets = [b for b, n in sizes_all.items() if n < min_n]

    metrics = []
    # 满意度均分
    for col in (satisfaction_cols or []):
        if col not in df.columns:
            continue
        by_bucket, _ = scale_means(df, col, labels, ordered)
        adj = adjacent_mean_tests(df, col, labels, ordered, min_n)
        metrics.append({
            "name": f"{col} 均分", "type": "satisfaction_mean", "source_col": col,
            "by_bucket": {b: round(by_bucket[b], 2) for b in ordered}, "adjacent": adj,
        })
    # NPS
    if nps_col and nps_col in df.columns:
        by_bucket, sizes = {}, {}
        for b in ordered:
            r = compute_nps(df.loc[labels == b, nps_col])
            by_bucket[b] = round(r["nps"], 1)
            sizes[b] = r["n"]
        adj = []
        for older, newer in _adjacent_pairs(ordered):
            r_o = compute_nps(df.loc[labels == older, nps_col])
            r_n = compute_nps(df.loc[labels == newer, nps_col])
            # 推荐者率 vs 贬损者率合并为 NPS 差；用推荐者两比例 z 近似显著性
            z, pval = two_prop_z(r_n["promoter"], r_n["n"], r_o["promoter"], r_o["n"])
            delta_pp = by_bucket[newer] - by_bucket[older]
            low_n = r_o["n"] < min_n or r_n["n"] < min_n
            adj.append({
                "from": older, "to": newer, "delta_pp": round(delta_pp, 1),
                "test": "two_prop_z", "p": round(pval, 4), "significant": pval < 0.05,
                "drift": (not low_n) and evaluate_drift(delta_pp, pval, "pp"),
                "low_n": low_n,
                "direction": "up" if delta_pp > 0 else ("down" if delta_pp < 0 else "flat"),
            })
        metrics.append({
            "name": "NPS", "type": "nps", "source_col": nps_col,
            "by_bucket": by_bucket, "adjacent": adj,
        })

    questions = []
    metric_source_cols = set(satisfaction_cols or []) | ({nps_col} if nps_col else set())
    # 单选题
    for col in classification.get("single_choice", []):
        by_bucket, sizes = single_choice_props(df, col, labels, ordered)
        opt_tests = adjacent_prop_tests(by_bucket, sizes, ordered, min_n)
        # 整体卡方（最新相邻期）
        overall = _overall_chi_square(by_bucket, sizes, ordered)
        drift = any(t["drift"] for t in opt_tests)
        questions.append({
            "question": col, "type": "single_choice",
            "options": sorted({o for b in ordered for o in by_bucket.get(b, {})}),
            "by_bucket": {b: {k: round(v, 4) for k, v in by_bucket.get(b, {}).items()} for b in ordered},
            "sizes": sizes, "overall_test": overall,
            "adjacent_option_tests": opt_tests, "drift": drift,
            "low_n": any(sizes.get(b, 0) < min_n for b in ordered),
        })
    # 多选题
    for root, subcols in classification.get("multi_choice", {}).items():
        by_bucket, sizes = multi_choice_rates(df, subcols, root, labels, ordered)
        opt_tests = adjacent_prop_tests(by_bucket, sizes, ordered, min_n)
        drift = any(t["drift"] for t in opt_tests)
        questions.append({
            "question": root, "type": "multi_choice",
            "options": sorted({o for b in ordered for o in by_bucket.get(b, {})}),
            "by_bucket": {b: {k: round(v, 4) for k, v in by_bucket.get(b, {}).items()} for b in ordered},
            "sizes": sizes, "overall_test": None,
            "adjacent_option_tests": opt_tests, "drift": drift,
            "low_n": any(sizes.get(b, 0) < min_n for b in ordered),
        })

    return {
        "granularity": granularity, "time_col": time_col,
        "buckets": ordered, "bucket_sizes": sizes_all, "low_n_buckets": low_n_buckets,
        "metrics": metrics, "questions": questions,
        "nps_col": nps_col, "satisfaction_cols": satisfaction_cols or [],
    }


def _overall_chi_square(by_bucket, sizes, ordered):
    """最新相邻期两桶 × 选项 的整体卡方。"""
    if len(ordered) < 2:
        return None
    older, newer = ordered[-2], ordered[-1]
    options = sorted({o for b in (older, newer) for o in by_bucket.get(b, {})})
    if not options:
        return None
    table = []
    for b in (older, newer):
        n = sizes.get(b, 0)
        table.append([round(by_bucket.get(b, {}).get(o, 0.0) * n) for o in options])
    arr = np.array(table)
    if arr.sum() == 0 or (arr.sum(axis=0) == 0).any():
        return None
    try:
        chi2, p, _, _ = stats.chi2_contingency(arr)
    except ValueError:
        return None
    return {"test": "chi_square", "from": older, "to": newer,
            "p": round(float(p), 4), "significant": p < 0.05}
```

在文件末尾加 CLI（`export` 分支先占位，Task 7 补全）：
```python
# ===================== CLI ===================== #

def default_output_filename(granularity):
    label = {"week": "按周", "month": "按月", "day": "按天"}.get(granularity, granularity)
    from datetime import datetime
    return f"回流异动诊断_{label}_{datetime.now():%Y%m%d_%H%M}.xlsx"


def _cmd_analyze(args):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from load_and_classify import classify_columns
    df = load_df(args.file_path)
    if args.time_col not in df.columns:
        return {"status": "need_input", "reason": "time_col_missing",
                "message": f"时间列 '{args.time_col}' 不存在，可用列：{list(df.columns[:20])}"}
    classification = classify_columns(df)
    single = classification["single_choice"]
    nps_col = args.nps_col or identify_metric_cols(single)[0]
    sat_cols = args.satisfaction_cols or identify_metric_cols(single)[1]
    if not nps_col and not sat_cols:
        return {"status": "need_input", "reason": "no_metric",
                "message": "未能自动识别 NPS/满意度题，请用 --nps_col / --satisfaction_cols 指定"}
    findings = build_findings(df, classification, args.granularity, args.time_col,
                              nps_col, sat_cols, args.min_n)
    out = args.findings_out or os.path.join(
        os.path.dirname(os.path.abspath(args.file_path)), "drift_findings.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(findings, f, ensure_ascii=False, indent=2)
    return {
        "status": "success", "granularity": args.granularity,
        "buckets": findings["buckets"], "bucket_sizes": findings["bucket_sizes"],
        "low_n_buckets": findings["low_n_buckets"],
        "questions_total": len(findings["questions"]),
        "questions_with_drift": sum(1 for q in findings["questions"] if q["drift"]),
        "findings_out": out, "nps_col": nps_col, "satisfaction_cols": sat_cols,
    }


def main():
    parser = argparse.ArgumentParser(description="问卷时间异动诊断")
    sub = parser.add_subparsers(dest="cmd", required=True)

    pa = sub.add_parser("analyze", help="分桶 + 检验 → findings JSON")
    pa.add_argument("--file_path", required=True)
    pa.add_argument("--granularity", required=True, choices=["week", "month", "day"])
    pa.add_argument("--time_col", default="结束答题时间")
    pa.add_argument("--nps_col", default=None)
    pa.add_argument("--satisfaction_cols", nargs="*", default=None)
    pa.add_argument("--min_n", type=int, default=30)
    pa.add_argument("--findings_out", default=None)

    pe = sub.add_parser("export", help="findings + conclusions → Excel")
    pe.add_argument("--findings", required=True)
    pe.add_argument("--conclusions", default=None)
    pe.add_argument("--output_path", default=None)

    args = parser.parse_args()
    if args.cmd == "analyze":
        result = _cmd_analyze(args)
    else:
        result = _cmd_export(args)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result.get("status") not in ("success", None):
        sys.exit(0 if result.get("status") == "need_input" else 1)


if __name__ == "__main__":
    main()
```

> 注：`_cmd_export` 在 Task 7 定义；本任务可临时加 `def _cmd_export(args): return {"status": "error", "message": "not implemented"}` 占位，Task 7 替换。

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -k "build_findings or analyze_cli" -q`
Expected: PASS

- [ ] **Step 5: Checkpoint** — 全量测试全绿。

---

## Task 7: export Excel（4 Sheet）+ export CLI

**Files:**
- Modify: `scripts/survey_drift.py`
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加：
```python
from openpyxl import load_workbook


def test_export_creates_four_sheets(tmp_path):
    findings = {
        "granularity": "week", "time_col": "结束答题时间",
        "buckets": ["第15周（4.6-4.12）", "第16周（4.13-4.19）"],
        "bucket_sizes": {"第15周（4.6-4.12）": 40, "第16周（4.13-4.19）": 40},
        "low_n_buckets": [],
        "metrics": [{
            "name": "Q1.整体满意度 均分", "type": "satisfaction_mean", "source_col": "Q1.整体满意度",
            "by_bucket": {"第15周（4.6-4.12）": 4.5, "第16周（4.13-4.19）": 3.0},
            "adjacent": [{"from": "第15周（4.6-4.12）", "to": "第16周（4.13-4.19）",
                          "delta": -1.5, "test": "t_test", "p": 0.001,
                          "significant": True, "drift": True, "low_n": False, "direction": "down"}],
        }],
        "questions": [{
            "question": "Q7.活动评价（单选）", "type": "single_choice",
            "options": ["满意", "一般"],
            "by_bucket": {"第15周（4.6-4.12）": {"满意": 1.0, "一般": 0.0},
                          "第16周（4.13-4.19）": {"满意": 0.5, "一般": 0.5}},
            "sizes": {"第15周（4.6-4.12）": 40, "第16周（4.13-4.19）": 40},
            "overall_test": {"test": "chi_square", "p": 0.001, "significant": True},
            "adjacent_option_tests": [{"option": "满意", "from": "第15周（4.6-4.12）",
                "to": "第16周（4.13-4.19）", "delta_pp": -50.0, "test": "two_prop_z",
                "p": 0.001, "significant": True, "drift": True, "low_n": False, "direction": "down"}],
            "drift": True, "low_n": False,
        }],
        "nps_col": None, "satisfaction_cols": ["Q1.整体满意度"],
    }
    conclusions = {"Q7.活动评价（单选）": "满意占比从100%骤降至50%，显著恶化，需排查活动体验。"}
    out = tmp_path / "report.xlsx"
    r = survey_drift.export_excel(findings, conclusions, str(out))
    assert r["status"] == "success"
    wb = load_workbook(out)
    assert "📊 指标总览" in wb.sheetnames
    assert "📈 逐题异动明细" in wb.sheetnames
    assert "⚠️ 异动汇总" in wb.sheetnames
    assert "ℹ️ 方法与样本" in wb.sheetnames


def test_default_output_filename():
    name = survey_drift.default_output_filename("week")
    assert name.startswith("回流异动诊断_按周_")
    assert name.endswith(".xlsx")
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_survey_drift.py -k "export or default_output" -q`
Expected: FAIL

- [ ] **Step 3: 写实现**

在 CLI 区块前追加 Excel 导出（复用 `_styles`）：
```python
# ===================== Excel 导出 ===================== #

def _trend_mark(delta, significant, is_pp):
    unit = "pp" if is_pp else "分"
    if not significant:
        return f"— {delta:+.1f}{unit}（不显著）"
    arrow = "▲" if delta > 0 else "▼"
    return f"{arrow} {delta:+.1f}{unit}"


def export_excel(findings, conclusions, output_path):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import _styles as st
    from openpyxl import Workbook

    conclusions = conclusions or {}
    buckets = findings["buckets"]
    wb = Workbook()

    # ---- Sheet 1: 指标总览 ----
    ws1 = wb.active
    ws1.title = "📊 指标总览"
    header = ["指标"] + buckets + ["最新vs上期", "是否显著"]
    ws1.append(header)
    sizes = findings["bucket_sizes"]
    ws1.append(["样本量"] + [sizes.get(b, 0) for b in buckets] + ["—", "—"])
    for m in findings["metrics"]:
        is_pp = m["type"] == "nps"
        row = [m["name"]] + [m["by_bucket"].get(b, "") for b in buckets]
        if m["adjacent"]:
            last = m["adjacent"][-1]
            d = last.get("delta_pp", last.get("delta", 0.0))
            row += [_trend_mark(d, last["significant"], is_pp),
                    f"✅ p={last['p']}" if last["significant"] else f"✗ p={last['p']}"]
        else:
            row += ["—", "—"]
        ws1.append(row)
    st.format_score_sheet(ws1)

    # ---- Sheet 2: 逐题异动明细 ----
    ws2 = wb.create_sheet("📈 逐题异动明细")
    ws2.append(["题目", "选项"] + buckets + ["最新Δ", "p", "异动", "AI结论"])
    for q in findings["questions"]:
        opts = q["options"]
        start_row = ws2.max_row + 1
        latest_tests = {t["option"]: t for t in q["adjacent_option_tests"]
                        if t["to"] == buckets[-1]} if len(buckets) >= 2 else {}
        for i, opt in enumerate(opts):
            t = latest_tests.get(opt, {})
            ws2.append([
                q["question"] if i == 0 else "", opt,
                *[q["by_bucket"].get(b, {}).get(opt, 0.0) for b in buckets],
                t.get("delta_pp", ""), t.get("p", ""),
                "✅" if t.get("drift") else "", "",
            ])
        end_row = ws2.max_row
        concl = conclusions.get(q["question"], "")
        concl_col = 2 + len(buckets) + 4  # AI结论列号
        if end_row >= start_row:
            ws2.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws2.merge_cells(start_row=start_row, start_column=concl_col,
                            end_row=end_row, end_column=concl_col)
            ws2.cell(row=start_row, column=concl_col, value=concl)
    _format_detail_sheet(ws2, len(buckets))

    # ---- Sheet 3: 异动汇总 ----
    ws3 = wb.create_sheet("⚠️ 异动汇总")
    ws3.append(["题目/指标", "变化项", "方向", "幅度", "显著性", "AI结论"])
    any_drift = False
    for m in findings["metrics"]:
        for t in m["adjacent"]:
            if t.get("drift"):
                any_drift = True
                d = t.get("delta_pp", t.get("delta", 0.0))
                arrow = "▲" if d > 0 else "▼"
                ws3.append([m["name"], "整体", arrow, f"{d:+.2f}", f"p={t['p']}",
                            conclusions.get(m["source_col"], "")])
    for q in findings["questions"]:
        for t in q["adjacent_option_tests"]:
            if t.get("drift") and t["to"] == buckets[-1]:
                any_drift = True
                arrow = "▲" if t["delta_pp"] > 0 else "▼"
                ws3.append([q["question"], t["option"], arrow,
                            f"{t['delta_pp']:+.1f}pp", f"p={t['p']}",
                            conclusions.get(q["question"], "")])
    if not any_drift:
        ws3.append(["本期各指标/题目均无显著异动", "", "", "", "", ""])
    st.format_basic_stats_sheet(ws3, index_cols=2)

    # ---- Sheet 4: 方法与样本 ----
    ws4 = wb.create_sheet("ℹ️ 方法与样本")
    ws4.append(["项", "说明"])
    ws4.append(["分桶粒度", {"week": "按周", "month": "按月", "day": "按天"}.get(findings["granularity"])])
    ws4.append(["时间列", findings["time_col"]])
    ws4.append(["各桶样本量", "; ".join(f"{b}={sizes.get(b,0)}" for b in buckets)])
    ws4.append(["样本不足桶(n<30)", "; ".join(findings["low_n_buckets"]) or "无"])
    ws4.append(["判异动门槛", "p<0.05 且（占比Δ≥5pp 或 均分Δ≥0.1）"])
    ws4.append(["检验方法", "均分:t检验/Mann-Whitney; 占比:两比例z; 单选整体:卡方"])
    ws4.append(["免责", "样本不足桶仅供参考，不判异动"])
    st.format_basic_stats_sheet(ws4, index_cols=1)

    wb.save(output_path)
    return {"status": "success", "output_path": output_path, "sheets": wb.sheetnames}


def _format_detail_sheet(ws, n_buckets):
    import _styles as st
    ws.freeze_panes = "C2"
    ws.sheet_view.showGridLines = False
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = st.header_fill()
        cell.font = st.header_font(size=10)
        cell.alignment = st.ALIGN_CENTER
    # 占比列百分比格式
    for r in range(2, ws.max_row + 1):
        for c in range(3, 3 + n_buckets):
            ws.cell(row=r, column=c).number_format = "0.0%"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions[st_get_letter(2 + n_buckets + 4)].width = 40  # AI结论列


def st_get_letter(idx):
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx)
```

替换 Task 6 里的 `_cmd_export` 占位为：
```python
def _cmd_export(args):
    with open(args.findings, encoding="utf-8") as f:
        findings = json.load(f)
    conclusions = None
    if args.conclusions:
        with open(args.conclusions, encoding="utf-8") as f:
            conclusions = json.load(f)
    out = args.output_path or os.path.join(
        os.path.dirname(os.path.abspath(args.findings)),
        default_output_filename(findings["granularity"]))
    return export_excel(findings, conclusions, out)
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_survey_drift.py -k "export or default_output" -q`
Expected: PASS

- [ ] **Step 5: Checkpoint** — `python -m pytest tests/test_survey_drift.py -q` 全绿。

---

## Task 8: reference 文档 + SKILL.md 集成

**Files:**
- Create: `references/18-drift-workflow.md`
- Modify: `SKILL.md`

- [ ] **Step 1: 创建 `references/18-drift-workflow.md`**

内容包含以下小节（完整写出，非占位）：
1. **何时读取本文档 / 触发条件**（"按周/月/天诊断回流异动"等，见 spec 第9节触发词）
2. **前置**：输入为含时间列的量化 CSV；粒度由用户指定（周/月/天），未指定用 `ask_user_question` 三选一
3. **三步编排（一句话触发，中途不停）**：
   - Step A 运行 `python {SKILL_DIR}/scripts/survey_drift.py analyze --file_path "量化CSV" --granularity week --findings_out "…/drift_findings.json"`；读 stdout，若 `status=need_input` 按 `reason` 用 `ask_user_question` 补 `--time_col`/`--nps_col`/`--satisfaction_cols` 后重跑
   - Step B 读 `drift_findings.json`，**逐题**写一句话结论 → `conclusions.json`（`{题目: 结论}`）。结论规范：只写定性判断（数字脚本已算）；异动题写"哪个选项/指标+方向+幅度pp/分+是否显著"；无异动写"本期无显著变化"；`low_n=true` 的题结论末尾加"（样本不足，仅供参考）"
   - Step C 运行 `python {SKILL_DIR}/scripts/survey_drift.py export --findings "…/drift_findings.json" --conclusions "…/conclusions.json"`
4. **findings.json / conclusions.json 结构**（引用 spec 第5节 + 结论映射示例）
5. **Excel 4 Sheet 说明**（引用 spec 第7节）
6. **统计方法与双门槛/样本守卫**（引用 spec 第6节）
7. **错误处理**：`need_input`（缺时间列/识别不到指标）→ 追问；桶数<2 → 提示"当前时间跨度不足以分期，请扩大数据范围或换更细粒度"
8. **后续操作提示**：换粒度重跑、补充文本新增反馈检测（下一阶段）、下载最新回流数据

- [ ] **Step 2: 修改 `SKILL.md` — 脚本路径清单**

在脚本路径代码块（`enrich_columns.py` 行后）追加一行：
```
{SKILL_DIR}/scripts/survey_drift.py
```

- [ ] **Step 3: 修改 `SKILL.md` — 依赖要求**

把依赖行 `pip install pandas numpy openpyxl requests` 改为：
```
pip install pandas numpy scipy openpyxl requests
```

- [ ] **Step 4: 修改 `SKILL.md` — 新增阶段6**

在"### 阶段 5：生成报告"小节之后、"---" 之前，插入：
```
### 阶段 6：时间异动诊断（按需）

**触发条件**：用户有单份含时间列的回流问卷数据，想按周/月/天自动对比、
诊断满意度/NPS/单选/多选的异动（如"按周诊断这份回流数据的变化"、
"逐题对比各月满意度和NPS有没有显著变化"、"回流数据有没有异常波动"）。

→ **读取 `references/18-drift-workflow.md` 获取完整执行步骤。**
```

- [ ] **Step 5: 修改 `SKILL.md` — 后续操作提示 + 方法文档表**

在"📊 分析方面"选项列表中追加一条：
```
• 做时间异动诊断（按周/月/天对比满意度/NPS/单选/多选，定位显著变化并写初步结论）
```
并在"分析方法参考文档"表格追加一行：
```
| `references/18-drift-workflow.md` | 单份回流问卷按时间分桶的异动诊断流程（阶段 6） |
```

- [ ] **Step 6: Checkpoint** — 人工检查 SKILL.md 渲染正常、无破坏既有结构。

---

## Task 9: 端到端 smoke（analyze → 伪结论 → export）

**Files:**
- Test: `tests/test_survey_drift.py`

- [ ] **Step 1: 写失败测试**

追加：
```python
def test_full_pipeline_smoke(tmp_path):
    csv = tmp_path / "reflow.csv"
    pd.DataFrame({
        "结束答题时间": (["2026-04-06 10:00"] * 50 + ["2026-04-13 10:00"] * 50),
        "Q1.整体满意度": [5] * 50 + [3] * 50,
        "Q51.您有多大可能将本游戏推荐给朋友？": [10] * 50 + [4] * 50,
        "Q7.活动评价（单选）": (["满意"] * 50) + (["满意"] * 25 + ["一般"] * 25),
    }).to_csv(csv, index=False, encoding="utf-8-sig")

    findings_out = tmp_path / "drift_findings.json"
    r1 = subprocess.run(
        [sys.executable, str(MODULE_PATH), "analyze",
         "--file_path", str(csv), "--granularity", "week",
         "--findings_out", str(findings_out)],
        capture_output=True, text=True, encoding="utf-8")
    assert r1.returncode == 0, r1.stderr
    p1 = json.loads(r1.stdout)
    assert p1["status"] == "success"
    assert p1["questions_with_drift"] >= 1

    # 伪结论（模拟 Agent 写结论）
    findings = json.loads(findings_out.read_text(encoding="utf-8"))
    conclusions = {q["question"]: "自动结论" for q in findings["questions"]}
    conclusions_out = tmp_path / "conclusions.json"
    conclusions_out.write_text(json.dumps(conclusions, ensure_ascii=False), encoding="utf-8")

    out_xlsx = tmp_path / "report.xlsx"
    r2 = subprocess.run(
        [sys.executable, str(MODULE_PATH), "export",
         "--findings", str(findings_out), "--conclusions", str(conclusions_out),
         "--output_path", str(out_xlsx)],
        capture_output=True, text=True, encoding="utf-8")
    assert r2.returncode == 0, r2.stderr
    assert json.loads(r2.stdout)["status"] == "success"
    assert out_xlsx.exists()
```

- [ ] **Step 2: 跑测试确认失败/通过**

Run: `python -m pytest tests/test_survey_drift.py -k "smoke" -q`
Expected: 若前序任务完整，应 PASS；若失败，按报错定位缺口修复。

- [ ] **Step 3: 跑全量测试**

Run: `python -m pytest tests/test_survey_drift.py -q`
Expected: 全部 PASS

- [ ] **Step 4: Checkpoint** — 全绿即完成本计划。

---

## 自检（写计划后）

**Spec 覆盖：**
- 分桶周/月/天 → Task 2 ✅
- 相邻期对比 + 全时间线展示 → Task 5/6/7 ✅
- 满意度 t 检验、NPS 两比例 z、单选卡方+z、多选 z → Task 3/4/5/6 ✅
- 双门槛判异动 + n<30 守卫 → Task 3/5/6 ✅
- LLM 逐题结论注入 Excel → Task 7（AI结论列）+ Task 8（工作流 Step B）✅
- 4-Sheet Excel → Task 7 ✅
- analyze/export 双子命令 + need_input → Task 6 ✅
- scipy 依赖 → Task 1 ✅
- reference 18 + SKILL.md 阶段6 → Task 8 ✅
- 范围外（文本新增反馈/FDR/HTML）→ 计划未含，符合 spec 第10节 ✅

**占位扫描：** 无 TODO/TBD；每个代码步骤含完整代码。`_cmd_export` 占位在 Task 6 明确说明并于 Task 7 替换。

**类型一致性：** 各函数签名与"函数签名契约"一致；`by_bucket`/`sizes`/`ordered`/`adjacent`/`adjacent_option_tests` 字段跨 Task 命名统一；Excel Sheet 名四处一致。
