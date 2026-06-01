# survey_compare 多期问卷对比工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增 `survey_compare.py` 脚本，支持两期（或多期）问卷数据自动匹配 + 对比，输出 4-Sheet 对比 Excel，并沉淀进 survey-research skill 供月度复用。

**Architecture:** 单文件 CLI 脚本，沿用现有 `_styles.py` 样式体系和 `_detect_csv_encoding` 辅助函数惯例；题目匹配用 `difflib.SequenceMatcher`；Excel 生成用 `openpyxl`；结果通过 stdout JSON 返回。不依赖 basic_stats/crosstab，独立运行。

**Tech Stack:** Python 3.x, pandas, openpyxl, difflib（标准库）, json, argparse

---

## 文件清单

| 操作 | 路径 |
|------|------|
| **新建** | `scripts/survey_compare.py` |
| **新建** | `references/16-compare-workflow.md` |
| **修改** | `README.md`（skill 入口文档，追加多期对比触发条件） |

> skill 路径前缀：`C:\Users\lijinghui03\.agents\skills\survey-research\`

---

## Task 1：脚本骨架 + CLI + 数据加载

**Files:**
- Create: `scripts/survey_compare.py`

- [ ] **Step 1.1：创建脚本文件，写入文件头和 CLI 入口**

```python
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 多期对比分析
============================

对比两期（或多期）问卷量化数据，生成四个 Sheet 的对比 Excel：
  1. 指标总览  —— 关键量表/NPS 均分趋势
  2. 逐题对比  —— 每题各选项占比 + Δ 差值
  3. 人群结构  —— 人口学/行为变量分布对比
  4. 文本主题  —— 文本分析主题变化（需传入 --text_results）

用法:
    python survey_compare.py \\
        --files "survey_A.csv" "survey_B.csv" \\
        --labels "S21飞龙" "S20X武器" \\
        [--mapping "compare_map.json"] \\
        [--text_results "text_A.json" "text_B.json"] \\
        [--output_path "对比报告.xlsx"]
"""

import argparse
import json
import sys
import os
import re
import difflib
from collections import defaultdict, OrderedDict
from datetime import datetime
from typing import Optional, List, Dict, Tuple

import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _styles import (
    Theme, thin_border, header_fill, header_font,
    body_font, even_fill, odd_fill, make_fill,
    ALIGN_CENTER, ALIGN_LEFT, ALIGN_RIGHT,
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ========================================================================= #
#                         CLI 入口
# ========================================================================= #

def _parse_args():
    p = argparse.ArgumentParser(description="多期问卷对比分析")
    p.add_argument("--files", nargs="+", required=True,
                   help="2~N 份量化数据 CSV，按新→旧排列")
    p.add_argument("--labels", nargs="+", required=True,
                   help="各期标签，与 --files 一一对应")
    p.add_argument("--mapping", default=None,
                   help="手动题目映射 JSON 文件路径")
    p.add_argument("--text_results", nargs="+", default=None,
                   help="各期文本分析 JSON 文件路径，与 --files 一一对应")
    p.add_argument("--output_path", default=None,
                   help="输出 Excel 路径，默认与第一个 CSV 同目录")
    return p.parse_args()


def main():
    args = _parse_args()

    if len(args.files) < 2:
        _err("至少需要 2 份数据文件")
    if len(args.labels) != len(args.files):
        _err(f"--labels 数量({len(args.labels)}) 与 --files({len(args.files)}) 不一致")
    if args.text_results and len(args.text_results) != len(args.files):
        _err(f"--text_results 数量({len(args.text_results)}) 与 --files({len(args.files)}) 不一致")

    # 加载数据
    dfs = []
    for f in args.files:
        enc = _detect_csv_encoding(f)
        df = pd.read_csv(f, encoding=enc, low_memory=False)
        dfs.append(df)

    # 加载手动映射
    manual_pairs, exclude_cols = _load_mapping(args.mapping)

    # 确定输出路径
    output_path = args.output_path or _default_output(args.files[0])

    # 匹配题目
    matched, a_only, b_only = _match_questions(dfs, args.labels, manual_pairs, exclude_cols)

    # 加载文本分析结果
    text_results_list = []
    if args.text_results:
        for tr_path in args.text_results:
            try:
                with open(tr_path, encoding="utf-8") as f:
                    text_results_list.append(json.load(f))
            except Exception:
                text_results_list.append([])

    # 生成 Excel
    wb = Workbook()
    wb.remove(wb.active)  # 删除默认 Sheet

    _write_overview_sheet(wb, dfs, args.labels, matched)
    _write_question_compare_sheet(wb, dfs, args.labels, matched, a_only, b_only)
    _write_population_sheet(wb, dfs, args.labels)
    if text_results_list:
        _write_text_compare_sheet(wb, text_results_list, args.labels)

    wb.save(output_path)

    result = {
        "status": "success",
        "output_path": output_path,
        "matched_questions": len(matched),
        "unmatched_a_only": len(a_only),
        "unmatched_b_only": len(b_only),
        "sheets": [s.title for s in wb.worksheets],
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
```

- [ ] **Step 1.2：添加辅助函数（编码检测、错误输出、路径生成）**

在 `main()` 函数上方（`_parse_args` 之前）插入：

```python
# ========================================================================= #
#                         辅助工具函数
# ========================================================================= #

def _err(msg: str):
    print(json.dumps({"status": "error", "message": msg}, ensure_ascii=False), file=sys.stderr)
    sys.exit(1)


def _detect_csv_encoding(filepath, sample_size=8192) -> str:
    """检测 CSV 文件编码"""
    with open(filepath, 'rb') as f:
        raw = f.read(sample_size)
    if raw.startswith(b'\xef\xbb\xbf'):
        return 'utf-8-sig'
    try:
        raw.decode('utf-8')
        return 'utf-8'
    except UnicodeDecodeError:
        return 'gbk'


def _default_output(first_csv: str) -> str:
    """生成默认输出路径"""
    dir_ = os.path.dirname(os.path.abspath(first_csv))
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    return os.path.join(dir_, f"survey_compare_{ts}.xlsx")


def _strip_q_prefix(col: str) -> str:
    """去除 Qxx. 前缀、（非必填）等修饰词，用于相似度比较"""
    s = re.sub(r'^[QY]\d+[.\s]+', '', str(col))
    s = re.sub(r'[（(][^）)]*[）)]', '', s)
    s = s.strip()
    return s


RE_Q_ROOT = re.compile(r'^([QY]\d+)\.')


def _q_root(col: str) -> str:
    """提取 Q 编号根，如 Q12"""
    m = RE_Q_ROOT.match(str(col))
    return m.group(1) if m else ""
```

- [ ] **Step 1.3：运行脚本检查 import 无报错**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research\scripts
python survey_compare.py --help
```

预期输出：显示 argparse 帮助文本，无 ImportError。

---

## Task 2：题目匹配引擎

**Files:**
- Modify: `scripts/survey_compare.py`（在 `_strip_q_prefix` 之后插入）

- [ ] **Step 2.1：添加手动映射加载函数**

```python
def _load_mapping(mapping_path: Optional[str]) -> Tuple[List[Dict], List[str]]:
    """
    加载 compare_map.json。
    返回 (manual_pairs, exclude_cols)
    manual_pairs 格式: [{"label": "...", "a": "列名A", "b": "列名B"}, ...]
    """
    if not mapping_path or not os.path.exists(mapping_path):
        return [], []
    with open(mapping_path, encoding="utf-8") as f:
        cfg = json.load(f)
    pairs = cfg.get("manual_pairs", [])
    excl = cfg.get("exclude", [])
    return pairs, excl
```

- [ ] **Step 2.2：添加自动题目匹配函数**

```python
def _match_questions(
    dfs: List[pd.DataFrame],
    labels: List[str],
    manual_pairs: List[Dict],
    exclude_cols: List[str],
) -> Tuple[List[Dict], List[str], List[str]]:
    """
    对 dfs[0]（A 期）和 dfs[1]（B 期）的列名做题目匹配。
    
    返回:
      matched  : [{"label": "题目标签", "a_col": "...", "b_col": "...", "score": 0.95}, ...]
      a_only   : 仅在 A 期出现的列名列表
      b_only   : 仅在 B 期出现的列名列表
    
    匹配优先级：手动映射 > 自动相似度匹配（阈值 0.70）
    """
    df_a, df_b = dfs[0], dfs[1]

    # 排除 exclude 列
    exclude_set = set(exclude_cols)

    # 获取有效列（Q 开头）
    cols_a = [c for c in df_a.columns if RE_Q_ROOT.match(str(c)) and c not in exclude_set]
    cols_b = [c for c in df_b.columns if RE_Q_ROOT.match(str(c)) and c not in exclude_set]

    # 手动映射：建立精确配对字典
    manual_a_used = set()
    manual_b_used = set()
    matched = []

    for mp in manual_pairs:
        a_col = mp.get("a", "")
        b_col = mp.get("b", "")
        lbl = mp.get("label", _strip_q_prefix(a_col) or _strip_q_prefix(b_col))
        if a_col in df_a.columns and b_col in df_b.columns:
            matched.append({"label": lbl, "a_col": a_col, "b_col": b_col, "score": 1.0, "manual": True})
            manual_a_used.add(a_col)
            manual_b_used.add(b_col)

    # 自动匹配：对未被手动配对的列做相似度计算
    remaining_a = [c for c in cols_a if c not in manual_a_used]
    remaining_b = [c for c in cols_b if c not in manual_b_used]

    # 构建 b 的 stripped → col 映射
    b_stripped = {_strip_q_prefix(c): c for c in remaining_b}

    auto_matched_b = set()
    for a_col in remaining_a:
        a_stripped = _strip_q_prefix(a_col)
        if not a_stripped:
            continue
        best_score = 0.0
        best_b_col = None
        for b_s, b_col in b_stripped.items():
            if b_col in auto_matched_b:
                continue
            score = difflib.SequenceMatcher(None, a_stripped, b_s).ratio()
            if score > best_score:
                best_score = score
                best_b_col = b_col
        if best_score >= 0.70 and best_b_col:
            lbl = a_stripped
            matched.append({"label": lbl, "a_col": a_col, "b_col": best_b_col, "score": best_score, "manual": False})
            auto_matched_b.add(best_b_col)

    # 剩余未匹配
    matched_a_cols = {m["a_col"] for m in matched}
    matched_b_cols = {m["b_col"] for m in matched}
    a_only = [c for c in cols_a if c not in matched_a_cols]
    b_only = [c for c in cols_b if c not in matched_b_cols]

    return matched, a_only, b_only
```

- [ ] **Step 2.3：快速验证匹配逻辑（命令行测试）**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research\scripts
python -c "
import sys; sys.path.insert(0,'.')
from survey_compare import _strip_q_prefix, _match_questions
import pandas as pd
df_a = pd.DataFrame(columns=['Q1.请问整体满意度如何？（单选）','Q2.您的性别是？（单选）'])
df_b = pd.DataFrame(columns=['Q1.请问整体满意度如何？（单选）','Q3.您的性别是？（单选）'])
matched, a_only, b_only = _match_questions([df_a, df_b], ['A','B'], [], [])
print('matched:', len(matched), [m['label'] for m in matched])
print('a_only:', a_only)
print('b_only:', b_only)
"
```

预期：matched=2，a_only=[]，b_only=[]

- [ ] **Step 2.4：Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py
git commit -m "feat: survey_compare - skeleton + question matching engine"
```

---

## Task 3：指标总览 Sheet（Sheet 1）

**Files:**
- Modify: `scripts/survey_compare.py`

指标总览自动识别 **量表题**（列名含"满意度"/"评分"/"推荐"或均值可计算）和 **NPS 题**（0-10分单选），计算均分和 Δ。

- [ ] **Step 3.1：添加量表题识别和均分计算辅助函数**

```python
# ========================================================================= #
#                         指标计算
# ========================================================================= #

def _is_scale_col(col: str, df: pd.DataFrame) -> bool:
    """判断是否为可计算均分的量表题"""
    kw = ["满意度", "评分", "满意", "推荐", "评价", "体验"]
    if not any(k in str(col) for k in kw):
        return False
    vals = df[col].dropna()
    numeric_vals = pd.to_numeric(vals, errors='coerce').dropna()
    return len(numeric_vals) / max(len(vals), 1) >= 0.5


def _is_nps_col(col: str, df: pd.DataFrame) -> bool:
    """判断是否为 NPS 题（0-10 分单选）"""
    if "推荐" not in str(col):
        return False
    vals = pd.to_numeric(df[col].dropna(), errors='coerce').dropna()
    return vals.between(0, 10).all() and len(vals) > 0


def _calc_mean(df: pd.DataFrame, col: str) -> Optional[float]:
    """计算列的数值均分，处理如'1星'/'非常满意(1)'等形式"""
    vals = df[col].dropna().astype(str)
    nums = []
    for v in vals:
        m = re.search(r'(\d+(?:\.\d+)?)', v)
        if m:
            nums.append(float(m.group(1)))
    return float(np.mean(nums)) if nums else None


def _calc_nps(df: pd.DataFrame, col: str) -> Optional[float]:
    """计算 NPS 分数 = 推荐型(9-10)占比 - 批评型(0-6)占比"""
    vals = pd.to_numeric(df[col].dropna(), errors='coerce').dropna()
    total = len(vals)
    if total == 0:
        return None
    promoters = (vals >= 9).sum() / total * 100
    detractors = (vals <= 6).sum() / total * 100
    return round(promoters - detractors, 1)


def _trend_label(delta: float, is_nps: bool = False) -> Tuple[str, str]:
    """
    根据差值返回趋势文字和颜色代码。
    返回 (text, color_hex)
    """
    threshold_small = 0.1 if not is_nps else 2.0
    threshold_big = 0.5 if not is_nps else 15.0
    if abs(delta) < threshold_small:
        return "— 持平", "666666"
    if delta > 0:
        if delta >= threshold_big:
            return f"▲▲ +{delta:.1f}{'pp' if is_nps else ''}", "375623"
        return f"▲ +{delta:.1f}{'pp' if is_nps else ''}", "375623"
    else:
        if abs(delta) >= threshold_big:
            return f"▼▼ {delta:.1f}{'pp' if is_nps else ''}", "C00000"
        return f"▼ {delta:.1f}{'pp' if is_nps else ''}", "C00000"
```

- [ ] **Step 3.2：添加写入指标总览 Sheet 函数**

```python
def _write_overview_sheet(
    wb: Workbook,
    dfs: List[pd.DataFrame],
    labels: List[str],
    matched: List[Dict],
):
    """Sheet 1：指标总览"""
    ws = wb.create_sheet("📊 指标总览")
    df_a, df_b = dfs[0], dfs[1]
    label_a, label_b = labels[0], labels[1]

    # 样式常量
    FONT_NAME = Theme.FONT_NAME
    hdr_fill = PatternFill("solid", fgColor=Theme.HEADER_BG)
    hdr_font = Font(name=FONT_NAME, bold=True, color=Theme.HEADER_FONT, size=11)
    idx_fill = PatternFill("solid", fgColor=Theme.INDEX_BG)
    idx_font = Font(name=FONT_NAME, bold=True, color=Theme.INDEX_FONT, size=10)
    body_font_ = Font(name=FONT_NAME, size=10)
    total_fill_ = PatternFill("solid", fgColor=Theme.TOTAL_BG)
    total_font_ = Font(name=FONT_NAME, bold=True, color=Theme.TOTAL_FONT, size=10)
    border = thin_border()

    def _write_cell(row, col, val, fill=None, font=None, align=None, bold=False):
        cell = ws.cell(row=row, column=col, value=val)
        if fill:
            cell.fill = fill
        if font:
            cell.font = font
        elif bold:
            cell.font = Font(name=FONT_NAME, bold=True, size=10)
        else:
            cell.font = body_font_
        cell.alignment = align or ALIGN_CENTER
        cell.border = border
        return cell

    # 标题行
    row = 1
    headers = ["指标名", label_a, label_b, "趋势", "备注"]
    col_widths = [35, 15, 15, 20, 25]
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        _write_cell(row, ci, h, fill=hdr_fill, font=hdr_font)
        ws.column_dimensions[get_column_letter(ci)].width = w

    # 样本量行
    row += 1
    _write_cell(row, 1, "样本量", fill=idx_fill, font=idx_font, align=ALIGN_LEFT)
    _write_cell(row, 2, len(df_a))
    _write_cell(row, 3, len(df_b))
    delta_n = len(df_a) - len(df_b)
    trend_txt, trend_color = _trend_label(delta_n / max(len(df_b), 1) * 100, is_nps=True)
    tc = ws.cell(row=row, column=4, value=trend_txt)
    tc.font = Font(name=FONT_NAME, bold=True, color=trend_color, size=10)
    tc.alignment = ALIGN_CENTER
    tc.border = border
    _write_cell(row, 5, "")

    # 量表/NPS 指标行
    seen_labels = set()
    for m in matched:
        a_col, b_col, lbl = m["a_col"], m["b_col"], m["label"]
        if lbl in seen_labels:
            continue

        is_nps = _is_nps_col(a_col, df_a) or _is_nps_col(b_col, df_b)
        is_scale = _is_scale_col(a_col, df_a) or _is_scale_col(b_col, df_b)
        if not (is_nps or is_scale):
            continue

        row += 1
        seen_labels.add(lbl)
        _write_cell(row, 1, lbl, fill=idx_fill, font=idx_font, align=ALIGN_LEFT)

        if is_nps:
            val_a = _calc_nps(df_a, a_col)
            val_b = _calc_nps(df_b, b_col)
            unit = "（NPS分）"
        else:
            val_a = _calc_mean(df_a, a_col)
            val_b = _calc_mean(df_b, b_col)
            unit = "（均分）"

        val_a_disp = round(val_a, 2) if val_a is not None else "—"
        val_b_disp = round(val_b, 2) if val_b is not None else "—"
        _write_cell(row, 2, val_a_disp)
        _write_cell(row, 3, val_b_disp)

        if val_a is not None and val_b is not None:
            delta = val_a - val_b
            trend_txt, trend_color = _trend_label(delta, is_nps=is_nps)
            tc = ws.cell(row=row, column=4, value=trend_txt)
            tc.font = Font(name=FONT_NAME, bold=True, color=trend_color, size=10)
            tc.alignment = ALIGN_CENTER
            tc.border = border
            note = f"{label_b}→{label_a}: {'+' if delta>=0 else ''}{round(delta,2)}{unit}"
        else:
            _write_cell(row, 4, "数据不足")
            note = ""
        _write_cell(row, 5, note, align=ALIGN_LEFT)

    ws.freeze_panes = "A2"
```

- [ ] **Step 3.3：Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py
git commit -m "feat: survey_compare - overview sheet with scale/NPS metrics"
```

---

## Task 4：逐题对比 Sheet（Sheet 2）

**Files:**
- Modify: `scripts/survey_compare.py`

每道匹配题目的各选项占比对比，Δ > 5pp 高亮。

- [ ] **Step 4.1：添加单题占比计算辅助函数**

```python
def _get_option_pcts(df: pd.DataFrame, col: str) -> Dict[str, float]:
    """
    计算单选/量表题各选项占比。
    返回 {选项文本: 占比(0~1)}，按出现频率排序。
    """
    vals = df[col].dropna().astype(str).str.strip()
    vals = vals[vals.str.len() > 0]
    total = len(vals)
    if total == 0:
        return {}
    counts = vals.value_counts()
    return {opt: cnt / total for opt, cnt in counts.items()}
```

- [ ] **Step 4.2：添加写入逐题对比 Sheet 函数**

```python
def _write_question_compare_sheet(
    wb: Workbook,
    dfs: List[pd.DataFrame],
    labels: List[str],
    matched: List[Dict],
    a_only: List[str],
    b_only: List[str],
):
    """Sheet 2：逐题对比"""
    ws = wb.create_sheet("📋 逐题对比")
    df_a, df_b = dfs[0], dfs[1]
    label_a, label_b = labels[0], labels[1]
    FONT_NAME = Theme.FONT_NAME

    hdr_fill = PatternFill("solid", fgColor=Theme.HEADER_BG)
    hdr_font = Font(name=FONT_NAME, bold=True, color=Theme.HEADER_FONT, size=11)
    idx_fill = PatternFill("solid", fgColor=Theme.INDEX_BG)
    idx_font = Font(name=FONT_NAME, bold=True, color=Theme.INDEX_FONT, size=10)
    body_font_ = Font(name=FONT_NAME, size=10)
    pos_fill = PatternFill("solid", fgColor="E2EFDA")  # 正向变化
    neg_fill = PatternFill("solid", fgColor="FCE4EC")  # 负向变化
    border = thin_border()
    pos_font = Font(name=FONT_NAME, bold=True, color="375623", size=10)
    neg_font = Font(name=FONT_NAME, bold=True, color="C00000", size=10)
    gray_font = Font(name=FONT_NAME, color="666666", size=10)

    col_widths = [38, 22, 14, 14, 12, 10]
    headers = ["题目", "选项", label_a, label_b, "Δ (A-B)", "显著"]
    row = 1
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font; c.alignment = ALIGN_CENTER; c.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w

    # 每道匹配题
    for m in matched:
        a_col, b_col, lbl = m["a_col"], m["b_col"], m["label"]
        pcts_a = _get_option_pcts(df_a, a_col)
        pcts_b = _get_option_pcts(df_b, b_col)
        all_opts = list(dict.fromkeys(list(pcts_a.keys()) + list(pcts_b.keys())))

        first_opt = True
        for opt in all_opts:
            row += 1
            va = pcts_a.get(opt, 0.0)
            vb = pcts_b.get(opt, 0.0)
            delta = va - vb
            sig = "✅" if abs(delta) >= 0.05 else ""

            # 题目名（仅第一行显示）
            q_cell = ws.cell(row=row, column=1, value=lbl if first_opt else "")
            q_cell.fill = idx_fill; q_cell.font = idx_font
            q_cell.alignment = ALIGN_LEFT; q_cell.border = border
            first_opt = False

            # 选项
            opt_cell = ws.cell(row=row, column=2, value=opt)
            opt_cell.font = body_font_; opt_cell.alignment = ALIGN_LEFT; opt_cell.border = border

            # 两期占比
            for ci, val in [(3, va), (4, vb)]:
                c = ws.cell(row=row, column=ci, value=f"{val:.1%}")
                c.font = body_font_; c.alignment = ALIGN_CENTER; c.border = border

            # Δ
            delta_cell = ws.cell(row=row, column=5, value=f"{delta:+.1%}")
            if abs(delta) >= 0.05:
                delta_cell.fill = pos_fill if delta > 0 else neg_fill
                delta_cell.font = pos_font if delta > 0 else neg_font
            else:
                delta_cell.font = gray_font
            delta_cell.alignment = ALIGN_CENTER; delta_cell.border = border

            # 显著标记
            sig_cell = ws.cell(row=row, column=6, value=sig)
            sig_cell.font = body_font_; sig_cell.alignment = ALIGN_CENTER; sig_cell.border = border

        # 题目间空行
        row += 1

    # 仅出现在 A/B 的题目
    if a_only or b_only:
        row += 1
        section_cell = ws.cell(row=row, column=1, value="── 未匹配题目 ──")
        section_cell.font = Font(name=FONT_NAME, bold=True, color="666666", size=11)
        section_cell.alignment = ALIGN_LEFT

        for col_name in a_only:
            row += 1
            c = ws.cell(row=row, column=1, value=f"[仅{labels[0]}] {col_name}")
            c.font = Font(name=FONT_NAME, italic=True, color="666666", size=10)
            c.alignment = ALIGN_LEFT

        for col_name in b_only:
            row += 1
            c = ws.cell(row=row, column=1, value=f"[仅{labels[1]}] {col_name}")
            c.font = Font(name=FONT_NAME, italic=True, color="666666", size=10)
            c.alignment = ALIGN_LEFT

    ws.freeze_panes = "A2"
```

- [ ] **Step 4.3：Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py
git commit -m "feat: survey_compare - question-by-question comparison sheet"
```

---

## Task 5：人群结构对比 Sheet（Sheet 3）

**Files:**
- Modify: `scripts/survey_compare.py`

- [ ] **Step 5.1：添加人口学维度识别和写入函数**

```python
# ========================================================================= #
#                         人群结构 Sheet
# ========================================================================= #

_DEMO_KEYWORDS = ["性别", "年龄", "职业", "段位", "游玩情况", "玩家类型", "付费"]


def _find_demo_cols(df: pd.DataFrame) -> List[str]:
    """找出人口学/行为特征列"""
    result = []
    for col in df.columns:
        if any(kw in str(col) for kw in _DEMO_KEYWORDS):
            if RE_Q_ROOT.match(str(col)):
                result.append(col)
    return result


def _write_population_sheet(
    wb: Workbook,
    dfs: List[pd.DataFrame],
    labels: List[str],
):
    """Sheet 3：人群结构对比"""
    ws = wb.create_sheet("👥 人群结构")
    df_a, df_b = dfs[0], dfs[1]
    label_a, label_b = labels[0], labels[1]
    FONT_NAME = Theme.FONT_NAME

    hdr_fill = PatternFill("solid", fgColor=Theme.HEADER_BG)
    hdr_font = Font(name=FONT_NAME, bold=True, color=Theme.HEADER_FONT, size=11)
    idx_fill = PatternFill("solid", fgColor=Theme.INDEX_BG)
    idx_font = Font(name=FONT_NAME, bold=True, color=Theme.INDEX_FONT, size=10)
    body_font_ = Font(name=FONT_NAME, size=10)
    border = thin_border()
    pos_font = Font(name=FONT_NAME, bold=True, color="375623", size=10)
    neg_font = Font(name=FONT_NAME, bold=True, color="C00000", size=10)
    gray_font = Font(name=FONT_NAME, color="666666", size=10)

    col_widths = [35, 22, 14, 14, 12]
    headers = ["维度", "选项", label_a, label_b, "Δ (A-B)"]
    row = 1
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font; c.alignment = ALIGN_CENTER; c.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w

    demo_a = _find_demo_cols(df_a)
    demo_b = _find_demo_cols(df_b)

    # 尝试按关键词对应两期列
    used_b = set()
    for a_col in demo_a:
        stripped_a = _strip_q_prefix(a_col)
        # 找 B 中最接近的列
        best_b = None
        best_score = 0.0
        for b_col in demo_b:
            if b_col in used_b:
                continue
            score = difflib.SequenceMatcher(None, stripped_a, _strip_q_prefix(b_col)).ratio()
            if score > best_score:
                best_score = score
                best_b = b_col
        if best_b is None or best_score < 0.50:
            continue
        used_b.add(best_b)

        pcts_a = _get_option_pcts(df_a, a_col)
        pcts_b = _get_option_pcts(df_b, best_b)
        all_opts = list(dict.fromkeys(list(pcts_a.keys()) + list(pcts_b.keys())))
        dim_label = stripped_a

        first_opt = True
        for opt in all_opts:
            row += 1
            va = pcts_a.get(opt, 0.0)
            vb = pcts_b.get(opt, 0.0)
            delta = va - vb

            dim_cell = ws.cell(row=row, column=1, value=dim_label if first_opt else "")
            dim_cell.fill = idx_fill; dim_cell.font = idx_font
            dim_cell.alignment = ALIGN_LEFT; dim_cell.border = border
            first_opt = False

            opt_cell = ws.cell(row=row, column=2, value=opt)
            opt_cell.font = body_font_; opt_cell.alignment = ALIGN_LEFT; opt_cell.border = border

            for ci, val in [(3, va), (4, vb)]:
                c = ws.cell(row=row, column=ci, value=f"{val:.1%}")
                c.font = body_font_; c.alignment = ALIGN_CENTER; c.border = border

            delta_cell = ws.cell(row=row, column=5, value=f"{delta:+.1%}")
            if abs(delta) >= 0.05:
                delta_cell.font = pos_font if delta > 0 else neg_font
            else:
                delta_cell.font = gray_font
            delta_cell.alignment = ALIGN_CENTER; delta_cell.border = border

        row += 1  # 维度间空行

    ws.freeze_panes = "A2"
```

- [ ] **Step 5.2：Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py
git commit -m "feat: survey_compare - population structure comparison sheet"
```

---

## Task 6：文本主题变化 Sheet（Sheet 4）

**Files:**
- Modify: `scripts/survey_compare.py`

- [ ] **Step 6.1：添加写入文本主题 Sheet 函数**

```python
# ========================================================================= #
#                         文本主题变化 Sheet
# ========================================================================= #

def _write_text_compare_sheet(
    wb: Workbook,
    text_results_list: List[List[Dict]],
    labels: List[str],
):
    """Sheet 4：文本主题变化（需要两期 text_results JSON）"""
    ws = wb.create_sheet("💬 文本主题")
    label_a, label_b = labels[0], labels[1]
    FONT_NAME = Theme.FONT_NAME

    hdr_fill = PatternFill("solid", fgColor=Theme.HEADER_BG)
    hdr_font = Font(name=FONT_NAME, bold=True, color=Theme.HEADER_FONT, size=11)
    idx_fill = PatternFill("solid", fgColor=Theme.INDEX_BG)
    idx_font = Font(name=FONT_NAME, bold=True, color=Theme.INDEX_FONT, size=10)
    q_title_fill = PatternFill("solid", fgColor="4472C4")
    q_title_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    body_font_ = Font(name=FONT_NAME, size=10)
    border = thin_border()
    pos_font = Font(name=FONT_NAME, bold=True, color="375623", size=10)
    neg_font = Font(name=FONT_NAME, bold=True, color="C00000", size=10)
    gray_font = Font(name=FONT_NAME, color="666666", size=10)

    col_widths = [35, 16, 16, 16]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # 按题目匹配两期文本结果
    results_a = text_results_list[0] if len(text_results_list) > 0 else []
    results_b = text_results_list[1] if len(text_results_list) > 1 else []

    # 建立 B 的 question → entry 映射
    b_map = {}
    for entry in results_b:
        b_map[_strip_q_prefix(entry["question"])] = entry

    row = 0
    for entry_a in results_a:
        q_label = _strip_q_prefix(entry_a["question"])
        entry_b = b_map.get(q_label)

        row += 1
        # 题目标题行
        q_cell = ws.cell(row=row, column=1, value=entry_a["question"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        q_cell.fill = q_title_fill; q_cell.font = q_title_font
        q_cell.alignment = ALIGN_LEFT; q_cell.border = border

        row += 1
        # 列头
        for ci, h in enumerate(["主题", f"{label_a}占比", f"{label_b}占比", "变化"], 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.fill = hdr_fill; c.font = hdr_font; c.alignment = ALIGN_CENTER; c.border = border

        dims_a = {d["name"]: d for d in entry_a.get("dimensions", [])}
        dims_b = {d["name"]: d for d in (entry_b.get("dimensions", []) if entry_b else [])}

        # 尝试匹配主题名
        all_dim_names = list(dims_a.keys())
        for dn in dims_b:
            if dn not in all_dim_names:
                all_dim_names.append(dn)

        for dim_name in all_dim_names:
            row += 1
            d_a = dims_a.get(dim_name)
            d_b = dims_b.get(dim_name)

            pct_a_str = d_a["percentage"] if d_a else "—"
            pct_b_str = d_b["percentage"] if d_b else "—"

            # 解析占比数值
            def _pct_val(s):
                m = re.search(r'(\d+\.?\d*)', str(s))
                return float(m.group(1)) / 100 if m else None

            va = _pct_val(pct_a_str)
            vb = _pct_val(pct_b_str)

            dim_cell = ws.cell(row=row, column=1, value=dim_name)
            dim_cell.font = body_font_; dim_cell.alignment = ALIGN_LEFT; dim_cell.border = border

            for ci, val_str in [(2, pct_a_str), (3, pct_b_str)]:
                c = ws.cell(row=row, column=ci, value=val_str)
                c.font = body_font_; c.alignment = ALIGN_CENTER; c.border = border

            # 变化列
            if va is not None and vb is not None:
                delta = va - vb
                trend_txt, trend_color = _trend_label(delta * 100, is_nps=True)
                tc = ws.cell(row=row, column=4, value=trend_txt)
                tc.font = Font(name=FONT_NAME, bold=True, color=trend_color, size=10)
            else:
                tc = ws.cell(row=row, column=4, value="本期未分析")
                tc.font = gray_font
            tc.alignment = ALIGN_CENTER; tc.border = border

        row += 2  # 题目间空行

    ws.freeze_panes = "A2"
```

- [ ] **Step 6.2：Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py
git commit -m "feat: survey_compare - text theme comparison sheet"
```

---

## Task 7：端到端测试（用真实数据验证）

**Files:**
- 使用已有数据：`C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\`

- [ ] **Step 7.1：不传 text_results 的基础运行**

```bash
python C:\Users\lijinghui03\.agents\skills\survey-research\scripts\survey_compare.py ^
  --files "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\survey_93048【量化数据】20260512-20260519_1779940072.csv" "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\survey_92034【量化数据】20260413-20260421_1779940110.csv" ^
  --labels "S21飞龙赛季" "S20X武器赛季" ^
  --output_path "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\survey_compare_test.xlsx"
```

预期 stdout JSON：`status: success`，`matched_questions > 20`，4 个 sheets。

- [ ] **Step 7.2：传入 text_results 的完整运行**

```bash
python C:\Users\lijinghui03\.agents\skills\survey-research\scripts\survey_compare.py ^
  --files "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\survey_93048【量化数据】20260512-20260519_1779940072.csv" "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\survey_92034【量化数据】20260413-20260421_1779940110.csv" ^
  --labels "S21飞龙赛季" "S20X武器赛季" ^
  --text_results "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\text_results_93048_fixed.json" ^
  --output_path "C:\Users\lijinghui03\Desktop\问卷交叉分析_20260528114508\survey_compare_93048_vs_92034.xlsx"
```

预期：Sheet "💬 文本主题" 存在，有 19 个题目对应的主题表。

- [ ] **Step 7.3：验证 Excel 可以正常打开，检查各 Sheet 数据是否正确**

手动打开 `survey_compare_93048_vs_92034.xlsx`，确认：
- Sheet 1 指标总览：有满意度均分行、趋势标识颜色正确
- Sheet 2 逐题对比：Δ ≥ 5pp 的行有颜色高亮
- Sheet 3 人群结构：有性别/年龄/段位分布对比
- Sheet 4 文本主题：Q2/Q3 等题目有主题变化表

- [ ] **Step 7.4：Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py
git commit -m "test: survey_compare - end-to-end validation with real data"
```

---

## Task 8：沉淀进 skill（reference + SKILL.md 更新）

**Files:**
- Create: `references/16-compare-workflow.md`
- Modify: `README.md`（skill 文档入口）

- [ ] **Step 8.1：创建 `references/16-compare-workflow.md`**

```markdown
# 多期问卷对比分析工作流程（阶段 6）

> 📌 **何时读取本文档**：用户有两期或以上问卷数据，想对比趋势变化时，由 SKILL.md 指引跳转至此。

## 触发条件

- 用户有两份不同时期的问卷数据（如两个赛季的满意度调研）
- 用户说"对比一下上两期数据"、"两份问卷的差异"、"这个月和上个月对比"
- 用户想沉淀成月度复用的对比报告

## 执行命令

```bash
python {SKILL_DIR}/scripts/survey_compare.py \
  --files "survey_A.csv" "survey_B.csv" \
  --labels "S21飞龙" "S20X武器" \
  [--mapping "compare_map.json"] \
  [--text_results "text_A.json" "text_B.json"] \
  --output_path "survey_compare_报告.xlsx"
```

**参数说明：**

| 参数 | 必填 | 说明 |
|------|------|------|
| `--files` | ✅ | 2~N 份量化 CSV，按**新→旧**排列 |
| `--labels` | ✅ | 各期标签，如 "S21飞龙" "S20X武器" |
| `--mapping` | ❌ | 手动题目映射 JSON（见下方格式） |
| `--text_results` | ❌ | 各期文本分析 JSON（text_results_xxx.json），对应 --files 顺序 |
| `--output_path` | ❌ | 默认与第一个 CSV 同目录 |

## 输出 Excel 结构

| Sheet | 内容 |
|-------|------|
| 📊 指标总览 | 样本量、满意度均分、NPS，含 ▲▼ 趋势标 |
| 📋 逐题对比 | 每道匹配题目的各选项占比 + Δ差值，≥5pp 高亮 |
| 👥 人群结构 | 性别/年龄/职业/段位分布两期对比 |
| 💬 文本主题 | 各文本题的主题维度占比对比（需传 --text_results） |

## 手动映射配置（compare_map.json）

适用于跨赛季同题但 Q 编号不同的情况（如 Q13"飞龙满意度" vs Q12"X武器满意度"）：

```json
{
  "manual_pairs": [
    {
      "label": "赛季满意度",
      "a": "Q13.总体而言，您对本赛季【飞龙】赛季的满意度如何？（单选）",
      "b": "Q12.总体而言，您对本赛季【X武器】赛季的满意度如何？（单选）"
    }
  ],
  "exclude": [
    "Q55.请问您的性别是？"
  ]
}
```

将 `compare_map.json` 保存在数据目录中，下月只需替换 `--files` 参数即可复用。

## 月度复用建议

1. 每月下载新问卷 CSV 后，直接运行命令（替换 --files 和 --labels）
2. 如题目结构未变，旧的 `compare_map.json` 可继续沿用
3. 如题目有增减，更新 `compare_map.json` 的 `manual_pairs` 部分

## stdout JSON 格式

```json
{
  "status": "success",
  "output_path": "...",
  "matched_questions": 42,
  "unmatched_a_only": 8,
  "unmatched_b_only": 5,
  "sheets": ["📊 指标总览", "📋 逐题对比", "👥 人群结构", "💬 文本主题"]
}
```
```

（将上述内容写入 `references/16-compare-workflow.md` 文件，注意去掉多余的代码块嵌套）

- [ ] **Step 8.2：更新 `README.md`，在工作流部分新增多期对比触发条件**

在 README.md 的整体工作流 section 内，现有的"阶段 5：生成报告"之后追加：

```markdown
### 阶段 6：多期对比分析（按需）

**触发条件**：用户有两期或以上问卷数据，想对比趋势/变化（如"两个赛季对比"、"这月和上月差异"）。

→ **读取 `references/16-compare-workflow.md` 获取完整执行步骤。**
```

- [ ] **Step 8.3：最终 Commit**

```bash
cd C:\Users\lijinghui03\.agents\skills\survey-research
git add scripts/survey_compare.py references/16-compare-workflow.md README.md
git commit -m "feat: add survey_compare tool + reference doc + skill integration"
```

---

## 自查清单

- [x] 设计文档 Task 1-8 覆盖所有 4 个 Sheet
- [x] `compare_map.json` 格式在 reference doc 和 Task 2.1 中均有定义
- [x] 趋势标识规则（▲▼ 颜色）与设计文档 Section 6 一致
- [x] `_trend_label` 函数在 Task 3.1 定义，Task 3.2、6.1 均复用，无命名不一致
- [x] `_get_option_pcts` 在 Task 4.1 定义，Task 5.1 复用
- [x] `_find_demo_cols` 覆盖性别/年龄/职业/段位/游玩情况关键词
- [x] 端到端测试使用真实存在的数据文件路径
- [x] 文本主题 Sheet 在 `--text_results` 未传入时不生成，无崩溃
