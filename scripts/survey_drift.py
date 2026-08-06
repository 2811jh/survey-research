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
    if a.std(ddof=1) == 0 and b.std(ddof=1) == 0:
        p = 1.0 if mean_a == mean_b else 0.0
        return {"test": "degenerate", "p": p, "mean_a": mean_a, "mean_b": mean_b, "delta": delta}
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
            rates[opt_name(sc)] = float(sel / n) if n else 0.0
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


# ===================== findings 组装 ===================== #

def _multi_choice_label(root, subcols):
    """从多选子列还原完整题干：取首个子列冒号前的部分（含 root）。
    如 'Q6.为什么回归？:游戏版本更新' → 'Q6.为什么回归？'；失败回退 root。"""
    for sc in subcols:
        s = str(sc)
        idxs = [i for i in (s.find(":"), s.find("：")) if i != -1]
        if idxs:
            stem = s[:min(idxs)].strip()
            if stem:
                return stem
    return root


def build_findings(df, classification, granularity, time_col,
                   nps_col, satisfaction_cols, min_n=30):
    if time_col not in df.columns:
        raise KeyError(f"时间列不存在：{time_col}")
    dt = pd.to_datetime(df[time_col], errors="coerce")
    labels, ordered = bucketize(dt, granularity)
    sizes_all = {b: int((labels == b).sum()) for b in ordered}
    low_n_buckets = [b for b, n in sizes_all.items() if n < min_n]

    metrics = []
    for col in (satisfaction_cols or []):
        if col not in df.columns:
            continue
        by_bucket, _ = scale_means(df, col, labels, ordered)
        adj = adjacent_mean_tests(df, col, labels, ordered, min_n)
        metrics.append({
            "name": f"{col} 均分", "type": "satisfaction_mean", "source_col": col,
            "by_bucket": {b: round(by_bucket[b], 2) for b in ordered}, "adjacent": adj,
        })
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
    for col in classification.get("single_choice", []):
        by_bucket, sizes = single_choice_props(df, col, labels, ordered)
        opt_tests = adjacent_prop_tests(by_bucket, sizes, ordered, min_n)
        overall = _overall_chi_square(by_bucket, sizes, ordered)
        drift = any(t["drift"] for t in opt_tests)
        questions.append({
            "question": col, "type": "single_choice",
            "question_label": col,
            "options": sorted({o for b in ordered for o in by_bucket.get(b, {})}),
            "by_bucket": {b: {k: round(v, 4) for k, v in by_bucket.get(b, {}).items()} for b in ordered},
            "sizes": sizes, "overall_test": overall,
            "adjacent_option_tests": opt_tests, "drift": drift,
            "low_n": any(sizes.get(b, 0) < min_n for b in ordered),
        })
    for root, subcols in classification.get("multi_choice", {}).items():
        by_bucket, sizes = multi_choice_rates(df, subcols, root, labels, ordered)
        opt_tests = adjacent_prop_tests(by_bucket, sizes, ordered, min_n)
        drift = any(t["drift"] for t in opt_tests)
        questions.append({
            "question": root, "type": "multi_choice",
            "question_label": _multi_choice_label(root, subcols),
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
            "p": round(float(p), 4), "significant": bool(p < 0.05)}


# ===================== Excel 导出 ===================== #

def _trend_mark(delta, significant, is_pp):
    unit = "pp" if is_pp else "分"
    prec = 1 if is_pp else 2
    if not significant:
        return f"— {delta:+.{prec}f}{unit}（不显著）"
    arrow = "▲" if delta > 0 else "▼"
    return f"{arrow} {delta:+.{prec}f}{unit}"


def export_excel(findings, conclusions, output_path, summary_scope="latest"):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
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
    _format_overview_sheet(ws1, len(buckets))

    # ---- Sheet 2: 逐题异动明细 ----
    ws2 = wb.create_sheet("📈 逐题异动明细")
    ws2.append(["题目", "选项"] + buckets + ["异动周", "AI 结论"])
    block_ranges = []          # (start_row, end_row) 每题一块，供分块斑马纹
    week_marks = {}            # (row, col) -> (kind, direction) 逐周环比标注
    concl_col = 2 + len(buckets) + 2  # AI结论列号
    for q in findings["questions"]:
        opts = q["options"]
        start_row = ws2.max_row + 1
        # 逐选项、逐"到达桶"的相邻期检验：{option: {to_bucket: test}}
        tests_by_opt = {}
        for t in q["adjacent_option_tests"]:
            tests_by_opt.setdefault(t["option"], {})[t["to"]] = t
        for i, opt in enumerate(opts):
            row_idx = ws2.max_row + 1
            drift_weeks = []
            for bi in range(1, len(buckets)):  # 从第 2 桶起才有环比
                t = tests_by_opt.get(opt, {}).get(buckets[bi])
                if not t:
                    continue
                col = 3 + bi
                if t.get("drift"):
                    week_marks[(row_idx, col)] = ("drift", t["direction"])
                    arrow = "▲" if t["delta_pp"] > 0 else "▼"
                    drift_weeks.append(f"{buckets[bi]}{arrow}{t['delta_pp']:+.1f}pp")
                elif t.get("significant"):
                    week_marks[(row_idx, col)] = ("sig", t["direction"])
            ws2.append([
                q.get("question_label", q["question"]) if i == 0 else "", opt,
                *[q["by_bucket"].get(b, {}).get(opt, 0.0) for b in buckets],
                "；".join(drift_weeks), "",
            ])
        end_row = ws2.max_row
        block_ranges.append((start_row, end_row))
        concl = conclusions.get(q["question"], "")
        if end_row >= start_row:
            ws2.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws2.merge_cells(start_row=start_row, start_column=concl_col,
                            end_row=end_row, end_column=concl_col)
            ws2.cell(row=start_row, column=concl_col, value=concl)
    _format_detail_sheet(ws2, len(buckets), block_ranges, week_marks)

    # ---- Sheet 3: 异动汇总 ----
    ws3 = wb.create_sheet("⚠️ 异动汇总")
    ws3.append(["题目/指标", "时段", "变化项", "方向", "幅度", "显著性", "AI结论"])

    def _in_scope(t):
        return True if summary_scope == "all" else (t["to"] == buckets[-1])

    any_drift = False
    for m in findings["metrics"]:
        for t in m["adjacent"]:
            if t.get("drift") and _in_scope(t):
                any_drift = True
                d = t.get("delta_pp", t.get("delta", 0.0))
                arrow = "▲" if d > 0 else "▼"
                ws3.append([m["name"], t["to"], "整体", arrow, f"{d:+.2f}",
                            f"p={t['p']}", conclusions.get(m["source_col"], "")])
    for q in findings["questions"]:
        for t in q["adjacent_option_tests"]:
            if t.get("drift") and _in_scope(t):
                any_drift = True
                arrow = "▲" if t["delta_pp"] > 0 else "▼"
                ws3.append([q.get("question_label", q["question"]), t["to"], t["option"], arrow,
                            f"{t['delta_pp']:+.1f}pp", f"p={t['p']}",
                            conclusions.get(q["question"], "")])
    if not any_drift:
        msg = "本期各指标/题目均无显著异动" if summary_scope == "latest" \
            else "全时间线各指标/题目均无显著异动"
        ws3.append([msg, "", "", "", "", "", ""])
    _format_summary_sheet(ws3, no_drift=(not any_drift))

    # ---- Sheet 4: 方法与样本 ----
    ws4 = wb.create_sheet("ℹ️ 方法与样本")
    ws4.append(["项", "说明"])
    ws4.append(["分桶粒度", {"week": "按周", "month": "按月", "day": "按天"}.get(findings["granularity"])])
    ws4.append(["时间列", findings["time_col"]])
    ws4.append(["各桶样本量", "; ".join(f"{b}={sizes.get(b,0)}" for b in buckets)])
    ws4.append(["样本不足桶(n<30)", "; ".join(findings["low_n_buckets"]) or "无"])
    ws4.append(["判异动门槛", "p<0.05 且（占比Δ≥5pp 或 均分Δ≥0.1）"])
    ws4.append(["检验方法", "均分:t检验/Mann-Whitney; 占比:两比例z; 单选整体:卡方"])
    ws4.append(["明细表颜色", "逐题明细中，某周单元格相对前一周显著变化会着色：琥珀底+加粗=大幅异动(双门槛)，红/绿字=一般显著(升绿/降红)，灰字=无显著环比变化"])
    ws4.append(["免责", "样本不足桶仅供参考，不判异动"])
    _format_method_sheet(ws4)

    wb.save(output_path)
    return {"status": "success", "output_path": output_path, "sheets": wb.sheetnames}


def _format_detail_sheet(ws, n_buckets, block_ranges=None, week_marks=None):
    """逐题异动明细：Slate + Indigo 设计系统（对齐文本分析 Excel 风格）。
    深色表头 + 按题分块斑马纹 + 占比 DataBar + 逐周环比热力标注 + 异动周列 + 结论靛蓝卡片。
    week_marks: {(row, col): (kind, direction)}，kind ∈ {'drift','sig'}，标注某周相对前一周的显著变化。"""
    import _styles as st
    from _styles import TextReportTheme as TR, Theme
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import DataBarRule

    week_marks = week_marks or {}
    border = st.thin_border()
    max_col = ws.max_column
    max_row = ws.max_row
    b_first, b_last = 3, 2 + n_buckets
    drift_weeks_col, concl_col = b_last + 1, b_last + 2

    DRIFT_BG = "FEF3C7"   # amber-100 异动周高亮
    UP_FONT = "1E7D32"    # green-800 升
    DOWN_FONT = "C0392B"  # red-700 降
    center = st.ALIGN_CENTER
    left = st.ALIGN_LEFT
    top_left = st.ALIGN_TOP_LEFT

    # ---- 表头 ----
    ws.row_dimensions[1].height = 40
    for c in range(1, max_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = st.make_fill(TR.TITLE_BG)
        cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.WHITE)
        cell.alignment = center
        cell.border = border

    # ---- 分块斑马纹底色（每题一色，块间交替）----
    if not block_ranges:
        block_ranges = [(r, r) for r in range(2, max_row + 1)]
    row_base = {}
    for bi, (s, e) in enumerate(block_ranges):
        base = TR.WHITE if bi % 2 == 0 else TR.ZEBRA_ALT
        for r in range(s, e + 1):
            row_base[r] = base

    # ---- 数据行 ----
    for r in range(2, max_row + 1):
        base = row_base.get(r, TR.WHITE)
        ws.row_dimensions[r].height = 22
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if c == 1:  # 题目（合并列）
                cell.fill = st.make_fill(TR.INDIGO_ACCENT_BG)
                cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.INDIGO_DEEP)
                cell.alignment = top_left
            elif c == 2:  # 选项
                cell.fill = st.make_fill(base)
                cell.font = Font(name=Theme.FONT_NAME, size=10, color=TR.TEXT_MAIN)
                cell.alignment = left
            elif b_first <= c <= b_last:  # 各周占比 + 逐周环比热力标注
                cell.number_format = "0.0%"
                cell.alignment = center
                mark = week_marks.get((r, c))
                if mark:
                    kind, direction = mark
                    color = UP_FONT if direction == "up" else DOWN_FONT
                    if kind == "drift":  # 大幅异动：琥珀底 + 加粗彩字
                        cell.fill = st.make_fill(DRIFT_BG)
                        cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=color)
                    else:                # 一般显著：彩字，不加底
                        cell.fill = st.make_fill(base)
                        cell.font = Font(name=Theme.FONT_NAME, size=10, color=color)
                else:                    # 无显著环比变化
                    cell.fill = st.make_fill(base)
                    cell.font = Font(name=Theme.FONT_NAME, size=10, color=TR.TEXT_SUB)
            elif c == drift_weeks_col:  # 异动周
                cell.fill = st.make_fill(base)
                has = bool(cell.value)
                cell.font = Font(name=Theme.FONT_NAME, size=9, bold=has,
                                 color=("B45309" if has else TR.TEXT_MUTE))
                cell.alignment = st.ALIGN_TOP_LEFT
            elif c == concl_col:  # AI 结论（合并卡片）
                cell.fill = st.make_fill(TR.INDIGO_BG)
                cell.font = Font(name=Theme.FONT_NAME, size=10, color=TR.INDIGO_DEEP)
                cell.alignment = top_left

    # ---- 占比列 DataBar（固定 0~1 刻度，跨题可比）----
    if max_row >= 2:
        for c in range(b_first, b_last + 1):
            col = get_column_letter(c)
            rng = f"{col}2:{col}{max_row}"
            rule = DataBarRule(start_type="num", start_value=0,
                               end_type="num", end_value=1,
                               color=TR.INDIGO_CHIP, showValue=True,
                               minLength=0, maxLength=100)
            ws.conditional_formatting.add(rng, rule)

    # ---- 列宽 / 冻结 ----
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 26
    for c in range(b_first, b_last + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11
    ws.column_dimensions[get_column_letter(drift_weeks_col)].width = 30
    ws.column_dimensions[get_column_letter(concl_col)].width = 46
    ws.freeze_panes = "C2"
    ws.sheet_view.showGridLines = False


# ===================== Slate + Indigo 皮肤（其余 3 Sheet） ===================== #

_UP_FONT = "1E7D32"    # green-800 升
_DOWN_FONT = "C0392B"  # red-700 降
_DRIFT_BG = "FEF3C7"   # amber-100


def _slate_header(ws, height=38):
    """统一深色表头（slate-800 + 白色粗体）。"""
    import _styles as st
    from _styles import TextReportTheme as TR, Theme
    from openpyxl.styles import Font
    ws.row_dimensions[1].height = height
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = st.make_fill(TR.TITLE_BG)
        cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.WHITE)
        cell.alignment = st.ALIGN_CENTER
        cell.border = st.thin_border()


def _format_overview_sheet(ws, n_buckets):
    """指标总览：深色表头 + 靛蓝指标列 + 斑马纹 + 趋势/显著性颜色编码。"""
    import _styles as st
    from _styles import TextReportTheme as TR, Theme
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    border = st.thin_border()
    _slate_header(ws)
    trend_col = 1 + n_buckets + 1
    sig_col = 1 + n_buckets + 2
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 26
        is_sample = str(ws.cell(r, 1).value or "").strip() == "样本量"
        zebra = TR.WHITE if r % 2 == 0 else TR.ZEBRA_ALT
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if c == 1:  # 指标
                cell.fill = st.make_fill(TR.INDIGO_ACCENT_BG)
                cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.INDIGO_DEEP)
                cell.alignment = st.ALIGN_LEFT
            elif c <= 1 + n_buckets:  # 各期数值
                cell.fill = st.make_fill(zebra)
                cell.font = Font(name=Theme.FONT_NAME, size=11, bold=True, color=TR.INDIGO_MAIN)
                cell.alignment = st.ALIGN_CENTER
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0" if is_sample else "0.00"
            elif c == trend_col:  # 最新vs上期
                cell.fill = st.make_fill(zebra)
                cell.alignment = st.ALIGN_CENTER
                v = str(cell.value or "")
                color = _UP_FONT if v.startswith("▲") else (_DOWN_FONT if v.startswith("▼") else TR.TEXT_MUTE)
                cell.font = Font(name=Theme.FONT_NAME, size=10, bold=v[:1] in ("▲", "▼"), color=color)
            elif c == sig_col:  # 是否显著
                cell.fill = st.make_fill(zebra)
                cell.alignment = st.ALIGN_CENTER
                sig = str(cell.value or "").startswith("✅")
                cell.font = Font(name=Theme.FONT_NAME, size=9, bold=sig,
                                 color=(TR.INDIGO_MAIN if sig else TR.TEXT_MUTE))
    ws.column_dimensions["A"].width = 40
    for c in range(2, 2 + n_buckets):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.column_dimensions[get_column_letter(trend_col)].width = 18
    ws.column_dimensions[get_column_letter(sig_col)].width = 14
    ws.freeze_panes = "B2"
    ws.sheet_view.showGridLines = False


def _format_summary_sheet(ws, no_drift=False):
    """异动汇总：深色表头 + 方向/幅度颜色编码 + AI 结论靛蓝卡片。
    列：题目/指标 | 时段 | 变化项 | 方向 | 幅度 | 显著性 | AI结论。"""
    import _styles as st
    from _styles import TextReportTheme as TR, Theme
    from openpyxl.styles import Font
    border = st.thin_border()
    ncol = ws.max_column  # 7
    _slate_header(ws)
    if no_drift:  # 单行提示：整行合并、居中弱化
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
        cell = ws.cell(row=2, column=1)
        cell.fill = st.make_fill(TR.NOTE_BG)
        cell.font = Font(name=Theme.FONT_NAME, size=11, color=TR.TEXT_SUB)
        cell.alignment = st.ALIGN_CENTER
        ws.row_dimensions[2].height = 40
        for c in range(1, ncol + 1):
            ws.cell(row=2, column=c).border = border
    else:
        for r in range(2, ws.max_row + 1):
            zebra = TR.WHITE if r % 2 == 0 else TR.ZEBRA_ALT
            arrow = str(ws.cell(r, 4).value or "")
            dcolor = _UP_FONT if arrow.startswith("▲") else (_DOWN_FONT if arrow.startswith("▼") else TR.TEXT_MAIN)
            ws.row_dimensions[r].height = 30
            for c in range(1, ncol + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                if c == 1:  # 题目/指标
                    cell.fill = st.make_fill(TR.INDIGO_ACCENT_BG)
                    cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.INDIGO_DEEP)
                    cell.alignment = st.ALIGN_LEFT
                elif c == 2:  # 时段
                    cell.fill = st.make_fill(zebra)
                    cell.font = Font(name=Theme.FONT_NAME, size=9, color=TR.TEXT_SUB)
                    cell.alignment = st.ALIGN_CENTER
                elif c == 3:  # 变化项
                    cell.fill = st.make_fill(zebra)
                    cell.font = Font(name=Theme.FONT_NAME, size=10, color=TR.TEXT_MAIN)
                    cell.alignment = st.ALIGN_LEFT
                elif c in (4, 5):  # 方向 / 幅度
                    cell.fill = st.make_fill(zebra)
                    cell.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=dcolor)
                    cell.alignment = st.ALIGN_CENTER
                elif c == 6:  # 显著性
                    cell.fill = st.make_fill(zebra)
                    cell.font = Font(name=Theme.FONT_NAME, size=9, color=TR.TEXT_SUB)
                    cell.alignment = st.ALIGN_CENTER
                else:  # AI 结论
                    cell.fill = st.make_fill(TR.INDIGO_BG)
                    cell.font = Font(name=Theme.FONT_NAME, size=10, color=TR.INDIGO_DEEP)
                    cell.alignment = st.ALIGN_TOP_LEFT
    for col, w in zip("ABCDEFG", (32, 20, 24, 8, 12, 12, 48)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _format_method_sheet(ws):
    """方法与样本：深色表头 + 靛蓝项名列 + 说明列浅底斑马纹。"""
    import _styles as st
    from _styles import TextReportTheme as TR, Theme
    from openpyxl.styles import Font
    border = st.thin_border()
    _slate_header(ws)
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 28
        zebra = TR.WHITE if r % 2 == 0 else TR.ZEBRA_ALT
        c1 = ws.cell(row=r, column=1)
        c1.fill = st.make_fill(TR.INDIGO_ACCENT_BG)
        c1.font = Font(name=Theme.FONT_NAME, size=10, bold=True, color=TR.INDIGO_DEEP)
        c1.alignment = st.ALIGN_LEFT
        c1.border = border
        c2 = ws.cell(row=r, column=2)
        c2.fill = st.make_fill(zebra)
        c2.font = Font(name=Theme.FONT_NAME, size=10, color=TR.TEXT_MAIN)
        c2.alignment = st.ALIGN_LEFT
        c2.border = border
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 72
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


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
    return export_excel(findings, conclusions, out, summary_scope=args.summary_scope)


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
    pe.add_argument("--summary-scope", dest="summary_scope",
                    choices=["latest", "all"], default="latest",
                    help="异动汇总范围：latest=仅最新相邻期（默认）；all=全时间线任意相邻期")

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
