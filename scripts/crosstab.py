#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 交叉分析
========================

完整的交叉分析流水线：
合并选项 → 交叉计算 → 得分计算 → 差异摘要 → 导出 Excel

输出专业格式化的 Excel + stdout JSON（交叉摘要 + 差异 + 得分）。

用法:
    python crosstab.py \
        --file_path "C:/xxx/data.xlsx" \
        --row_questions '["all"]' \
        --col_questions '["Q17.性别"]' \
        [--merge_rules '{"Q1.满意度": {"不满意": [1,2,3], "满意": [4,5]}}'] \
        [--calc_scores auto] \
        [--output_path "C:/xxx/data_交叉分析.xlsx"] \
        [--report_json '{"per_question":[...]}']
"""

import argparse
import json
import sys
import os
import re
import warnings
import pandas as pd
import numpy as np
from collections import defaultdict
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_and_classify import classify_columns
from _styles import (
    format_data_sheet, format_score_sheet, write_structured_report,
    Theme, header_fill, header_font, body_font,
    thin_border, make_fill, ALIGN_CENTER, ALIGN_LEFT, ALIGN_RIGHT,
)


# ========================================================================= #
#                           辅助函数
# ========================================================================= #

def _detect_csv_encoding(filepath, sample_size=8192):
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


def _extract_subcol_number(subcol: str, prefix: str) -> int:
    """从多选子列名中提取选项序号"""
    suffix = subcol.split(prefix)[1].strip()
    match = re.search(r'^(\d+)', suffix)
    return int(match.group(1)) if match else 0


def _extract_score_from_option(option) -> Optional[float]:
    """从选项文本中提取数值分数"""
    if option is None:
        return None
    s = str(option).strip()
    if not s:
        return None
    match = re.match(r'^-?\d+(?:\.\d+)?', s)
    if not match:
        match = re.search(r'-?\d+(?:\.\d+)?', s)
    return float(match.group(0)) if match else None


# ========================================================================= #
#                        合并选项 (Merge / Recode)
# ========================================================================= #

def merge_options(
    df: pd.DataFrame,
    column: str,
    merge_rules: dict,
    new_column_name: Optional[str] = None,
) -> str:
    """
    合并指定列的选项值，在 df 上原地添加新列。

    Args:
        df: 数据 DataFrame
        column: 原始列名
        merge_rules: {"不满意": [1,2,3], "满意": [4,5]}
        new_column_name: 新列名（默认自动生成）

    Returns:
        新列名
    """
    if column not in df.columns:
        raise ValueError(f"列 '{column}' 不存在")

    mapping = {}
    for label, values in merge_rules.items():
        for v in values:
            mapping[v] = label

    if new_column_name is None:
        short_name = re.sub(r'^Q\d+\.', '', column).strip()
        if len(short_name) > 20:
            short_name = short_name[:20]
        new_column_name = f"recode_{short_name}"

    df[new_column_name] = df[column].map(mapping)
    return new_column_name


# ========================================================================= #
#                        交叉分析核心
# ========================================================================= #

def run_crosstab(
    df: pd.DataFrame,
    classification: dict,
    row_questions: list,
    col_questions: list,
) -> dict:
    """
    执行交叉分析。

    Args:
        df: 数据 DataFrame
        classification: 列分类信息
        row_questions: 行变量列表（支持 "all"、具体列名、多选题根如 "Q8."）
        col_questions: 列变量列表（分组维度）

    Returns:
        {
            "freq_df": DataFrame,       # 频数表
            "percent_df": DataFrame,    # 列百分比
            "col_totals": dict,         # 列合计
            "col_labels": list,         # 列标签
            "valid_rows_map": dict,     # 行变量类型映射
        }
    """
    multi_dict = classification["multi_choice"]

    # --- 处理 "all" ---
    if row_questions == ["all"] or row_questions == "all":
        row_questions = list(classification["valid_for_crosstab"])
        col_sources = set()
        for cq in col_questions:
            col_sources.add(cq)
        row_questions = [q for q in row_questions if q not in col_sources]

    # --- 验证并分类问题 ---
    def validate_and_classify(questions):
        valid = []
        invalid = []
        for q in questions:
            q_clean = str(q).strip()
            if q_clean in multi_dict:
                valid.append(("multi", q_clean))
            elif re.match(r'^Q\d+\.$', q_clean) and q_clean in multi_dict:
                valid.append(("multi", q_clean))
            elif q_clean in df.columns:
                valid.append(("single", q_clean))
            else:
                invalid.append(q_clean)
        return valid, invalid

    valid_rows, invalid_rows = validate_and_classify(row_questions)
    valid_cols, invalid_cols = validate_and_classify(col_questions)

    if invalid_rows:
        warnings.warn(f"无效行问题将被跳过：{invalid_rows}")
    if invalid_cols:
        warnings.warn(f"无效列问题将被跳过：{invalid_cols}")

    # --- 列条件生成 ---
    col_conditions = []
    col_totals = {}
    seen_cols = defaultdict(int)

    for q_type, q in valid_cols:
        q_clean = str(q).strip()
        seen_cols[q_clean] += 1
        instance_id = seen_cols[q_clean]

        if q_type == "multi":
            root = q_clean
            subcols = multi_dict[root]
            example_subcol = subcols[0]
            rest_part = example_subcol.split(root)[1].strip()
            if ':' in rest_part:
                question_text = rest_part.split(':', 1)[0].strip()
            elif '：' in rest_part:
                question_text = rest_part.split('：', 1)[0].strip()
            else:
                question_text = rest_part
            full_question = f"{root}{question_text}"
            if instance_id > 1:
                full_question += f" #{instance_id}"

            for subcol in subcols:
                rest_subcol = subcol.split(root)[1].strip()
                if ':' in rest_subcol:
                    option_text = rest_subcol.split(':', 1)[1].strip()
                elif '：' in rest_subcol:
                    option_text = rest_subcol.split('：', 1)[1].strip()
                else:
                    option_text = rest_subcol
                label = f"{full_question}\n{option_text}"
                cond = df[subcol] == 1
                col_conditions.append((label, cond))
                col_totals[label] = int(cond.sum())

            total_label = f"{full_question}\n总计"
            total_cond = (df[subcols] == 1).any(axis=1)
            col_conditions.append((total_label, total_cond))
            col_totals[total_label] = int(total_cond.sum())

        else:
            values = df[q_clean].dropna().unique()
            try:
                sorted_values = sorted(
                    values,
                    key=lambda x: int(re.match(r'^(\d+)', str(x)).group(1))
                )
            except Exception:
                sorted_values = sorted(values, key=str)

            unique_question = q_clean
            if instance_id > 1:
                unique_question += f" #{instance_id}"

            for value in sorted_values:
                label = f"{unique_question}\n{value}"
                cond = df[q_clean] == value
                col_conditions.append((label, cond))
                col_totals[label] = int(cond.sum())

            total_label = f"{unique_question}\n总计"
            total_cond = df[q_clean].notna()
            col_conditions.append((total_label, total_cond))
            col_totals[total_label] = int(total_cond.sum())

    # --- 行条件生成 ---
    row_conditions = []
    for q_type, q in valid_rows:
        if q_type == "multi":
            root = q
            subcols = multi_dict[root]
            first_rest = subcols[0].split(root)[1].strip()
            if ':' in first_rest:
                q_text = first_rest.split(':', 1)[0].strip()
            elif '：' in first_rest:
                q_text = first_rest.split('：', 1)[0].strip()
            else:
                q_text = first_rest
            full_question = f"{root}{q_text}"

            for subcol in subcols:
                rest = subcol.split(root)[1].strip()
                if ':' in rest:
                    option_text = rest.split(':', 1)[1].strip()
                elif '：' in rest:
                    option_text = rest.split('：', 1)[1].strip()
                else:
                    option_text = rest
                cond = df[subcol] == 1
                row_conditions.append(((full_question, option_text), cond))
            total_cond = (df[subcols] == 1).any(axis=1)
            row_conditions.append(((full_question, "总计"), total_cond))
        else:
            values = df[q].dropna().unique()
            try:
                sorted_values = sorted(
                    values,
                    key=lambda x: int(re.match(r'^(\d+)', str(x)).group(1))
                )
            except Exception:
                sorted_values = sorted(values, key=str)
            for value in sorted_values:
                cond = df[q] == value
                row_conditions.append(((q, str(value)), cond))
            total_cond = df[q].notna()
            row_conditions.append(((q, "总计"), total_cond))

    # --- 交叉统计计算 ---
    freq_results = []
    for (r_question, r_option), r_cond in row_conditions:
        row_data = {}
        for c_label, c_cond in col_conditions:
            count = int((r_cond & c_cond).sum())
            row_data[c_label] = count
        freq_results.append(row_data)

    index = pd.MultiIndex.from_tuples(
        [(rl[0], rl[1]) for rl, _ in row_conditions],
        names=["问题", "选项"]
    )
    col_labels = [cl for cl, _ in col_conditions]

    freq_df = pd.DataFrame(freq_results, index=index, columns=col_labels)

    # --- 列百分比 ---
    percent_df = freq_df.astype(float).copy()
    for question in percent_df.index.get_level_values(0).unique():
        q_mask = percent_df.index.get_level_values(0) == question
        total_idx = (question, "总计")
        if total_idx in freq_df.index:
            denom = freq_df.loc[total_idx].replace(0, np.nan)
        else:
            denom = pd.Series(col_totals).reindex(percent_df.columns).replace(0, np.nan)
        percent_df.loc[q_mask] = freq_df.loc[q_mask].div(denom, axis=1)
    percent_df = percent_df.fillna(0)

    return {
        "freq_df": freq_df,
        "percent_df": percent_df,
        "col_totals": col_totals,
        "col_labels": col_labels,
        "valid_rows_map": {q: q_type for q_type, q in valid_rows},
        "invalid_rows": invalid_rows,
        "invalid_cols": invalid_cols,
    }


# ========================================================================= #
#                      满意度 / NPS 得分计算
# ========================================================================= #

def _detect_score_type(question_name: str, df: pd.DataFrame) -> str:
    """自动识别题目是满意度还是 NPS"""
    q_lower = question_name.lower()
    if "nps" in q_lower or "推荐" in question_name:
        return "nps"
    if "满意度" in question_name or "满意" in question_name:
        return "satisfaction"
    if question_name in df.columns:
        values = df[question_name].dropna().unique()
        numeric_vals = []
        for v in values:
            score = _extract_score_from_option(v)
            if score is not None:
                numeric_vals.append(score)
        if numeric_vals:
            min_val, max_val = min(numeric_vals), max(numeric_vals)
            if min_val >= 0 and max_val >= 9:
                return "nps"
    return "satisfaction"


def _is_scoreable_question(question_name: str, df: pd.DataFrame) -> Optional[str]:
    """判断题目是否适合计算得分"""
    q_str = str(question_name)

    satisfaction_keywords = ["满意度", "满意", "评价如何", "评价是", "体验感受"]
    nps_keywords = ["NPS", "nps", "推荐"]

    has_satisfaction = any(kw in q_str for kw in satisfaction_keywords)
    has_nps = any(kw in q_str or kw.lower() in q_str.lower() for kw in nps_keywords)

    if not has_satisfaction and not has_nps:
        return None

    if question_name not in df.columns:
        return None

    values = df[question_name].dropna().unique()
    numeric_vals = []
    for v in values:
        score = _extract_score_from_option(v)
        if score is not None:
            numeric_vals.append(score)

    if len(numeric_vals) < 2:
        return None

    min_val, max_val = min(numeric_vals), max(numeric_vals)

    if has_nps:
        if min_val >= 0 and max_val >= 9 and max_val <= 10:
            return "nps"
        if has_satisfaction and min_val >= 1 and max_val <= 7:
            return "satisfaction"
        return None

    if has_satisfaction:
        if min_val >= 1 and max_val <= 10 and (max_val - min_val) >= 2:
            return "satisfaction"
        return None

    return None


def auto_detect_score_questions(df: pd.DataFrame, ct_result: dict) -> list:
    """自动检测所有需要计算得分的题目"""
    row_type_map = ct_result["valid_rows_map"]
    freq_df = ct_result["freq_df"]

    scoreable = []
    for q_name in freq_df.index.get_level_values(0).unique():
        if row_type_map.get(q_name) != "single":
            continue
        score_type = _is_scoreable_question(q_name, df)
        if score_type is not None:
            scoreable.append(q_name)
    return scoreable


def calc_scores(df: pd.DataFrame, ct_result: dict, score_questions: list) -> Optional[pd.DataFrame]:
    """
    计算满意度得分或 NPS。

    Returns:
        score_df（得分 DataFrame）或 None
    """
    freq_df = ct_result["freq_df"]
    row_type_map = ct_result["valid_rows_map"]

    score_results = []
    score_index = []
    score_type_info = {}

    for q in score_questions:
        q = str(q).strip()

        if q not in freq_df.index.get_level_values(0).unique():
            warnings.warn(f"题目 '{q}' 不在行变量中，已跳过")
            continue
        if row_type_map.get(q) != "single":
            warnings.warn(f"得分计算仅支持单选/量表题，已跳过：{q}")
            continue

        score_type = _detect_score_type(q, df)
        score_type_info[q] = score_type

        q_slice = freq_df.xs(q, level=0)

        if score_type == "satisfaction":
            value_map = {}
            for opt in q_slice.index:
                opt_str = str(opt).strip()
                if opt_str in ("总计", "合计", "Total"):
                    continue
                score_val = _extract_score_from_option(opt_str)
                if score_val is not None:
                    value_map[opt] = score_val

            if not value_map:
                continue

            q_counts = q_slice.loc[list(value_map.keys())]
            weights = pd.Series(value_map)
            numerator = (q_counts.T * weights).T.sum(axis=0)
            denominator = q_counts.sum(axis=0).replace(0, np.nan)
            score = numerator / denominator

            score_results.append(score)
            score_index.append((q, "满意度得分(加权均值)"))

        else:
            value_map = {}
            for opt in q_slice.index:
                opt_str = str(opt).strip()
                if opt_str in ("总计", "合计", "Total"):
                    continue
                score_val = _extract_score_from_option(opt_str)
                if score_val is not None:
                    value_map[opt] = score_val

            if not value_map:
                continue

            promoter_opts = [opt for opt, s in value_map.items() if s >= 9]
            detractor_opts = [opt for opt, s in value_map.items() if s <= 6]

            q_counts = q_slice.loc[list(value_map.keys())]
            total_count = q_counts.sum(axis=0).replace(0, np.nan)

            promoter_count = q_counts.loc[promoter_opts].sum(axis=0) if promoter_opts else 0
            detractor_count = q_counts.loc[detractor_opts].sum(axis=0) if detractor_opts else 0

            nps_score = (promoter_count - detractor_count) / total_count

            score_results.append(nps_score)
            score_index.append((q, "NPS得分(%)"))

    if not score_results:
        return None

    score_df = pd.DataFrame(
        score_results,
        index=pd.MultiIndex.from_tuples(score_index, names=["问题", "指标"]),
    )
    score_df = score_df.reindex(columns=freq_df.columns)
    return score_df


# ========================================================================= #
#                      差异摘要
# ========================================================================= #

def get_crosstab_summary(ct_result: dict) -> dict:
    """提取关键差异摘要"""
    percent_df = ct_result["percent_df"]
    col_labels = ct_result["col_labels"]

    total_cols = [c for c in col_labels if c.endswith("\n总计")]
    non_total_cols = [c for c in col_labels if not c.endswith("\n总计")]

    summary = {}

    for question in percent_df.index.get_level_values(0).unique():
        q_data = percent_df.xs(question, level=0)
        option_rows = [opt for opt in q_data.index if opt != "总计"]
        if not option_rows:
            continue

        question_summary = {
            "options": {},
            "max_diff_option": None,
            "max_diff_value": 0,
        }

        for opt in option_rows:
            opt_percents = {}
            for col in non_total_cols:
                pct = float(q_data.loc[opt, col]) if opt in q_data.index else 0
                opt_percents[col] = round(pct, 4)

            pct_values = list(opt_percents.values())
            diff = max(pct_values) - min(pct_values) if pct_values else 0

            question_summary["options"][str(opt)] = {
                "percentages": opt_percents,
                "max_min_diff": round(diff, 4),
            }

            if diff > question_summary["max_diff_value"]:
                question_summary["max_diff_value"] = round(diff, 4)
                question_summary["max_diff_option"] = str(opt)

        summary[question] = question_summary

    return summary


# ========================================================================= #
#                      Excel 导出
# ========================================================================= #

def export_crosstab_excel(
    ct_result: dict,
    output_path: str,
    score_df: Optional[pd.DataFrame] = None,
    report_text: str = "",
) -> str:
    """导出交叉分析 Excel 报告"""
    freq_df = ct_result["freq_df"]
    percent_df = ct_result["percent_df"]
    col_labels = ct_result["col_labels"]

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except PermissionError:
            raise PermissionError(f"请关闭正在使用的文件：{output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: 交叉分析（频数）
        freq_df.to_excel(writer, sheet_name='交叉分析', merge_cells=True)
        format_data_sheet(writer.sheets['交叉分析'], is_percent=False)
        writer.sheets['交叉分析'].sheet_properties.tabColor = "B4C6E7"

        # Sheet 2: 列百分比
        percent_df.to_excel(writer, sheet_name='列百分比', merge_cells=True)
        format_data_sheet(writer.sheets['列百分比'], is_percent=True)
        writer.sheets['列百分比'].sheet_properties.tabColor = "F8CBAD"

        # Sheet 3: 得分分析（如有）
        if score_df is not None and not score_df.empty:
            score_df.to_excel(writer, sheet_name='得分分析', merge_cells=True)
            format_score_sheet(writer.sheets['得分分析'])
            writer.sheets['得分分析'].sheet_properties.tabColor = "C6EFCE"

        # Sheet 4: 分析报告（如有）
        if report_text:
            _write_report_sheet(writer, report_text, percent_df, col_labels)

    return output_path


def _write_report_sheet(writer, report_text: str, percent_df=None, col_labels=None):
    """写入 AI 分析报告 sheet"""
    try:
        report_data = json.loads(report_text)

        # v3 结构化 JSON
        if isinstance(report_data, dict) and "per_question" in report_data:
            ws = writer.book.create_sheet('分析报告')
            write_structured_report(ws, report_data, percent_df, col_labels or [])
            ws.sheet_properties.tabColor = "D9C4EC"
            return

        # v2 列表 JSON
        if isinstance(report_data, list):
            report_df = pd.DataFrame(report_data)
            col_mapping = {
                "question": "题目",
                "finding": "关键发现",
                "detail": "详细说明",
            }
            report_df = report_df.rename(columns={
                k: v for k, v in col_mapping.items() if k in report_df.columns
            })
            report_df.to_excel(writer, sheet_name='分析报告', index=False)
            ws = writer.sheets['分析报告']
            ws.sheet_properties.tabColor = "D9C4EC"
            return
    except (json.JSONDecodeError, TypeError):
        pass

    # 纯文本模式
    ws = writer.book.create_sheet('分析报告')
    from _styles import Theme
    from openpyxl.styles import Font, Alignment
    ws.cell(row=1, column=1, value="交叉分析报告")
    ws.cell(row=1, column=1).font = Font(name=Theme.FONT_NAME, size=16, bold=True, color=Theme.HEADER_BG)

    lines = report_text.split("\n")
    for i, line in enumerate(lines, start=3):
        cell = ws.cell(row=i, column=1, value=line)
        cell.font = Font(name=Theme.FONT_NAME, size=11)
        cell.alignment = Alignment(wrap_text=True, vertical='top')

    ws.column_dimensions['A'].width = 120
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "D9C4EC"


# ========================================================================= #
#                      JSON 输出生成
# ========================================================================= #

def _generate_output_json(
    ct_result: dict,
    diff_summary: dict,
    score_df: Optional[pd.DataFrame],
    output_path: str,
) -> dict:
    """生成 stdout JSON 输出"""
    freq_df = ct_result["freq_df"]
    percent_df = ct_result["percent_df"]

    # 百分比表摘要
    percent_summary = {}
    for (q, opt) in percent_df.index:
        if q not in percent_summary:
            percent_summary[q] = {}
        percent_summary[q][opt] = {
            col: round(float(percent_df.loc[(q, opt), col]), 4)
            for col in percent_df.columns
        }

    # 得分摘要
    score_summary = None
    if score_df is not None and not score_df.empty:
        score_summary = {}
        non_total_cols = [c for c in ct_result["col_labels"] if not c.endswith("\n总计")]
        for (q, indicator) in score_df.index:
            scores_by_col = {}
            for col in non_total_cols:
                if col in score_df.columns:
                    scores_by_col[col] = round(float(score_df.loc[(q, indicator), col]), 4)
            score_summary[f"{q} - {indicator}"] = scores_by_col

    return {
        "status": "success",
        "output_path": output_path,
        "row_questions_count": len(ct_result["valid_rows_map"]),
        "col_conditions_count": len(ct_result["col_labels"]),
        "invalid_rows": ct_result.get("invalid_rows", []),
        "invalid_cols": ct_result.get("invalid_cols", []),
        "percent_table": percent_summary,
        "diff_summary": diff_summary,
        "score_summary": score_summary,
    }


# ========================================================================= #
#                        主函数
# ========================================================================= #

def run_crosstab_pipeline(
    file_path: str,
    row_questions: list,
    col_questions: list,
    sheet_name=0,
    merge_rules: dict = None,
    calc_scores_mode: str = None,
    output_path: str = None,
    report_json: str = "",
) -> dict:
    """
    完整的交叉分析流水线。

    Args:
        file_path: 数据文件路径
        row_questions: 行变量列表
        col_questions: 列变量列表
        sheet_name: 工作表名或编号
        merge_rules: {"列名": {"新标签": [原始值]}} 合并规则
        calc_scores_mode: "auto" 自动检测 / None 不计算
        output_path: 输出路径
        report_json: AI 报告 JSON 文本

    Returns:
        JSON 输出
    """
    # 加载数据
    ext = file_path.rsplit('.', 1)[-1].lower()
    if ext == 'csv':
        df = pd.read_csv(file_path, encoding=_detect_csv_encoding(file_path))
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]

    classification = classify_columns(df)

    # 合并选项
    if merge_rules:
        for col_name, rules in merge_rules.items():
            new_col = merge_options(df, col_name, rules)
            # 将合并后的列替换到 col_questions 中
            if col_name in col_questions:
                idx = col_questions.index(col_name)
                col_questions[idx] = new_col
            # 更新分类信息
            classification["single_choice"].append(new_col)
            classification["valid_for_crosstab"].append(new_col)

    # 交叉分析
    ct_result = run_crosstab(df, classification, row_questions, col_questions)

    # 得分计算
    score_df = None
    if calc_scores_mode == "auto":
        score_questions = auto_detect_score_questions(df, ct_result)
        if score_questions:
            score_df = calc_scores(df, ct_result, score_questions)
    elif calc_scores_mode and calc_scores_mode != "none":
        try:
            score_questions = json.loads(calc_scores_mode)
            score_df = calc_scores(df, ct_result, score_questions)
        except json.JSONDecodeError:
            pass

    # 差异摘要
    diff_summary = get_crosstab_summary(ct_result)

    # 输出路径
    if output_path is None:
        base = os.path.splitext(file_path)[0]
        output_path = f"{base}_交叉分析.xlsx"

    # 导出 Excel
    export_crosstab_excel(ct_result, output_path, score_df, report_json)

    # 生成 JSON 输出
    return _generate_output_json(ct_result, diff_summary, score_df, output_path)


# ========================================================================= #
#                        CLI 入口
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(description="问卷交叉分析")
    parser.add_argument("--file_path", required=True, help="数据文件的绝对路径")
    parser.add_argument("--row_questions", required=True, help='行变量 JSON，如 \'["all"]\'')
    parser.add_argument("--col_questions", required=True, help='列变量 JSON，如 \'["Q17.性别"]\'')
    parser.add_argument("--sheet_name", default="0", help="工作表名或编号")
    parser.add_argument("--merge_rules", default=None, help='合并规则 JSON')
    parser.add_argument("--calc_scores", default=None, help='"auto" 或题目列表 JSON')
    parser.add_argument("--output_path", default=None, help="输出 Excel 路径")
    parser.add_argument("--report_json", default="", help="AI 报告 JSON 文本")
    args = parser.parse_args()

    sheet_name = args.sheet_name
    try:
        sheet_name = int(sheet_name)
    except ValueError:
        pass

    try:
        row_questions = json.loads(args.row_questions)
    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"row_questions JSON 解析失败: {e}"}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    try:
        col_questions = json.loads(args.col_questions)
    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"col_questions JSON 解析失败: {e}"}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    merge_rules = None
    if args.merge_rules:
        try:
            merge_rules = json.loads(args.merge_rules)
        except json.JSONDecodeError as e:
            print(json.dumps({"error": f"merge_rules JSON 解析失败: {e}"}, ensure_ascii=False), file=sys.stderr)
            sys.exit(1)

    try:
        result = run_crosstab_pipeline(
            file_path=args.file_path,
            row_questions=row_questions,
            col_questions=col_questions,
            sheet_name=sheet_name,
            merge_rules=merge_rules,
            calc_scores_mode=args.calc_scores,
            output_path=args.output_path,
            report_json=args.report_json,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
