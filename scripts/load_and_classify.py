#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 数据加载与列自动分类
====================================

读取问卷数据文件（CSV / Excel），自动识别并分类所有列为：
- 单选题、多选题、矩阵量表题、文本题、元数据列、排除列

输出 JSON 到 stdout，供后续脚本和 AI 使用。

用法:
    python load_and_classify.py --file_path "C:/xxx/data.xlsx" [--sheet_name 0]
"""

import argparse
import json
import sys
import re
import pandas as pd
from collections import defaultdict
from typing import Optional


# ========================================================================= #
#                           辅助函数
# ========================================================================= #

def _extract_subcol_number(subcol: str, prefix: str) -> int:
    """从多选子列名中提取选项序号"""
    suffix = subcol.split(prefix)[1].strip()
    match = re.search(r'^(\d+)', suffix)
    return int(match.group(1)) if match else 0


def _is_text_column(series: pd.Series, col_name: str) -> bool:
    """判断某列是否为文本题（非结构化答案）"""
    if "输入文本" in col_name:
        return True
    if "非必填" in col_name:
        return True

    non_null = series.dropna()
    if len(non_null) == 0:
        return True

    # 纯数值列（量表题/单选编码）→ 非文本
    try:
        pd.to_numeric(non_null)
        return False
    except (ValueError, TypeError):
        pass

    str_vals = non_null.astype(str)
    avg_len = str_vals.str.len().mean()
    unique_rate = non_null.nunique() / len(non_null) if len(non_null) > 0 else 0

    # 高唯一率 + 平均长度较长 → 文本题
    if unique_rate > 0.6 and avg_len > 8:
        return True
    # 超长文本
    if avg_len > 20:
        return True

    return False


def _is_meta_column(col_name: str) -> bool:
    """
    判断是否为元数据列（非题目列）。
    以 Q+数字 开头的列优先视为题目，除非是附属文本列。
    """
    col_clean = col_name.strip()

    # 以 Q+数字 开头 → 大概率是题目，但先检查个人信息类
    if re.match(r'^Q\d+[\.\s]', col_clean):
        personal_keywords = ["姓名", "手机", "电话", "微信", "邮箱", "称呼",
                             "个人信息", "联系方式", "联系到您"]
        for pk in personal_keywords:
            if pk in col_clean:
                return True
        return False

    # recode_ 开头 → 合并后的列，不算元数据
    if col_clean.startswith('recode_'):
        return False

    # 明确的元数据前缀
    meta_exact_prefixes = ["序号", "开始答题时间", "结束答题时间", "答题时长"]
    for prefix in meta_exact_prefixes:
        if col_clean.startswith(prefix):
            return True

    # 元数据特征关键词
    meta_keywords = [
        "uid", "UID", "IP地址", "来源渠道",
        "怎么称呼", "姓名", "手机号", "电话", "微信号", "邮箱",
    ]
    for kw in meta_keywords:
        if kw in col_clean:
            return True

    # 附属文本列
    if "输入文本" in col_clean:
        return True

    # 非 Q 开头的其余未知列 → 按元数据处理
    if not re.match(r'^Q\d+', col_clean):
        return True

    return False


# ========================================================================= #
#                        核心分类逻辑
# ========================================================================= #

def classify_columns(df: pd.DataFrame) -> dict:
    """
    自动分类所有列。

    返回:
        {
            "single_choice": [...],
            "multi_choice": {root: [subcols]},
            "matrix_scale": {root: [subcols]},
            "text": [...],
            "meta": [...],
            "excluded": [...],
            "valid_for_crosstab": [...]
        }
    """
    # 1. 先识别带冒号的子列根
    multi_roots = defaultdict(list)
    for col in df.columns:
        col_clean = str(col).strip()
        match = re.match(r'^(Q\d+\.)', col_clean)
        if match:
            root = match.group(1)
            rest = col_clean[len(root):]
            if ':' in rest or '：' in rest:
                multi_roots[root].append(col)

    # 2. 区分多选题 vs 矩阵量表题
    multi_choice_dict = {}
    matrix_scale_dict = {}
    multi_subcols = set()
    matrix_subcols = set()

    for root, subcols in multi_roots.items():
        if len(subcols) <= 1:
            continue
        sorted_cols = sorted(subcols, key=lambda x: _extract_subcol_number(x, root))

        is_matrix = False
        sample = df[sorted_cols].head(1000)
        for sc in sorted_cols:
            if "输入文本" in str(sc):
                continue
            try:
                unique_vals = sample[sc].dropna().unique()
                numeric_vals = set()
                for v in unique_vals:
                    try:
                        numeric_vals.add(float(v))
                    except (ValueError, TypeError):
                        pass
                if numeric_vals and max(numeric_vals) > 1:
                    is_matrix = True
                    break
            except Exception:
                pass

        if is_matrix:
            matrix_scale_dict[root] = sorted_cols
            matrix_subcols.update(sorted_cols)
        else:
            multi_choice_dict[root] = sorted_cols
            multi_subcols.update(sorted_cols)

    # 3. 分类其余列
    single_choice = []
    text_cols = []
    meta_cols = []
    excluded_cols = []

    for col in df.columns:
        col_str = str(col).strip()

        if col in multi_subcols:
            if "输入文本" in col_str:
                excluded_cols.append(col)
            continue

        if col in matrix_subcols:
            if "输入文本" in col_str:
                excluded_cols.append(col)
            else:
                single_choice.append(col)
            continue

        if _is_meta_column(col_str):
            meta_cols.append(col)
            continue

        if _is_text_column(df[col], col_str):
            text_cols.append(col)
            continue

        single_choice.append(col)

    # 4. 可用于交叉分析的列
    valid_for_crosstab = list(single_choice)
    for root in multi_choice_dict:
        valid_for_crosstab.append(root)

    return {
        "single_choice": single_choice,
        "multi_choice": multi_choice_dict,
        "matrix_scale": matrix_scale_dict,
        "text": text_cols,
        "meta": meta_cols,
        "excluded": excluded_cols,
        "valid_for_crosstab": valid_for_crosstab,
    }


# ========================================================================= #
#                        数据加载入口
# ========================================================================= #

def load_and_classify(file_path: str, sheet_name=0) -> dict:
    """
    加载问卷数据并自动分类列。

    Args:
        file_path: 数据文件绝对路径
        sheet_name: 工作表名或编号

    Returns:
        包含分类信息的字典
    """
    ext = file_path.rsplit('.', 1)[-1].lower()
    if ext == 'csv':
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

    df.columns = [str(c).strip() for c in df.columns]
    classification = classify_columns(df)

    return {
        "file_path": file_path,
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "single_choice": classification["single_choice"],
        "multi_choice": {
            root: cols for root, cols in classification["multi_choice"].items()
        },
        "matrix_scale": {
            root: cols for root, cols in classification.get("matrix_scale", {}).items()
        },
        "text": classification["text"],
        "meta": classification["meta"],
        "excluded": classification["excluded"],
        "valid_for_crosstab": classification["valid_for_crosstab"],
    }


# ========================================================================= #
#                        CLI 入口
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(description="问卷数据加载与列自动分类")
    parser.add_argument("--file_path", required=True, help="数据文件的绝对路径")
    parser.add_argument("--sheet_name", default="0", help="工作表名称或编号（默认 0）")
    args = parser.parse_args()

    # 处理 sheet_name：如果是纯数字则转为 int
    sheet_name = args.sheet_name
    try:
        sheet_name = int(sheet_name)
    except ValueError:
        pass

    try:
        result = load_and_classify(args.file_path, sheet_name)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except Exception as e:
        error_result = {"error": str(e)}
        print(json.dumps(error_result, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
