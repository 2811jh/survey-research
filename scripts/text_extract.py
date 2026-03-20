#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 文本题提取与清洗
================================

从问卷数据中提取指定文本题的回答，自动过滤无效文本，输出清洗后的 JSON。

功能：
- 自动检测候选文本题列表（当不指定 column 时）
- 文本清洗：过滤无效回答（"无"、"没有"、纯数字等）
- 支持抽样（大数据量时先看一部分）

用法:
    # 检测所有文本题
    python text_extract.py --file_path "C:/xxx/data.xlsx" --detect

    # 提取指定列的文本
    python text_extract.py --file_path "C:/xxx/data.xlsx" --column "Q10.您还有什么建议？"

    # 抽样 200 条
    python text_extract.py --file_path "C:/xxx/data.xlsx" --column "Q10.建议" --sample_n 200
"""

import argparse
import json
import sys
import os
import re
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_and_classify import classify_columns


# ========================================================================= #
#                        辅助函数
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


# ========================================================================= #
#                        文本清洗
# ========================================================================= #

# 无效文本黑名单（精确匹配）
_INVALID_EXACT = {
    "无", "没有", "没", "无意见", "没意见", "无建议", "没建议",
    "暂无", "暂时没有", "暂时无", "暂无建议", "暂无意见",
    "无所谓", "没什么", "没啥", "不知道", "不清楚",
    "没有了", "无了", "就这些", "以上", "同上",
    "好", "好的", "可以", "还好", "一般", "ok", "OK", "Ok",
    "都很好", "很好", "非常好", "挺好的",
    "不错", "还不错", "可以的", "都行",
    "嗯", "额", "呃", "哦", "啊",
    ".", "..", "...", "。", "-", "--", "——",
    "n/a", "N/A", "na", "NA", "null", "none", "None",
}

# 无效文本正则（匹配则过滤）
_INVALID_PATTERNS = [
    re.compile(r'^\d+$'),                    # 纯数字
    re.compile(r'^[.\-_=+*#@!？！。，、]+$'),  # 纯标点符号
    re.compile(r'^\s*$'),                     # 纯空白
]


def _clean_text(text: str) -> str:
    """基础清理：去除多余空白"""
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _is_invalid_text(text: str) -> bool:
    """判断是否为无效文本"""
    cleaned = text.strip()

    # 精确匹配黑名单
    if cleaned in _INVALID_EXACT:
        return True

    # 正则匹配
    for pattern in _INVALID_PATTERNS:
        if pattern.match(cleaned):
            return True

    # 太短（1个字符）
    if len(cleaned) <= 1:
        return True

    return False


def clean_column_texts(df: pd.DataFrame, column: str, sample_n: int = 0, sample_seed: int = 42) -> dict:
    """
    对指定列做文本清洗。

    Args:
        df: 数据 DataFrame
        column: 列名
        sample_n: 抽样数量（0=全部）
        sample_seed: 随机种子

    Returns:
        {
            "column": str,
            "raw_count": int,         # 原始非空条数
            "valid_count": int,       # 清洗后有效条数
            "unique_count": int,      # 去重后条数
            "dropped_count": int,     # 被过滤的无效条数
            "sample_n": int,          # 实际返回的文本数量
            "texts": [str],           # 文本列表
            "unique_texts": [str],    # 去重后的文本列表
        }
    """
    if column not in df.columns:
        return {"error": f"列 '{column}' 不存在"}

    # 原始非空
    raw_series = df[column].dropna().astype(str)
    raw_count = len(raw_series)

    # 清洗
    valid_texts = []
    for text in raw_series:
        cleaned = _clean_text(text)
        if cleaned and not _is_invalid_text(cleaned):
            valid_texts.append(cleaned)

    valid_count = len(valid_texts)
    dropped_count = raw_count - valid_count

    # 去重
    unique_texts = list(dict.fromkeys(valid_texts))  # 保序去重
    unique_count = len(unique_texts)

    # 抽样
    output_texts = valid_texts
    if sample_n > 0 and sample_n < len(valid_texts):
        import random
        rng = random.Random(sample_seed)
        output_texts = rng.sample(valid_texts, sample_n)

    return {
        "column": column,
        "raw_count": raw_count,
        "valid_count": valid_count,
        "unique_count": unique_count,
        "dropped_count": dropped_count,
        "sample_n": len(output_texts),
        "texts": output_texts,
        "unique_texts": unique_texts,
    }


# ========================================================================= #
#                    自动检测文本题候选列
# ========================================================================= #

def detect_text_questions(df: pd.DataFrame, classification: dict = None) -> list:
    """
    自动检测问卷中的文本题候选列。

    Returns:
        列表，每项:
        {
            "index": int,
            "column": str,
            "raw_n": int,
            "valid_n": int,
            "avg_len": float,
            "uniq_rate": float,
            "example": str,
        }
    """
    if classification is None:
        classification = classify_columns(df)

    text_cols = classification.get("text", [])
    results = []

    for idx, col in enumerate(text_cols, 1):
        raw_series = df[col].dropna().astype(str)
        raw_n = len(raw_series)

        # 清洗
        valid_texts = []
        for text in raw_series:
            cleaned = _clean_text(text)
            if cleaned and not _is_invalid_text(cleaned):
                valid_texts.append(cleaned)

        valid_n = len(valid_texts)
        avg_len = np.mean([len(t) for t in valid_texts]) if valid_texts else 0
        uniq_rate = len(set(valid_texts)) / valid_n if valid_n > 0 else 0
        example = valid_texts[0] if valid_texts else ""

        results.append({
            "index": idx,
            "column": col,
            "raw_n": raw_n,
            "valid_n": valid_n,
            "avg_len": round(avg_len, 1),
            "uniq_rate": round(uniq_rate, 3),
            "example": example[:100],
        })

    return results


# ========================================================================= #
#                        CLI 入口
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(description="问卷文本题提取与清洗")
    parser.add_argument("--file_path", required=True, help="数据文件的绝对路径")
    parser.add_argument("--column", default=None, help="要提取的列名（不指定则用 --detect 模式）")
    parser.add_argument("--detect", action="store_true", help="自动检测文本题候选列")
    parser.add_argument("--sheet_name", default="0", help="工作表名或编号")
    parser.add_argument("--sample_n", type=int, default=0, help="抽样数量（0=全部）")
    parser.add_argument("--sample_seed", type=int, default=42, help="随机种子")
    args = parser.parse_args()

    sheet_name = args.sheet_name
    try:
        sheet_name = int(sheet_name)
    except ValueError:
        pass

    try:
        # 加载数据
        ext = args.file_path.rsplit('.', 1)[-1].lower()
        if ext == 'csv':
            df = pd.read_csv(args.file_path, encoding=_detect_csv_encoding(args.file_path))
        else:
            df = pd.read_excel(args.file_path, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]

        if args.detect or args.column is None:
            # 检测模式
            classification = classify_columns(df)
            questions = detect_text_questions(df, classification)
            result = {
                "file_path": args.file_path,
                "total_rows": len(df),
                "text_questions": questions,
            }
        else:
            # 提取模式
            result = clean_column_texts(df, args.column, args.sample_n, args.sample_seed)

        print(json.dumps(result, ensure_ascii=False, indent=2))

    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
