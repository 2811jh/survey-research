#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 数据字段扩充
============================

在问卷 CSV/Excel 的指定日期列右侧，自动插入派生字段：
  --add-week   插入"周"列：第X周（M.D-M.D），ISO周，周一为起始
  --add-month  插入"月"列：26年X月（两位年份 + 月份）

可同时指定多个日期列，每列单独在其右侧插入。

用法:
    python enrich_columns.py \\
        --file_path "C:/xxx/量化数据.csv" \\
        --date_cols "结束答题时间" \\
        --add-week --add-month \\
        [--output "C:/xxx/输出.csv"]

    # 多列
    python enrich_columns.py \\
        --file_path "data.csv" \\
        --date_cols "开始答题时间" "结束答题时间" \\
        --add-week --add-month

输出:
    成功 → stdout 输出 JSON，同时将文件写入 --output 路径（默认在原文件名加 _带周月 后缀）
    失败 → stderr 输出错误信息，exit code 1
"""

import argparse
import json
import os
import sys
from datetime import timedelta

import pandas as pd


# ========================================================================= #
#                           编码检测
# ========================================================================= #

def _detect_csv_encoding(filepath: str, sample_size: int = 8192) -> str:
    """检测 CSV 文件编码，优先返回 utf-8-sig / utf-8 / gbk"""
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
#                           格式化函数
# ========================================================================= #

def _fmt_md(d) -> str:
    """date → M.D（不补零）"""
    return f'{d.month}.{d.day}'


def week_label(dt) -> str:
    """
    datetime → 第X周（M.D-M.D）
    ISO 周编号（周一为每周第一天）
    示例：2026-04-06 → 第15周（4.6-4.12）
    """
    iso = dt.isocalendar()          # (year, week, weekday)  weekday: 1=Mon
    week_num = iso[1]
    monday = dt - timedelta(days=iso[2] - 1)
    sunday = monday + timedelta(days=6)
    return f'第{week_num}周（{_fmt_md(monday)}-{_fmt_md(sunday)}）'


def month_label(dt) -> str:
    """
    datetime → 26年X月
    取年份后两位 + 月份数字
    示例：2026-04-06 → 26年4月
    """
    year_short = str(dt.year)[2:]
    return f'{year_short}年{dt.month}月'


# ========================================================================= #
#                           核心逻辑
# ========================================================================= #

def enrich(
    file_path: str,
    date_cols: list[str],
    add_week: bool,
    add_month: bool,
    output: str | None = None,
) -> dict:
    """
    读取文件，为每个 date_col 在其右侧插入周/月列，写出新文件。

    Returns dict with status / output_path / stats
    """
    if not add_week and not add_month:
        return {"status": "error", "message": "请至少指定 --add-week 或 --add-month 之一"}

    # ── 读取 ──────────────────────────────────────────────────────────────
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        df = pd.read_excel(file_path, dtype=str)
    else:
        enc = _detect_csv_encoding(file_path)
        df = pd.read_csv(file_path, encoding=enc, dtype=str, low_memory=False)

    original_rows = len(df)

    # ── 验证列是否存在 ────────────────────────────────────────────────────
    missing = [c for c in date_cols if c not in df.columns]
    if missing:
        return {
            "status": "error",
            "message": f"以下列不存在：{missing}，可用列（前20）：{list(df.columns[:20])}",
        }

    # ── 逐列插入（倒序插入，避免偏移） ───────────────────────────────────
    # 先正序处理，收集 (insert_pos, col_name, series)，再倒序插入
    inserts = []
    week_counts: dict[str, dict] = {}

    for date_col in date_cols:
        dt_series = pd.to_datetime(df[date_col], errors='coerce')
        null_cnt = int(dt_series.isna().sum())

        col_idx = df.columns.get_loc(date_col)
        offset = 1  # 第一个新列插在 date_col 右侧

        if add_week:
            w_series = dt_series.apply(lambda d: week_label(d) if pd.notna(d) else '')
            inserts.append((col_idx + offset, f'{date_col}_周', w_series))
            offset += 1

        if add_month:
            m_series = dt_series.apply(lambda d: month_label(d) if pd.notna(d) else '')
            inserts.append((col_idx + offset, f'{date_col}_月', m_series))
            offset += 1

        # 统计各周分布（用于报告）
        if add_week:
            w_counts = w_series[w_series != ''].value_counts().to_dict()
            week_counts[date_col] = w_counts

    # 倒序插入，保证前面插入不影响后面位置
    for pos, col_name, series in sorted(inserts, key=lambda x: x[0], reverse=True):
        df.insert(pos, col_name, series)

    # ── 输出路径 ──────────────────────────────────────────────────────────
    if output is None:
        base, suf = os.path.splitext(file_path)
        output = base + '_带周月' + (suf if suf in ('.xlsx', '.xls') else '.csv')

    # ── 写出 ──────────────────────────────────────────────────────────────
    out_ext = os.path.splitext(output)[1].lower()
    if out_ext in ('.xlsx', '.xls'):
        df.to_excel(output, index=False)
    else:
        df.to_csv(output, index=False, encoding='utf-8-sig')

    file_size = os.path.getsize(output)

    return {
        "status": "success",
        "output_path": output,
        "rows": original_rows,
        "file_size_bytes": file_size,
        "columns_added": [col for _, col, _ in inserts],
        "week_distribution": week_counts,
    }


# ========================================================================= #
#                           CLI 入口
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(
        description='在日期列右侧插入"周"/"月"派生字段',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('--file_path', required=True, help='输入文件路径（CSV 或 Excel）')
    parser.add_argument(
        '--date_cols', nargs='+', default=['结束答题时间'],
        help='要处理的日期列名，可多个（默认：结束答题时间）',
    )
    parser.add_argument('--add-week',  action='store_true', help='插入"周"列')
    parser.add_argument('--add-month', action='store_true', help='插入"月"列')
    parser.add_argument('--output', default=None, help='输出路径（默认：原文件名 + _带周月）')

    # 默认两者都加（与原始用法一致）
    args = parser.parse_args()
    if not args.add_week and not args.add_month:
        args.add_week = True
        args.add_month = True

    result = enrich(
        file_path=args.file_path,
        date_cols=args.date_cols,
        add_week=args.add_week,
        add_month=args.add_month,
        output=args.output,
    )

    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result.get('status') == 'error':
        sys.exit(1)


if __name__ == '__main__':
    main()
