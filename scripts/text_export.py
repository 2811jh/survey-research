#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - 文本分析结果导出
================================

将 AI 的文本分析结果导出为专业 Excel 报告。
每道题生成两个 sheet：总结概览 + 逐条明细。

用法:
    python text_export.py --output_path "C:/xxx/data_文本分析.xlsx" --results_file "C:/xxx/text_results.json"

    # 也可以直接传入 JSON 字符串
    python text_export.py --output_path "C:/xxx/data_文本分析.xlsx" --results_json '[{...}]'

results JSON 格式:
    [
        {
            "question": "Q10.您还有什么建议？",
            "conclusion": "核心结论（2-3句话）",
            "dimensions": [
                {
                    "name": "维度名（如：性能问题）",
                    "count": 100,
                    "percentage": "20.0%",
                    "examples": ["用户原文1", "用户原文2", ...]
                }
            ],
            "details": [
                {"text": "用户原文", "labels": "维度A, 维度B"}
            ]
        }
    ]
"""

import argparse
import json
import sys
import os
import re
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _styles import (
    Theme, TextReportTheme,
    format_text_summary_sheet, format_text_detail_sheet,
    thin_border, header_fill, header_font, index_fill, index_font,
    body_font, even_fill, odd_fill, make_fill,
    ALIGN_CENTER, ALIGN_LEFT, ALIGN_RIGHT, ALIGN_TOP_LEFT,
)
from text_extract import clean_column_texts
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# ========================================================================= #
#                        Excel 导出核心
# ========================================================================= #

def _safe_sheet_name(name: str, max_len: int = 28) -> str:
    """生成安全的 sheet 名称（去除非法字符，截断长度）"""
    # 去除 Excel sheet 名非法字符
    cleaned = re.sub(r'[\\/:*?"<>|]', '', name)
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len]
    return cleaned.strip()


def _auto_label_texts(texts: list, dimensions: list) -> list:
    """
    基于维度名称中的关键词，自动为每条文本标注所属维度。

    逻辑：
    1. 从每个维度名中提取关键词（按 / ( ) 、 分割）
    2. 同时从 examples 中提取高频词作为辅助关键词
    3. 逐条文本匹配，命中关键词最多的维度即为标签
    4. 支持多标签（一条文本可属于多个维度）

    Args:
        texts: 清洗后的全量文本列表
        dimensions: 维度列表，每项含 name, examples 等

    Returns:
        details 列表: [{"text": "...", "labels": "维度A, 维度B"}, ...]
    """
    if not dimensions or not texts:
        return []

    # 为每个维度构建关键词集合
    dim_keywords = []
    for dim in dimensions:
        name = dim.get("name", "")
        # 从维度名提取关键词：按常见分隔符拆分
        keywords = set()
        parts = re.split(r'[/（）()、,，\s]+', name)
        for p in parts:
            p = p.strip()
            if len(p) >= 2:  # 至少2个字的词才有效
                keywords.add(p)

        # 从 examples 中提取辅助关键词（取每条的前4个字作为短语特征）
        for ex in dim.get("examples", [])[:5]:
            ex_clean = ex.strip()
            if len(ex_clean) >= 2:
                # 提取开头的关键动词短语（如"优化卡顿" -> "优化", "卡顿"）
                short_words = re.findall(r'[\u4e00-\u9fff]{2,4}', ex_clean)
                for w in short_words[:3]:
                    keywords.add(w)

        dim_keywords.append({
            "name": name,
            "keywords": keywords,
        })

    # 逐条文本匹配
    details = []
    for text in texts:
        matched_dims = []
        for dk in dim_keywords:
            # 统计匹配到的关键词数量
            hit_count = sum(1 for kw in dk["keywords"] if kw in text)
            if hit_count > 0:
                matched_dims.append((dk["name"], hit_count))

        if matched_dims:
            # 按命中数降序，取所有命中的维度
            matched_dims.sort(key=lambda x: -x[1])
            labels = ", ".join(d[0] for d in matched_dims)
        else:
            labels = "其他"

        details.append({"text": text, "labels": labels})

    return details




def _write_summary_sheet(writer, question_data: dict, sheet_idx: int):
    """
    写入单题的总结概览 sheet（纯 openpyxl 手写竖排布局）。

    布局：
      第1行：大标题 "文本分析报告"
      第2行：题目名称
      第3行：核心结论
      第4行：空行
      第5行：维度表表头（序号 | 问题类别 | 反馈条数 | 占比 | 典型用户原文）
      第6-N行：维度数据行，examples 用 bullet list 换行展示
    """
    question = question_data.get("question", f"题目{sheet_idx}")
    conclusion = question_data.get("conclusion", "")
    dimensions = question_data.get("dimensions", [])

    sheet_name = "总结概览"

    # 创建 sheet（不用 pandas）
    ws = writer.book.create_sheet(sheet_name)
    border = thin_border()
    total_width = 5  # 序号 | 问题类别 | 反馈条数 | 占比 | 典型用户原文

    row = 1

    # ---- 第1行：大标题 ----
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_width)
    cell = ws.cell(row=row, column=1, value="文本分析报告")
    cell.fill = make_fill(Theme.HEADER_BG)
    cell.font = Font(name=Theme.FONT_NAME, size=16, bold=True, color=Theme.HEADER_FONT)
    cell.alignment = ALIGN_CENTER
    cell.border = border
    ws.row_dimensions[row].height = 42
    for c in range(2, total_width + 1):
        ws.cell(row=row, column=c).border = border
    row += 1

    # ---- 第2行：题目名称 ----
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_width)
    cell = ws.cell(row=row, column=1, value=f"题目：{question}")
    cell.fill = make_fill(TextReportTheme.DIMENSION_HEADER_BG)
    cell.font = Font(name=Theme.FONT_NAME, size=12, bold=True, color=TextReportTheme.DIMENSION_HEADER_FONT)
    cell.alignment = ALIGN_LEFT
    cell.border = border
    ws.row_dimensions[row].height = 32
    for c in range(2, total_width + 1):
        ws.cell(row=row, column=c).border = border
    row += 1

    # ---- 第3行：核心结论 ----
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_width)
    # 清洗 %% → %（防止双重转义）
    conclusion_clean = conclusion.replace("%%", "%") if conclusion else ""
    conclusion_display = f"核心结论：{conclusion_clean}" if conclusion_clean else ""
    cell = ws.cell(row=row, column=1, value=conclusion_display)
    cell.fill = make_fill(TextReportTheme.CONCLUSION_BG)
    cell.font = Font(name=Theme.FONT_NAME, size=11, bold=True, color=TextReportTheme.CONCLUSION_FONT)
    cell.alignment = ALIGN_TOP_LEFT
    cell.border = border
    line_count = max(1, len(conclusion_display) // 80 + 1)
    ws.row_dimensions[row].height = max(60, line_count * 20)
    for c in range(2, total_width + 1):
        ws.cell(row=row, column=c).border = border
    row += 1

    # ---- 第4行：空行 ----
    ws.row_dimensions[row].height = 10
    row += 1

    # ---- 维度统计表 ----
    if dimensions:
        # 表头
        headers = ["序号", "问题类别", "反馈条数", "占比", "典型用户原文"]
        dim_header_fill = make_fill(TextReportTheme.DIMENSION_HEADER_BG)
        dim_header_font = Font(name=Theme.FONT_NAME, size=11, bold=True, color=TextReportTheme.DIMENSION_HEADER_FONT)
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=ci, value=h)
            cell.fill = dim_header_fill
            cell.font = dim_header_font
            cell.alignment = ALIGN_CENTER
            cell.border = border
        ws.row_dimensions[row].height = 30
        row += 1

        # 数据行
        for di, dim in enumerate(dimensions, 1):
            examples = dim.get("examples", [])
            # 用 bullet list 换行连接全部 examples
            example_text = "\n".join(f"• {ex}" for ex in examples)

            # 序号
            cell = ws.cell(row=row, column=1, value=di)
            cell.fill = even_fill() if di % 2 == 0 else odd_fill()
            cell.font = body_font()
            cell.alignment = ALIGN_CENTER
            cell.border = border

            # 问题类别
            cell = ws.cell(row=row, column=2, value=dim.get("name", ""))
            cell.fill = index_fill()
            cell.font = index_font(bold=True)
            cell.alignment = ALIGN_LEFT
            cell.border = border

            # 反馈条数
            cell = ws.cell(row=row, column=3, value=dim.get("count", 0))
            cell.fill = even_fill() if di % 2 == 0 else odd_fill()
            cell.font = body_font()
            cell.alignment = ALIGN_CENTER
            cell.border = border

            # 占比（清洗 %% → %）
            pct_val = str(dim.get("percentage", "0%")).replace("%%", "%")
            cell = ws.cell(row=row, column=4, value=pct_val)
            cell.fill = even_fill() if di % 2 == 0 else odd_fill()
            cell.font = body_font()
            cell.alignment = ALIGN_CENTER
            cell.border = border

            # 典型用户原文（bullet list）
            cell = ws.cell(row=row, column=5, value=example_text)
            cell.fill = even_fill() if di % 2 == 0 else odd_fill()
            cell.font = body_font()
            cell.alignment = ALIGN_TOP_LEFT
            cell.border = border

            # 根据 examples 数量自动调整行高（每条原声约 18px）
            example_lines = max(1, len(examples))
            ws.row_dimensions[row].height = max(28, example_lines * 18)
            row += 1

    # ---- 列宽 ----
    ws.column_dimensions['A'].width = 8   # 序号
    ws.column_dimensions['B'].width = 30  # 问题类别
    ws.column_dimensions['C'].width = 12  # 反馈条数
    ws.column_dimensions['D'].width = 10  # 占比
    ws.column_dimensions['E'].width = 70  # 典型用户原文

    ws.sheet_properties.tabColor = "C6EFCE"
    ws.sheet_view.showGridLines = False

    return sheet_name


def _write_detail_sheet(writer, question_data: dict, sheet_idx: int):
    """
    写入单题的逐条明细 sheet（纯 openpyxl 手写，含序号列）。

    表头：序号 | 用户原文 | 归属类别
    """
    question = question_data.get("question", f"题目{sheet_idx}")
    details = question_data.get("details", [])

    sheet_name = "逐条明细"

    if not details:
        ws = writer.book.create_sheet(sheet_name)
        ws.cell(row=1, column=1, value="暂无明细数据")
        ws.sheet_properties.tabColor = "F8CBAD"
        return sheet_name

    ws = writer.book.create_sheet(sheet_name)
    border = thin_border()

    # ---- 表头 ----
    headers = ["序号", "用户原文", "归属类别"]
    detail_header_fill = make_fill(TextReportTheme.DETAIL_HEADER_BG)
    detail_header_font = header_font(size=10)
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = detail_header_fill
        cell.font = detail_header_font
        cell.alignment = ALIGN_CENTER
        cell.border = border
    ws.row_dimensions[1].height = 35

    # ---- AutoFilter（表头筛选箭头）----
    ws.auto_filter.ref = f"A1:C{len(details) + 1}"

    # ---- 数据行 ----
    for ri, item in enumerate(details, 1):
        row_idx = ri + 1

        # 序号
        cell = ws.cell(row=row_idx, column=1, value=ri)
        cell.fill = even_fill() if ri % 2 == 0 else odd_fill()
        cell.font = body_font()
        cell.alignment = ALIGN_CENTER
        cell.border = border

        # 用户原文
        cell = ws.cell(row=row_idx, column=2, value=item.get("text", ""))
        cell.fill = even_fill() if ri % 2 == 0 else odd_fill()
        cell.font = body_font()
        cell.alignment = ALIGN_TOP_LEFT
        cell.border = border

        # 归属类别
        cell = ws.cell(row=row_idx, column=3, value=item.get("labels", ""))
        cell.fill = even_fill() if ri % 2 == 0 else odd_fill()
        cell.font = body_font()
        cell.alignment = ALIGN_CENTER
        cell.border = border

        ws.row_dimensions[row_idx].height = 40

    # ---- 列宽 ----
    ws.column_dimensions['A'].width = 8    # 序号
    ws.column_dimensions['B'].width = 80   # 用户原文
    ws.column_dimensions['C'].width = 30   # 归属类别

    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "F8CBAD"

    return sheet_name


# ========================================================================= #
#                        主函数
# ========================================================================= #

def export_text_report(results: list, output_path: str,
                       file_path: str = None, sheet_name=0) -> dict:
    """
    将文本分析结果导出为专业 Excel 报告。

    当 details 为空但提供了 file_path 时，自动从原始数据中提取文本，
    基于 dimensions 关键词做自动标注，生成逐条明细。

    Args:
        results: 分析结果列表
        output_path: 输出文件路径
        file_path: 原始数据文件路径（可选，用于自动标注）
        sheet_name: 工作表名或编号（默认 0）

    Returns:
        {"status": "success", "output_path": str, "sheets": [str]}
    """
    if not results:
        return {"error": "results 不能为空"}

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except PermissionError:
            raise PermissionError(f"请关闭正在使用的文件：{output_path}")

    # 如果需要自动标注，预加载原始数据
    source_df = None
    if file_path and os.path.exists(file_path):
        try:
            ext = file_path.rsplit('.', 1)[-1].lower()
            if ext == 'csv':
                source_df = pd.read_csv(file_path)
            else:
                source_df = pd.read_excel(file_path, sheet_name=sheet_name)
            source_df.columns = [str(c).strip() for c in source_df.columns]
        except Exception:
            source_df = None

    sheets_created = []
    auto_labeled_count = 0

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for idx, question_data in enumerate(results, 1):
            details = question_data.get("details", [])
            dimensions = question_data.get("dimensions", [])
            question_col = question_data.get("question", "")

            # 如果 details 为空，尝试自动标注
            if not details and dimensions and source_df is not None and question_col:
                # 尝试精确匹配列名，失败则模糊匹配
                matched_col = question_col if question_col in source_df.columns else None
                if not matched_col:
                    # 提取题号前缀（如 "Q20"）用于模糊匹配
                    q_prefix = question_col.split(".")[0] if "." in question_col else None
                    for c in source_df.columns:
                        if q_prefix and str(c).startswith(q_prefix + "."):
                            matched_col = c
                            break
                        elif question_col in str(c) or str(c) in question_col:
                            matched_col = c
                            break

                if matched_col:
                    extract_result = clean_column_texts(source_df, matched_col)
                    if "error" not in extract_result:
                        all_texts = extract_result.get("texts", [])
                        if all_texts:
                            details = _auto_label_texts(all_texts, dimensions)
                            question_data["details"] = details
                            auto_labeled_count += len(details)

            # 总结概览
            summary_name = _write_summary_sheet(writer, question_data, idx)
            sheets_created.append(summary_name)

            # 逐条明细
            detail_name = _write_detail_sheet(writer, question_data, idx)
            sheets_created.append(detail_name)

    result = {
        "status": "success",
        "output_path": output_path,
        "sheets": sheets_created,
        "questions_count": len(results),
    }
    if auto_labeled_count > 0:
        result["auto_labeled_count"] = auto_labeled_count

    return result


# ========================================================================= #
#                        CLI 入口
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(description="问卷文本分析结果导出")
    parser.add_argument("--output_path", required=True, help="输出 Excel 文件路径")
    parser.add_argument("--results_file", default=None, help="分析结果 JSON 文件路径")
    parser.add_argument("--results_json", default=None, help="分析结果 JSON 字符串")
    parser.add_argument("--file_path", default=None,
                        help="原始数据文件路径（可选）。当 details 为空时，从此文件提取文本并自动标注")
    parser.add_argument("--sheet_name", default="0", help="工作表名或编号（默认 0）")
    args = parser.parse_args()

    # 解析 sheet_name
    sheet_name = args.sheet_name
    try:
        sheet_name = int(sheet_name)
    except ValueError:
        pass

    # 读取分析结果
    results = None
    if args.results_file:
        try:
            with open(args.results_file, 'r', encoding='utf-8') as f:
                results = json.load(f)
        except Exception as e:
            print(json.dumps({"error": f"读取 results_file 失败: {e}"}, ensure_ascii=False), file=sys.stderr)
            sys.exit(1)
    elif args.results_json:
        try:
            results = json.loads(args.results_json)
        except json.JSONDecodeError as e:
            print(json.dumps({"error": f"results_json 解析失败: {e}"}, ensure_ascii=False), file=sys.stderr)
            sys.exit(1)
    else:
        print(json.dumps({"error": "请提供 --results_file 或 --results_json"}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)

    if not isinstance(results, list):
        results = [results]

    try:
        result = export_text_report(results, args.output_path,
                                    file_path=args.file_path,
                                    sheet_name=sheet_name)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
