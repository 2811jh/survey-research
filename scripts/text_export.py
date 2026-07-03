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


def _clean_filename_part(text: str, max_len: int = 24) -> str:
    """生成安全的文件名片段。"""
    cleaned = re.sub(r'[\\/:*?"<>|]', '', text)
    cleaned = re.sub(r'\s+', '', cleaned)
    return cleaned[:max_len].strip() or "文本分析"


def _summarize_question_for_filename(question: str) -> str:
    """根据题目生成简短文件名语义。"""
    q_match = re.match(r'^(Q\d+)', question.strip())
    qid = q_match.group(1) if q_match else "文本题"

    activity_match = re.search(r'[“"]([^”"]+)[”"]', question)
    activity = activity_match.group(1) if activity_match else ""
    if activity:
        activity = activity.replace("活动", "").replace("节", "")

    if "MC移动版" in question and "比较满意" in question:
        topic = "MC满意原因"
    elif "外挂" in question:
        topic = "外挂表现"
    elif "模组" in question and "搜不到" in question:
        topic = "想玩缺失模组"
    elif "BUG" in question.upper():
        topic = "近期BUG"
    elif "丢失" in question and "存档" in question:
        topic = "存档丢失场景"
    elif "美术" in question or "画面" in question:
        topic = "美术不满原因"
    elif "当前版本" in question and "建议" in question:
        topic = "版本建议"
    elif "性能问题" in question and "哪些模组" in question:
        topic = "性能问题模组"
    elif "性能问题" in question:
        topic = "性能问题场景"
    elif "不太愿意" in question and "推荐" in question:
        topic = "不推荐原因"
    elif "更愿意" in question and "推荐" in question:
        topic = "推荐改进建议"
    elif "推荐给他人" in question and "打动" in question:
        topic = "推荐打动原因"
    elif "模组售后" in question or "问题反馈" in question:
        topic = "模组售后诉求"
    elif activity:
        if "比较满意" in question or "较为满意" in question:
            topic = f"{activity}满意原因"
        elif "不太满意" in question or "不满" in question or "一般" in question:
            topic = f"{activity}不满原因"
        elif "建议" in question or "期待" in question or "意见" in question:
            topic = f"{activity}建议"
        else:
            topic = activity
    elif "建议" in question or "期待" in question or "意见" in question:
        topic = "建议"
    elif "满意" in question:
        topic = "满意原因"
    elif "不满" in question or "一般" in question:
        topic = "不满原因"
    else:
        topic = "文本分析"

    return f"{qid}_{_clean_filename_part(topic)}.xlsx"


def default_output_filename(results: list) -> str:
    """根据文本分析结果生成默认 Excel 文件名。"""
    if not results:
        return "文本分析.xlsx"
    question = str(results[0].get("question", "文本分析"))
    return _summarize_question_for_filename(question)


# ========================================================================= #
#                        Excel 导出核心
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

        # 从 examples 中提取辅助关键词（取每条前5个2-4字词）
        for ex in dim.get("examples", []):
            ex_clean = ex.strip()
            if len(ex_clean) >= 2:
                short_words = re.findall(r'[\u4e00-\u9fff]{2,4}', ex_clean)
                for w in short_words[:5]:
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




OTHER_CANON = "其他/未归类"
_OTHER_ALIASES = {
    "其他", "其它", "无效", "模糊", "无效/模糊", "模糊/无效",
    "其他/无效", "无效/其他", "未归类", "其他/未归类", "无",
}


def _split_labels(s):
    """把 labels 字符串拆成标签列表，兼容中英逗号、顿号分隔。"""
    return [x.strip() for x in re.split(r'[,，、]', str(s or "")) if x.strip()]


def _canon_label(lab):
    """归一化标签：所有"其他/无效"类别名合并为统一的 OTHER_CANON。"""
    lab = (lab or "").strip()
    if not lab or lab in _OTHER_ALIASES:
        return OTHER_CANON
    return lab


def _rebuild_dimensions_from_details(details: list, ai_dimensions: list) -> list:
    """
    以逐条 details 为唯一数据源，重算各维度 count/percentage。

    - 合并所有"其他/无效"类别为统一桶（OTHER_CANON），并显式保留展示（不再隐藏）；
    - "其他/未归类"永远排在最后；
    - 维度顺序优先沿用 AI 给出的顺序，details 中出现的新标签追加其后；
    - examples 优先用 AI 提供的典型原声，缺失时回退为该标签下真实命中的原文（可验证）。

    多标签文本会分别计入各标签，故占比之和可能 > 100%（正常现象）。
    """
    total = len(details)
    counts = {}
    ex_by = {}
    seen_order = []
    for d in details:
        seen = set()
        for lab in _split_labels(d.get("labels", "")):
            canon = _canon_label(lab)
            if canon in seen:
                continue
            seen.add(canon)
            if canon not in counts:
                counts[canon] = 0
                ex_by[canon] = []
                seen_order.append(canon)
            counts[canon] += 1
            txt = (d.get("text") or "").strip()
            if txt and len(ex_by[canon]) < 3:
                ex_by[canon].append(txt)

    ai_examples = {}
    ai_names = []
    for dim in ai_dimensions or []:
        nm = dim.get("name", "")
        if _canon_label(nm) == OTHER_CANON:
            continue
        ai_names.append(nm)
        ai_examples[nm] = dim.get("examples", []) or []

    final_order = []
    for nm in ai_names:
        if nm in counts and nm not in final_order:
            final_order.append(nm)
    for nm in seen_order:
        if nm != OTHER_CANON and nm not in final_order:
            final_order.append(nm)
    if OTHER_CANON in counts:
        final_order.append(OTHER_CANON)

    dims = []
    for nm in final_order:
        c = counts.get(nm, 0)
        pct = f"{c / total * 100:.1f}%" if total > 0 else "0%"
        examples = ai_examples.get(nm) or ex_by.get(nm, [])
        dims.append({"name": nm, "count": c, "percentage": pct, "examples": examples})
    return dims


def _augment_conclusion(conclusion: str, dims: list, total: int) -> str:
    """在结论末尾追加脚本自动计算的数据佐证行（数字由脚本回填，杜绝手写幻觉）。"""
    non_other = [d for d in dims if d["name"] != OTHER_CANON]
    top = sorted(non_other, key=lambda d: -d.get("count", 0))[:4]
    parts = [f"{d['name']} {d['percentage']}（{d['count']}条）" for d in top]
    other = next((d for d in dims if d["name"] == OTHER_CANON), None)
    stat = f"【数据佐证（脚本自动计算，N={total}）】" + "；".join(parts)
    if other:
        stat += f"；{OTHER_CANON} {other['percentage']}（{other['count']}条）"
    base = (conclusion or "").strip()
    return (base + "\n" + stat) if base else stat


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
                       file_path: str = None, sheet_name=0,
                       sample_n: int = 300) -> dict:
    """
    将文本分析结果导出为专业 Excel 报告。

    数据源单一化：一律以逐条 details 为唯一数据源反算维度统计（count/percentage）
    并回填结论中的数字，确保总结概览与逐条明细完全一致、结论不出现幻觉数字。

    details 缺失时的策略（硬制）：
      - sample_n > 0（抽样模式，默认）：必须由 AI 逐条编码填入 details，否则直接
        返回 error（error_type=missing_details），不生成表格。
      - sample_n == 0（全量模式）：允许用 file_path + dimensions 关键词自动兜底
        （覆盖率较低，结果带 keyword_fallback_used 标记）。

    Args:
        results: 分析结果列表
        output_path: 输出文件路径
        file_path: 原始数据文件路径（仅全量兜底模式用于关键词自动标注）
        sheet_name: 工作表名或编号（默认 0）
        sample_n: 明细行数限制（默认 300，0=全量关键词兜底）

    Returns:
        {"status": "success", "output_path": str, "sheets": [str]}
        或 {"status": "error", "error_type": "missing_details", ...}
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
                source_df = pd.read_csv(file_path, encoding=_detect_csv_encoding(file_path))
            else:
                source_df = pd.read_excel(file_path, sheet_name=sheet_name)
            source_df.columns = [str(c).strip() for c in source_df.columns]
        except Exception:
            source_df = None

    # ---- 硬制校验：抽样模式必须提供 details（逐条 LLM 编码），否则拒绝出表 ----
    missing_q = None
    for q in results:
        if q.get("details"):
            continue
        if sample_n == 0 and q.get("dimensions") and source_df is not None:
            continue  # 全量模式允许关键词兜底
        missing_q = q
        break
    if missing_q is not None:
        return {
            "status": "error",
            "error_type": "missing_details",
            "question": missing_q.get("question", ""),
            "message": (
                "抽样模式下必须先为每条抽样文本做逐条编码，再把结果填入 results JSON 的 details "
                "（格式 [{\"text\":\"用户原文\",\"labels\":\"维度A, 维度B\"}]）后才能导出。"
                "这样总结概览与逐条明细都以 AI 的编码为唯一数据源，准确度最高。"
                "如确需关键词自动兜底，请显式加 --sample_n 0 走全量标注（覆盖率较低）。"
            ),
        }

    sheets_created = []
    keyword_fallback_used = False
    warnings = []

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for idx, question_data in enumerate(results, 1):
            details = question_data.get("details", [])
            dimensions = question_data.get("dimensions", [])
            question_col = question_data.get("question", "")

            # 全量模式且 AI 未填 details → 关键词兜底（带告警）
            if not details and sample_n == 0 and dimensions and source_df is not None and question_col:
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
                            keyword_fallback_used = True

            # ---- 单一数据源：一律从 details 反算维度统计 + 脚本回填结论数字 ----
            if details:
                rebuilt = _rebuild_dimensions_from_details(details, dimensions)
                question_data["dimensions"] = rebuilt
                question_data["conclusion"] = _augment_conclusion(
                    question_data.get("conclusion", ""), rebuilt, len(details))
                # 未归类占比告警
                other_dim = next((d for d in rebuilt if d["name"] == OTHER_CANON), None)
                if other_dim and len(details) > 0:
                    ratio = other_dim["count"] / len(details)
                    if ratio > 0.2:
                        warnings.append(
                            f"[{question_col}] 未归类占比 {ratio * 100:.1f}%（>20%），"
                            "建议补充维度或重新逐条编码以提升准确度。")

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
    if keyword_fallback_used:
        result["keyword_fallback_used"] = True
    if warnings:
        result["warnings"] = warnings

    return result


# ========================================================================= #
#                        CLI 入口
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(description="问卷文本分析结果导出")
    parser.add_argument("--output_path", default=None,
                        help="输出 Excel 文件路径；不传时根据题目自动生成如 Q5_MC满意原因.xlsx")
    parser.add_argument("--results_file", default=None, help="分析结果 JSON 文件路径")
    parser.add_argument("--results_json", default=None, help="分析结果 JSON 字符串")
    parser.add_argument("--file_path", default=None,
                        help="原始数据文件路径（可选）。当 details 为空时，从此文件提取文本并自动标注")
    parser.add_argument("--sheet_name", default="0", help="工作表名或编号（默认 0）")
    parser.add_argument("--sample_n", type=int, default=300,
                        help="逐条明细行数限制（默认 300，0=全量标注并重新统计维度）")
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
        output_path = args.output_path or default_output_filename(results)
        result = export_text_report(results, output_path,
                                    file_path=args.file_path,
                                    sheet_name=sheet_name,
                                    sample_n=args.sample_n)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if result.get("status") == "error":
            sys.exit(1)
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
