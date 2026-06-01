#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - HTML 问卷结果报告生成
====================================

遍历全部题目 → 每题生成柱状图 + 自动结论 → 输出单文件 HTML 报告。

用法:
    python html_report.py \
        --file_path "量化数据.csv" \
        --survey_name "《我的世界》联机大厅调研" \
        --survey_id 93650 \
        --date_range "2026-05-01 ~ 2026-05-31" \
        --clean_desc "无清洗" \
        --cross_cols '["Q54.请问您的性别是？","Q56.请问您的职业是？"]' \
        --output "报告.html"
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime

import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_and_classify import classify_columns


# ========================================================================= #
#  工具
# ========================================================================= #

def _detect_csv_encoding(filepath, sample_size=8192):
    with open(filepath, 'rb') as f:
        raw = f.read(sample_size)
    if raw.startswith(b'\xef\xbb\xbf'):
        return 'utf-8-sig'
    try:
        raw.decode('utf-8')
        return 'utf-8'
    except UnicodeDecodeError:
        return 'gbk'


def _load_data(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        enc = _detect_csv_encoding(file_path)
        return pd.read_csv(file_path, encoding=enc, low_memory=False)
    return pd.read_excel(file_path)


def _qnum(col):
    """提取 Q 编号，用于排序"""
    m = re.match(r'^Q(\d+)', col)
    return int(m.group(1)) if m else 9999


def _col_title(col):
    """从列名提取纯题干"""
    raw = col.split(':')[0].strip()
    return re.sub(r'^Q\d+[.\s]', '', raw).strip()


def _short_label(col):
    """提取选项文字（冒号后）"""
    if ':' in col:
        return col.split(':')[-1].strip()
    return col.strip()


# ========================================================================= #
#  已知标签映射
# ========================================================================= #

_KNOWN_LABELS = {
    '性别':       {1: '男', 2: '女', 3: '其他/不愿透露'},
    '年龄':       {1: '6岁以下', 2: '7-9岁', 3: '10-12岁', 4: '13-15岁',
                   5: '16-18岁', 6: '19-22岁', 7: '23-25岁', 8: '26-30岁',
                   9: '31-35岁', 10: '36-40岁', 11: '18岁以下', 12: '41岁以上'},
    '职业':       {1: '在读小学生', 2: '在读初中生', 3: '在读高中/中职生',
                   4: '在读大学/大专生', 5: '在读硕博研究生',
                   6: 'IT/互联网', 7: '金融', 8: '教育', 9: '医疗', 10: '制造业',
                   11: '服务业', 12: '政府/事业单位', 13: '自由职业', 14: '待业',
                   15: '学生(未细分)', 16: '其他'},
    '频率':       {1: '几乎每天', 2: '每周4-6天', 3: '每周2-3天', 4: '每周1天',
                   5: '每月2-3次', 6: '每月1次', 7: '几乎不玩'},
    '时长':       {1: '30分钟以内', 2: '30分钟-1小时', 3: '1-2小时',
                   4: '2-3小时', 5: '3小时以上'},
    '意愿':       {1: '非常强烈', 2: '比较强烈', 3: '一般', 4: '不太强烈', 5: '非常不强烈'},
    '交流':       {1: '非常强烈', 2: '比较强烈', 3: '一般', 4: '不太强烈', 5: '非常不强烈'},
    '有几个人':   {1: '1人（自己一人）', 2: '2人', 3: '3-5人', 4: '6-10人', 5: '10人以上'},
    '小团体':     {1: '有，3-5人', 2: '有，6-10人', 3: '有，10人以上',
                   4: '没有，喜欢随机匹配', 5: '没有，主要自己玩'},
    '小群组':     {1: '有，3-5人', 2: '有，6-10人', 3: '有，10人以上',
                   4: '没有，喜欢随机匹配', 5: '没有，主要自己玩'},
    '付费内购':   {1: '有，且付费过', 2: '有，但没付费过', 3: '没有，不愿意付费'},
}


def _get_labels(col):
    for kw, lmap in _KNOWN_LABELS.items():
        if kw in col:
            return lmap
    return {}


# ========================================================================= #
#  题目统计
# ========================================================================= #

def _is_demo_col(col):
    return any(kw in col for kw in ['性别', '年龄', '职业'])


def _single_stats(df, col, total_n):
    """单选题：返回各选项 count/pct，并携带标签"""
    s = pd.to_numeric(df[col], errors='coerce').dropna()
    if len(s) == 0:
        return None
    lmap = _get_labels(col)
    vc = s.value_counts().sort_index()
    options = []
    for val, cnt in vc.items():
        options.append({
            'label': lmap.get(int(val), f'选项{int(val)}'),
            'count': int(cnt),
            'pct': round(int(cnt) / len(s) * 100, 1),
        })
    return {'n': int(len(s)), 'options': options}


def _multi_stats(df, sub_cols, total_n):
    """多选题：各子列 count/pct（分母 = 作答人数）"""
    # 计算作答人数（至少选1项的行数）
    mat = df[sub_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    answered = int((mat.sum(axis=1) > 0).sum())
    denom = answered if answered > 0 else total_n
    options = []
    for col in sub_cols:
        label = _short_label(col)
        cnt = int(mat[col].astype(bool).sum())
        options.append({
            'label': label,
            'count': cnt,
            'pct': round(cnt / denom * 100, 1),
        })
    options.sort(key=lambda x: x['count'], reverse=True)
    return {'n': denom, 'options': options}


def _matrix_stats(df, sub_cols, total_n):
    """矩阵题：每行子题给出均值和分布"""
    rows = []
    for col in sub_cols:
        s = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(s) == 0:
            continue
        vc = s.value_counts().sort_index()
        # 归一化分布为 pct
        dist = []
        for val, cnt in vc.items():
            dist.append({
                'label': str(int(val)),
                'count': int(cnt),
                'pct': round(int(cnt) / len(s) * 100, 1),
            })
        rows.append({
            'label': _short_label(col),
            'mean': round(float(s.mean()), 2),
            'n': int(len(s)),
            'dist': dist,
        })
    rows.sort(key=lambda x: x['mean'], reverse=True)
    return {'rows': rows}


# ========================================================================= #
#  自动结论
# ========================================================================= #

def _conclude_single(title, options, n):
    """为单选题自动生成一句结论"""
    if not options:
        return ''
    top = options[0] if options[0]['pct'] >= options[-1]['pct'] else sorted(options, key=lambda x: x['pct'], reverse=True)[0]
    sorted_opts = sorted(options, key=lambda x: x['pct'], reverse=True)
    top1 = sorted_opts[0]
    top2_pct = sorted_opts[0]['pct'] + sorted_opts[1]['pct'] if len(sorted_opts) >= 2 else sorted_opts[0]['pct']
    if top1['pct'] >= 50:
        return f"**{top1['label']}**占比最高（{top1['pct']}%），过半玩家选择此项。"
    else:
        tops = '、'.join([f"「{o['label']}」{o['pct']}%" for o in sorted_opts[:2]])
        return f"分布较分散，TOP2 为 {tops}，合计 {round(top2_pct, 1)}%。"


def _conclude_multi(title, options, n):
    """为多选题自动生成一句结论"""
    if not options:
        return ''
    top3 = options[:3]
    tops = '、'.join([f"「{o['label']}」({o['pct']}%)" for o in top3])
    return f"玩家选择最多的前三项为 {tops}。"


def _conclude_matrix(title, rows):
    """为矩阵题自动生成一句结论"""
    if not rows:
        return ''
    best = rows[0]
    worst = rows[-1]
    avg = round(sum(r['mean'] for r in rows) / len(rows), 2)
    if len(rows) == 1:
        return f"均值 {best['mean']} 分（N={best['n']}）。"
    diff = round(best['mean'] - worst['mean'], 2)
    if diff >= 0.3:
        return f"各项均值平均 {avg} 分，「{best['label']}」得分最高（{best['mean']}），「{worst['label']}」相对最低（{worst['mean']}），差距 {diff} 分。"
    return f"各项均值平均 {avg} 分，得分分布较为集中，最高 {best['mean']}（{best['label']}）、最低 {worst['mean']}（{worst['label']}）。"


# ========================================================================= #
#  人口学交叉分析
# ========================================================================= #

def _calc_cross_question(df, q_col, q_type, group_col, group_label_map, sub_cols=None):
    """
    对单题按分组变量做交叉统计。
    q_type: 'single' | 'multi'
    返回 {'groups': [...], 'rows': [{'label': ..., 'values': [...pct...]}]}
    """
    group_s = pd.to_numeric(df[group_col], errors='coerce')
    valid_groups = []
    for val in sorted(group_s.dropna().unique()):
        n = int((group_s == val).sum())
        if n >= 30:
            valid_groups.append((int(val), group_label_map.get(int(val), str(int(val))), n))

    group_names = [f"{g[1]}(n={g[2]})" for g in valid_groups]
    rows = []

    if q_type == 'single':
        s = pd.to_numeric(df[q_col], errors='coerce')
        lmap = _get_labels(q_col)
        # 所有选项
        all_vals = sorted(s.dropna().unique())
        for val in all_vals:
            label = lmap.get(int(val), f'选项{int(val)}')
            values = []
            for gval, glabel, gn in valid_groups:
                mask = group_s == gval
                sub = s[mask].dropna()
                pct = round(float((sub == val).mean()) * 100, 1) if len(sub) > 0 else 0
                values.append(pct)
            rows.append({'label': label, 'values': values})

    elif q_type == 'multi' and sub_cols:
        mat = df[sub_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
        for col in sub_cols:
            label = _short_label(col)
            values = []
            for gval, glabel, gn in valid_groups:
                mask = (group_s == gval)
                sub_mat = mat.loc[mask]
                answered = int((sub_mat.sum(axis=1) > 0).sum())
                denom = answered if answered > 0 else int(mask.sum())
                cnt = int(sub_mat[col].astype(bool).sum())
                pct = round(cnt / denom * 100, 1) if denom > 0 else 0
                values.append(pct)
            rows.append({'label': label, 'values': values})

    return {'groups': group_names, 'rows': rows}


# ========================================================================= #
#  主流程
# ========================================================================= #

def _build_questions(df, classification, total_n, cross_cols=None):
    """
    遍历所有题目，构建报告用的题目列表。
    每题格式:
    {
      qid, type, title, n,
      stats: {...},        # 频率统计
      conclusion: str,     # 自动结论
      cross: {...} | None  # 交叉分析（若指定 cross_cols）
    }
    """
    questions = []
    processed_prefixes = set()

    # 准备交叉分析 group_col 标签映射
    cross_maps = {}
    if cross_cols:
        for gcol in cross_cols:
            if gcol in df.columns:
                cross_maps[gcol] = _get_labels(gcol)

    # 汇总所有多选前缀和矩阵前缀
    multi_prefixes = set(classification.get('multi_choice', {}).keys())
    matrix_prefixes = set(classification.get('matrix_scale', {}).keys())

    # 按 Q 编号排序所有题目
    all_single = sorted(classification.get('single_choice', []), key=_qnum)
    all_multi = {k: v for k, v in sorted(
        classification.get('multi_choice', {}).items(), key=lambda x: _qnum(x[0])
    )}
    all_matrix = {k: v for k, v in sorted(
        classification.get('matrix_scale', {}).items(), key=lambda x: _qnum(x[0])
    )}

    # 已出现的 Q 编号（避免重复）
    shown_qnums = set()

    def _make_cross(qid, q_type, q_col=None, sub_cols=None):
        if not cross_maps:
            return None
        result = {}
        for gcol, gmap in cross_maps.items():
            if gcol not in df.columns:
                continue
            gname = _short_label(gcol) if ':' not in gcol else gcol.split(':')[-1].strip()
            for kw in ['性别', '年龄', '职业']:
                if kw in gcol:
                    gname = kw
                    break
            result[gname] = _calc_cross_question(df, q_col, q_type, gcol, gmap, sub_cols=sub_cols)
        return result if result else None

    # 遍历顺序：按 Q 编号混合排列
    # 先收集所有题目及其排序键
    all_items = []  # (sort_key, type, qid, data)

    for col in all_single:
        qn = _qnum(col)
        all_items.append((qn, 'single', col, None))

    for prefix, sub_cols in all_multi.items():
        qn = _qnum(prefix)
        all_items.append((qn, 'multi', prefix, sub_cols))

    for prefix, sub_cols in all_matrix.items():
        qn = _qnum(prefix)
        all_items.append((qn, 'matrix', prefix, sub_cols))

    all_items.sort(key=lambda x: x[0])

    seen_qnums = set()

    for sort_key, qtype, col_or_prefix, sub_cols in all_items:
        qn = sort_key
        if qn in seen_qnums:
            continue
        seen_qnums.add(qn)

        qid = f"Q{qn}"

        if qtype == 'single':
            col = col_or_prefix
            # 跳过[图片]列
            if '[图片]' in col:
                continue
            title = _col_title(col)
            stats = _single_stats(df, col, total_n)
            if stats is None:
                continue
            conclusion = _conclude_single(title, stats['options'], stats['n'])
            cross = _make_cross(qid, 'single', q_col=col)
            questions.append({
                'qid': qid, 'type': 'single', 'title': title,
                'n': stats['n'], 'stats': stats, 'conclusion': conclusion, 'cross': cross,
            })

        elif qtype == 'multi':
            sub_cols = sub_cols or []
            if not sub_cols:
                continue
            root = sub_cols[0]
            if '[图片]' in root:
                continue
            title = _col_title(root)
            stats = _multi_stats(df, sub_cols, total_n)
            conclusion = _conclude_multi(title, stats['options'], stats['n'])
            cross = _make_cross(qid, 'multi', q_col=None, sub_cols=sub_cols)
            questions.append({
                'qid': qid, 'type': 'multi', 'title': title,
                'n': stats['n'], 'stats': stats, 'conclusion': conclusion, 'cross': cross,
            })

        elif qtype == 'matrix':
            sub_cols = sub_cols or []
            if not sub_cols:
                continue
            root = sub_cols[0]
            if '[图片]' in root:
                continue
            title = _col_title(root)
            stats = _matrix_stats(df, sub_cols, total_n)
            if not stats['rows']:
                continue
            conclusion = _conclude_matrix(title, stats['rows'])
            # 矩阵题暂不做交叉（数据量大）
            questions.append({
                'qid': qid, 'type': 'matrix', 'title': title,
                'n': stats['rows'][0]['n'] if stats['rows'] else total_n,
                'stats': stats, 'conclusion': conclusion, 'cross': None,
            })

    return questions


# ========================================================================= #
#  HTML 渲染
# ========================================================================= #

def _render_html(report_data, theme='default'):
    try:
        from jinja2 import Environment, FileSystemLoader
    except ImportError:
        print(json.dumps({'error': '缺少 jinja2，请执行: pip install jinja2'}, ensure_ascii=False))
        sys.exit(1)

    template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
    echarts_path = os.path.join(template_dir, 'echarts.min.js')

    with open(echarts_path, 'r', encoding='utf-8') as f:
        echarts_js = f.read()

    report_data_json = json.dumps(report_data, ensure_ascii=False)

    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template('satisfaction_report.html')

    html = template.render(
        theme=theme,
        meta=report_data['meta'],
        questions=report_data['questions'],
        echarts_js=echarts_js,
        report_data_json=report_data_json,
    )
    return html


# ========================================================================= #
#  入口
# ========================================================================= #

def generate_report(
    file_path,
    survey_name='',
    survey_id='',
    date_range='',
    clean_desc='无清洗',
    cross_cols=None,
    theme='default',
    output_path=None,
):
    print(f"[html_report] Loading: {file_path}", file=sys.stderr)
    df = _load_data(file_path)
    total_n = len(df)

    print(f"[html_report] Classifying columns...", file=sys.stderr)
    classification = classify_columns(df)

    print(f"[html_report] Building questions...", file=sys.stderr)
    questions = _build_questions(df, classification, total_n, cross_cols=cross_cols)
    print(f"[html_report] Total questions: {len(questions)}", file=sys.stderr)

    report_data = {
        'meta': {
            'title': survey_name or '问卷调研报告',
            'survey_id': survey_id or '',
            'total_n': total_n,
            'date_range': date_range or '',
            'clean_desc': clean_desc,
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'questions_count': len(questions),
        },
        'questions': questions,
    }

    print(f"[html_report] Rendering HTML...", file=sys.stderr)
    html = _render_html(report_data, theme=theme)

    if not output_path:
        base = os.path.splitext(file_path)[0]
        output_path = f"{base}_调研报告.html"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    result = {
        'status': 'success',
        'output_path': os.path.abspath(output_path),
        'total_n': total_n,
        'questions_count': len(questions),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def main():
    parser = argparse.ArgumentParser(description="HTML 问卷结果报告生成")
    parser.add_argument("--file_path", required=True)
    parser.add_argument("--survey_name", default="")
    parser.add_argument("--survey_id", default="")
    parser.add_argument("--date_range", default="")
    parser.add_argument("--clean_desc", default="无清洗")
    parser.add_argument("--cross_cols", default=None, help="交叉分析分组列名 JSON 列表")
    parser.add_argument("--theme", default="default", choices=["default", "dark", "minimal"])
    parser.add_argument("--output", default=None)

    args = parser.parse_args()
    cross_cols = json.loads(args.cross_cols) if args.cross_cols else None

    generate_report(
        file_path=args.file_path,
        survey_name=args.survey_name,
        survey_id=args.survey_id,
        date_range=args.date_range,
        clean_desc=args.clean_desc,
        cross_cols=cross_cols,
        theme=args.theme,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
