#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
问卷分析工具 - HTML 满意度报告生成
====================================

一键读取量化数据 → 自动计算满意度/NPS/细分维度/交叉分析 → 输出单文件 HTML 报告。

用法:
    python html_report.py \
        --file_path "量化数据.csv" \
        --survey_name "《我的世界》3月版本调研" \
        --survey_id 90502 \
        --date_range "2026-03-20 ~ 2026-03-21" \
        --clean_desc "无清洗" \
        --cross_cols '["Q54.请问您的性别是？","Q56.请问您的职业是？"]' \
        --theme default \
        --output "报告.html"

主题: default (专业蓝灰) / dark (深色仪表盘) / minimal (简约素白)
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from collections import OrderedDict

import pandas as pd
import numpy as np

# 确保能导入同目录模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load_and_classify import classify_columns


# ========================================================================= #
#                           工具函数
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
    """加载 CSV 或 Excel 数据"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        enc = _detect_csv_encoding(file_path)
        return pd.read_csv(file_path, encoding=enc, low_memory=False)
    else:
        return pd.read_excel(file_path)


def _to_numeric_series(df, col):
    """安全转换列为数值"""
    return pd.to_numeric(df[col], errors='coerce').dropna()


def _short_label(col_name):
    """缩短列名用于显示"""
    # 取冒号后面的部分，或取问号后面的部分
    if ':' in col_name:
        return col_name.split(':')[-1].strip()
    return col_name


# ========================================================================= #
#                  自动识别满意度相关题目
# ========================================================================= #

def _identify_questions(df, classification):
    """自动识别满意度报告所需的核心题目"""
    result = {
        'overall_sat': None,     # 整体满意度列名
        'nps': None,             # NPS列名
        'dim_cols': [],          # 细分维度列名列表
        'dim_labels': [],        # 细分维度显示标签
        'reason_cols': [],       # 不满原因列名列表(多选子列)
        'single_choice': [],     # 其他单选题
        'multi_choice': {},      # 其他多选题
    }

    single_cols = classification.get('single_choice', [])
    multi_choice = classification.get('multi_choice', {})
    matrix_scale = classification.get('matrix_scale', {})

    # 1. 找整体满意度题
    for col in single_cols:
        if ('整体满意' in col or '总体满意' in col) and col.startswith('Q1'):
            result['overall_sat'] = col
            break
    if not result['overall_sat']:
        for col in single_cols:
            if '整体满意' in col:
                result['overall_sat'] = col
                break

    # 2. 找 NPS 题
    for col in single_cols:
        if ('推荐' in col or 'NPS' in col.upper()) and col.startswith('Q2'):
            result['nps'] = col
            break
    if not result['nps']:
        for col in single_cols:
            if '推荐' in col:
                s = _to_numeric_series(df, col)
                if s.max() >= 10:
                    result['nps'] = col
                    break

    # 3. 找细分维度 (矩阵量表)
    for prefix, cols in matrix_scale.items():
        for col in cols:
            if col == result['overall_sat'] or col == result['nps']:
                continue
            s = _to_numeric_series(df, col)
            if len(s) > 0 and s.max() <= 10:
                result['dim_cols'].append(col)
                result['dim_labels'].append(_short_label(col))

    # 也把独立的满意度单选题加入维度
    for col in single_cols:
        if col in [result['overall_sat'], result['nps']]:
            continue
        if '满意度' in col or '满意' in col:
            s = _to_numeric_series(df, col)
            if len(s) > 0 and 1 <= s.max() <= 5:
                result['dim_cols'].append(col)
                result['dim_labels'].append(_short_label(col))

    # 4. 找不满原因 (含"不满意"/"不太满意"的多选题)
    for prefix, sub_cols in multi_choice.items():
        if any('不太满意的主要原因' in c or '不满意的原因' in c for c in sub_cols):
            result['reason_cols'] = sub_cols
            break

    # 5. 其他单选/多选题（排除已识别的）
    identified = set()
    identified.add(result['overall_sat'])
    identified.add(result['nps'])
    identified.update(result['dim_cols'])

    for col in single_cols:
        if col not in identified and not col.startswith(('Q54', 'Q55', 'Q56', 'Q57', 'Q58', 'Q59')):
            # 跳过人口学和访谈邀约题
            result['single_choice'].append(col)

    for prefix, sub_cols in multi_choice.items():
        if sub_cols != result['reason_cols']:
            result['multi_choice'][prefix] = sub_cols

    return result


# ========================================================================= #
#                  指标计算
# ========================================================================= #

def _calc_overall(df, col):
    """计算整体满意度指标"""
    s = _to_numeric_series(df, col)
    if len(s) == 0:
        return {}
    return {
        'mean': round(float(s.mean()), 2),
        'top2': round(float((s >= 4).mean()) * 100, 1),
        'mid': round(float((s == 3).mean()) * 100, 1),
        'bot2': round(float((s <= 2).mean()) * 100, 1),
        'n': int(len(s)),
        'dist': {str(int(k)): int(v) for k, v in s.value_counts().sort_index().items()},
    }


def _calc_nps(df, col):
    """计算 NPS 指标"""
    s = _to_numeric_series(df, col)
    if len(s) == 0:
        return {}
    promoters = int((s >= 9).sum())
    passives = int(((s >= 7) & (s <= 8)).sum())
    detractors = int((s <= 6).sum())
    total = promoters + passives + detractors
    if total == 0:
        return {}
    return {
        'value': round((promoters / total - detractors / total) * 100, 1),
        'promoters': round(promoters / total * 100, 1),
        'passives': round(passives / total * 100, 1),
        'detractors': round(detractors / total * 100, 1),
        'n': total,
        'dist': {str(int(k)): int(v) for k, v in s.value_counts().sort_index().items()},
    }


def _calc_dimensions(df, dim_cols, dim_labels):
    """计算各维度满意度"""
    dims = []
    for col, label in zip(dim_cols, dim_labels):
        s = _to_numeric_series(df, col)
        if len(s) == 0:
            continue
        dims.append({
            'name': label,
            'mean': round(float(s.mean()), 2),
            'top2': round(float((s >= 4).mean()) * 100, 1),
            'bot2': round(float((s <= 2).mean()) * 100, 1),
            'n': int(len(s)),
        })
    return sorted(dims, key=lambda x: x['mean'], reverse=True)


def _calc_reasons(df, reason_cols, total_n):
    """计算不满原因统计"""
    stats = []
    for col in reason_cols:
        reason = col.split('？:')[-1] if '？:' in col else col.split('是？:')[-1] if '是？:' in col else _short_label(col)
        count = int(pd.to_numeric(df[col], errors='coerce').fillna(0).astype(bool).sum())
        stats.append({'reason': reason, 'count': count, 'pct': round(count / total_n * 100, 1)})
    return sorted(stats, key=lambda x: x['count'], reverse=True)


def _calc_question_stats(df, col, total_n):
    """计算单个单选/多选题的选项统计"""
    s = pd.to_numeric(df[col], errors='coerce').dropna()
    if len(s) == 0:
        return None
    vc = s.value_counts().sort_index()
    options = []
    for val, cnt in vc.items():
        options.append({'label': str(int(val)), 'count': int(cnt)})
    return options


def _calc_multi_question_stats(df, sub_cols, total_n):
    """计算多选题各子选项统计"""
    options = []
    for col in sub_cols:
        label = _short_label(col)
        count = int(pd.to_numeric(df[col], errors='coerce').fillna(0).astype(bool).sum())
        options.append({'label': label, 'count': count})
    return sorted(options, key=lambda x: x['count'], reverse=True)


# ========================================================================= #
#                  交叉分析计算
# ========================================================================= #

def _get_value_labels(df, col):
    """获取分组变量的值标签映射（从文本数据或编码推断）"""
    s = pd.to_numeric(df[col], errors='coerce').dropna()
    unique_vals = sorted(s.unique())

    # 常见映射
    if '性别' in col:
        label_map = {1: '男', 2: '女', 3: '其他/不愿透露'}
    elif '年龄' in col:
        label_map = {1: '6岁以下', 2: '7-9岁', 3: '10-12岁', 4: '13-15岁', 5: '16-18岁',
                     6: '19-22岁', 7: '23-25岁', 8: '26-30岁', 9: '31-35岁', 10: '36-40岁',
                     11: '18岁以下', 12: '41岁以上'}
    elif '职业' in col:
        label_map = {1: '在读小学生', 2: '在读初中生', 3: '在读高中/中职生',
                     4: '在读大学/大专生', 5: '在读硕博研究生',
                     6: 'IT/互联网', 7: '金融', 8: '教育', 9: '医疗', 10: '制造业',
                     11: '服务业', 12: '政府/事业单位', 13: '自由职业', 14: '待业',
                     15: '学生(未细分)', 16: '其他'}
    else:
        label_map = {int(v): str(int(v)) for v in unique_vals}

    return label_map


def _calc_cross_overall(df, sat_col, nps_col, group_col):
    """按分组变量交叉计算整体满意度和NPS"""
    label_map = _get_value_labels(df, group_col)
    group_s = pd.to_numeric(df[group_col], errors='coerce')
    results = []

    for val in sorted(group_s.dropna().unique()):
        val_int = int(val)
        label = label_map.get(val_int, str(val_int))
        mask = group_s == val
        n = int(mask.sum())
        if n < 30:
            continue

        row = {'group': label, 'n': n}

        # 满意度
        if sat_col:
            sat = _to_numeric_series(df.loc[mask], sat_col)
            if len(sat) > 0:
                row['sat_mean'] = round(float(sat.mean()), 2)
                row['sat_top2'] = round(float((sat >= 4).mean()) * 100, 1)
            else:
                row['sat_mean'] = 0
                row['sat_top2'] = 0
        else:
            row['sat_mean'] = 0
            row['sat_top2'] = 0

        # NPS
        if nps_col:
            nps_s = _to_numeric_series(df.loc[mask], nps_col)
            if len(nps_s) > 0:
                prom = float((nps_s >= 9).mean())
                detr = float((nps_s <= 6).mean())
                row['nps'] = round((prom - detr) * 100, 1)
                row['nps_promoters'] = round(prom * 100, 1)
                row['nps_detractors'] = round(detr * 100, 1)
            else:
                row['nps'] = 0
                row['nps_promoters'] = 0
                row['nps_detractors'] = 0
        else:
            row['nps'] = 0
            row['nps_promoters'] = 0
            row['nps_detractors'] = 0

        results.append(row)

    return results


def _calc_cross_dimensions(df, dim_cols, dim_labels, group_col):
    """按分组变量交叉计算各维度满意度均值"""
    label_map = _get_value_labels(df, group_col)
    group_s = pd.to_numeric(df[group_col], errors='coerce')

    # 确定有效分组（样本量 >= 30）
    valid_groups = []
    for val in sorted(group_s.dropna().unique()):
        val_int = int(val)
        if int(group_s.eq(val).sum()) >= 30:
            valid_groups.append((val_int, label_map.get(val_int, str(val_int))))

    group_names = [g[1] for g in valid_groups]
    dimensions = []

    for col, label in zip(dim_cols, dim_labels):
        values = []
        for val_int, _ in valid_groups:
            mask = group_s == val_int
            s = _to_numeric_series(df.loc[mask], col)
            values.append(round(float(s.mean()), 2) if len(s) > 0 else 0)
        dimensions.append({'name': label, 'values': values})

    return {'groups': group_names, 'dimensions': dimensions}


# ========================================================================= #
#                  预警检测
# ========================================================================= #

def _check_alerts(overall, nps_data, dimensions):
    """检查预警条件"""
    alerts = []

    # 整体满意度 < 3.5
    if overall.get('mean', 5) < 3.5:
        alerts.append({
            'level': 'red',
            'indicator': '整体满意度',
            'current': f"{overall['mean']}/5.0",
            'desc': '低于 3.5 健康线，需高度关注'
        })

    # NPS 检查
    nps_val = nps_data.get('value', 100)
    if nps_val < 0:
        alerts.append({
            'level': 'red',
            'indicator': 'NPS 净推荐值',
            'current': f"{nps_val}%",
            'desc': '贬损者多于推荐者，口碑风险'
        })
    elif nps_val < 30:
        alerts.append({
            'level': 'yellow',
            'indicator': 'NPS 净推荐值',
            'current': f"{nps_val}%",
            'desc': '有待改进（健康线为 30%+）'
        })

    # 维度检查
    for dim in dimensions:
        if dim['mean'] < 3.0:
            alerts.append({
                'level': 'red',
                'indicator': dim['name'],
                'current': f"{dim['mean']}/5.0",
                'desc': f"均值低于 3.0，不满率 {dim['bot2']}%，重点关注"
            })
        elif dim['bot2'] > 30:
            alerts.append({
                'level': 'yellow',
                'indicator': dim['name'],
                'current': f"不满率 {dim['bot2']}%",
                'desc': '超过 30% 玩家不满，需关注'
            })
        elif dim['bot2'] > 20:
            alerts.append({
                'level': 'yellow',
                'indicator': dim['name'],
                'current': f"{dim['mean']}/5.0 (不满率 {dim['bot2']}%)",
                'desc': '超过 20% 预警线'
            })

    return alerts


# ========================================================================= #
#                  HTML 渲染
# ========================================================================= #

def _render_html(report_data, theme='default'):
    """使用 Jinja2 渲染 HTML 报告"""
    try:
        from jinja2 import Environment, FileSystemLoader
    except ImportError:
        print(json.dumps({"error": "缺少 jinja2，请执行: pip install jinja2"}, ensure_ascii=False))
        sys.exit(1)

    template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
    echarts_path = os.path.join(template_dir, 'echarts.min.js')

    # 读取 ECharts JS
    with open(echarts_path, 'r', encoding='utf-8') as f:
        echarts_js = f.read()

    # 准备模板数据
    report_data_json = json.dumps(report_data, ensure_ascii=False)

    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template('satisfaction_report.html')

    html = template.render(
        theme=theme,
        meta=report_data['meta'],
        overall=report_data['overall'],
        nps=report_data['nps'],
        cross_overall=report_data.get('cross_overall'),
        dimensions=report_data['dimensions'],
        cross_dimensions=report_data.get('cross_dimensions'),
        reasons=report_data.get('reasons', []),
        questions=report_data.get('questions', []),
        alerts=report_data.get('alerts', []),
        echarts_js=echarts_js,
        report_data_json=report_data_json,
    )

    return html


# ========================================================================= #
#                  主流程
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
    """
    一键生成 HTML 满意度报告。

    Args:
        file_path: 量化数据文件路径 (CSV/Excel)
        survey_name: 问卷名称
        survey_id: 问卷 ID
        date_range: 数据时间范围
        clean_desc: 数据清洗说明
        cross_cols: 交叉分析分组列名列表 (JSON)
        theme: 主题风格 (default/dark/minimal)
        output_path: 输出 HTML 路径
    """

    # 1. 加载数据
    print(f"[html_report] Loading data: {file_path}", file=sys.stderr)
    df = _load_data(file_path)
    total_n = len(df)

    # 2. 分类
    print(f"[html_report] Classifying columns...", file=sys.stderr)
    classification = classify_columns(df)

    # 3. 识别题目
    questions_map = _identify_questions(df, classification)
    print(f"[html_report] Overall satisfaction: {questions_map['overall_sat']}", file=sys.stderr)
    print(f"[html_report] NPS: {questions_map['nps']}", file=sys.stderr)
    print(f"[html_report] Dimensions: {len(questions_map['dim_cols'])} items", file=sys.stderr)

    # 4. 计算指标
    print(f"[html_report] Calculating metrics...", file=sys.stderr)
    overall = _calc_overall(df, questions_map['overall_sat']) if questions_map['overall_sat'] else {}
    nps_data = _calc_nps(df, questions_map['nps']) if questions_map['nps'] else {}
    dimensions = _calc_dimensions(df, questions_map['dim_cols'], questions_map['dim_labels'])
    reasons = _calc_reasons(df, questions_map['reason_cols'], total_n) if questions_map['reason_cols'] else []

    # 5. 交叉分析
    cross_overall = OrderedDict()
    cross_dimensions = OrderedDict()
    if cross_cols:
        print(f"[html_report] Cross-tabulating by: {cross_cols}", file=sys.stderr)
        for group_col in cross_cols:
            if group_col not in df.columns:
                print(f"[html_report] WARNING: Column not found: {group_col}", file=sys.stderr)
                continue

            # 推断分组名称
            if '性别' in group_col:
                gname = '性别'
            elif '年龄' in group_col:
                gname = '年龄'
            elif '职业' in group_col:
                gname = '职业'
            else:
                gname = _short_label(group_col)

            cross_overall[gname] = _calc_cross_overall(
                df, questions_map['overall_sat'], questions_map['nps'], group_col
            )
            cross_dimensions[gname] = _calc_cross_dimensions(
                df, questions_map['dim_cols'], questions_map['dim_labels'], group_col
            )

    # 6. 单选/多选题统计
    print(f"[html_report] Computing question statistics...", file=sys.stderr)
    question_charts = []

    # 单选题 → 选项 <= 6 用饼图，>6 用柱状图
    for col in questions_map['single_choice'][:15]:  # 最多15题，避免报告太长
        options = _calc_question_stats(df, col, total_n)
        if options and len(options) >= 2:
            question_charts.append({
                'title': _short_label(col),
                'chart_type': 'pie' if len(options) <= 6 else 'bar',
                'options': options,
            })

    # 多选题 → 横向柱状图
    count = 0
    for prefix, sub_cols in questions_map['multi_choice'].items():
        if count >= 10:
            break
        options = _calc_multi_question_stats(df, sub_cols, total_n)
        if options and len(options) >= 2:
            # 用第一个子列的题干作为标题
            title = sub_cols[0].split('？:')[0].split('是？')[0] if '？:' in sub_cols[0] or '是？' in sub_cols[0] else prefix
            title = _short_label(title) if len(title) > 60 else title
            question_charts.append({
                'title': title,
                'chart_type': 'bar',
                'options': options,
            })
            count += 1

    # 7. 预警
    alerts = _check_alerts(overall, nps_data, dimensions)

    # 8. 组装数据
    report_data = {
        'meta': {
            'title': survey_name or '问卷',
            'survey_id': survey_id or '',
            'total_n': total_n,
            'date_range': date_range or '',
            'clean_desc': clean_desc,
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        },
        'overall': overall,
        'nps': nps_data,
        'dimensions': dimensions,
        'cross_overall': dict(cross_overall) if cross_overall else None,
        'cross_dimensions': dict(cross_dimensions) if cross_dimensions else None,
        'reasons': reasons,
        'questions': question_charts,
        'alerts': alerts,
    }

    # 9. 渲染 HTML
    print(f"[html_report] Rendering HTML (theme={theme})...", file=sys.stderr)
    html = _render_html(report_data, theme=theme)

    # 10. 输出
    if not output_path:
        base = os.path.splitext(file_path)[0]
        output_path = f"{base}_满意度报告.html"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    result = {
        'status': 'success',
        'output_path': os.path.abspath(output_path),
        'theme': theme,
        'total_n': total_n,
        'dimensions_count': len(dimensions),
        'cross_groups': list(cross_overall.keys()) if cross_overall else [],
        'alerts_count': len(alerts),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


# ========================================================================= #
#                  CLI
# ========================================================================= #

def main():
    parser = argparse.ArgumentParser(description="HTML 满意度报告生成")
    parser.add_argument("--file_path", required=True, help="量化数据文件路径")
    parser.add_argument("--survey_name", default="", help="问卷名称")
    parser.add_argument("--survey_id", default="", help="问卷 ID")
    parser.add_argument("--date_range", default="", help="数据时间范围")
    parser.add_argument("--clean_desc", default="无清洗", help="清洗逻辑描述")
    parser.add_argument("--cross_cols", default=None, help="交叉分析分组列名 JSON 列表")
    parser.add_argument("--theme", default="default", choices=["default", "dark", "minimal"],
                        help="主题风格")
    parser.add_argument("--output", default=None, help="输出 HTML 路径")

    args = parser.parse_args()

    cross_cols = None
    if args.cross_cols:
        cross_cols = json.loads(args.cross_cols)

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
