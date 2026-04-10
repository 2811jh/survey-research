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

def _should_skip_question(col_name):
    """判断是否应该跳过该题不展示在'各题统计图表'中"""
    # 1. 非必填题
    if '非必填' in col_name or '（非必填）' in col_name:
        return True
    # 2. 满意度题/满意程度题
    if '满意度' in col_name or '满意程度' in col_name or '体验感受如何' in col_name:
        return True
    # 3. 含[图片]的意义不明题
    if '[图片]' in col_name:
        return True
    # 4. 人口学和访谈邀约
    q_match = re.match(r'^Q(\d+)', col_name)
    if q_match:
        qnum = int(q_match.group(1))
        if qnum >= 54:  # Q54性别、Q55年龄、Q56职业、Q57-Q59访谈
            return True
    return False


def _identify_questions(df, classification):
    """自动识别满意度报告所需的核心题目"""
    result = {
        'overall_sat': None,     # 整体满意度列名
        'nps': None,             # NPS列名
        'dim_cols': [],          # 细分维度列名列表
        'dim_labels': [],        # 细分维度显示标签
        'dim_groups': [],        # 细分维度按题目分组 [{group_title, dims}]
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

    # 4.5 按题目前缀将维度分组，使用专业化标题
    # 关键词 → 专业标题映射
    _title_map = {
        '体验感受': '体验感受满意度',
        '基础性能': '基础性能满意度',
        '性能问题': '性能问题频率',
        '美术画面': '美术画面满意度',
        '界面使用': '界面使用满意度',
        '模组': '模组体验满意度',
        '社交': '社交体验满意度',
        'UGC': 'UGC 玩法满意度',
        '活动': '活动满意度',
        '付费': '付费体验满意度',
        '皮肤': '皮肤满意度',
        '内容乐趣': '内容乐趣满意度',
        '玩法乐趣': '玩法乐趣满意度',
    }

    def _professional_title(raw_title):
        """将原始题干转为专业分组标题"""
        for keyword, title in _title_map.items():
            if keyword in raw_title:
                return title
        # 兜底: 简短化处理
        short = re.sub(r'您对.*?的', '', raw_title)
        short = re.sub(r'\？.*$', '', short)
        short = re.sub(r'\*.*$', '', short)
        short = short.strip()
        return short if len(short) <= 12 else short[:12]

    current_group = None
    current_prefix = None
    for col, label in zip(result['dim_cols'], result['dim_labels']):
        # 提取 Q 编号前缀 (如 Q3, Q4, Q21)
        m = re.match(r'^(Q\d+)', col)
        prefix = m.group(1) if m else 'other'
        # 提取题干
        if ':' in col:
            raw_title = col.split(':')[0].strip()
            raw_title = re.sub(r'^Q\d+\.', '', raw_title).strip()
        else:
            raw_title = label

        if prefix != current_prefix:
            current_prefix = prefix
            pro_title = _professional_title(raw_title)
            current_group = {'group_title': pro_title, 'dims': []}
            result['dim_groups'].append(current_group)
        current_group['dims'].append({'col': col, 'label': label})

    # 5. 其他单选/多选题（排除已识别的 + 过滤规则）
    identified = set()
    identified.add(result['overall_sat'])
    identified.add(result['nps'])
    identified.update(result['dim_cols'])

    for col in single_cols:
        if col not in identified and not _should_skip_question(col):
            result['single_choice'].append(col)

    for prefix, sub_cols in multi_choice.items():
        if sub_cols != result['reason_cols']:
            # 检查多选题的母题是否应该跳过
            root_col = sub_cols[0] if sub_cols else ''
            if not _should_skip_question(root_col):
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


def _calc_reasons(df, reason_cols, total_n, text_cols=None):
    """计算不满原因统计，并从文本题中搜索每条原因的玩家原话"""
    stats = []
    for col in reason_cols:
        reason = col.split('？:')[-1] if '？:' in col else col.split('是？:')[-1] if '是？:' in col else _short_label(col)
        count = int(pd.to_numeric(df[col], errors='coerce').fillna(0).astype(bool).sum())
        stats.append({'reason': reason, 'count': count, 'pct': round(count / total_n * 100, 1), 'quote': ''})
    stats = sorted(stats, key=lambda x: x['count'], reverse=True)

    # 从文本题中为每个原因搜索匹配的玩家原话
    if text_cols:
        # 合并所有文本列成一个大文本池
        all_texts = []
        for tc in text_cols:
            if tc in df.columns:
                texts = df[tc].dropna().astype(str).tolist()
                all_texts.extend([t for t in texts if len(t) >= 4 and t not in ('无', '没有', '不知道', '暂无', 'nan')])

        if all_texts:
            for item in stats:
                # 提取原因中的关键词用于匹配
                reason = item['reason']
                # 常见关键词映射
                keywords = []
                if '性能' in reason or '卡顿' in reason or '延迟' in reason or '闪退' in reason:
                    keywords = ['卡顿', '延迟', '闪退', '卡', '掉帧', 'lag', '卡死']
                elif '环境' in reason or '外挂' in reason or '破坏' in reason:
                    keywords = ['外挂', '炸', '破坏', '骂', '挂', '恶意']
                elif '社交' in reason or '氛围' in reason:
                    keywords = ['社交', '好友', '聊天', '组队', '匹配']
                elif '界面' in reason or '操作' in reason:
                    keywords = ['界面', '操作', 'UI', '按钮', '误触', '难找']
                elif '皮肤' in reason or '付费' in reason:
                    keywords = ['皮肤', '付费', '贵', '充值', '氪', '抽奖', '价格']
                elif 'UGC' in reason or '模组' in reason:
                    keywords = ['模组', 'mod', 'MOD', '地图', '资源中心']
                elif '基础操作' in reason or '行走' in reason or '跳跃' in reason:
                    keywords = ['操作', '行走', '跳跃', '手感', '不流畅']
                elif '信誉' in reason or '禁言' in reason or '举报' in reason:
                    keywords = ['禁言', '举报', '封号', '信誉']
                elif 'BUG' in reason or 'bug' in reason.lower():
                    keywords = ['bug', 'BUG', 'Bug', '漏洞', '异常']
                elif '新手' in reason or '教学' in reason:
                    keywords = ['新手', '教程', '引导', '教学']
                elif '福利' in reason or '奖励' in reason:
                    keywords = ['福利', '奖励', '白嫖', '送']
                elif '更新' in reason or '版本' in reason:
                    keywords = ['更新', '版本', '速度慢']
                elif '存档' in reason or '丢失' in reason:
                    keywords = ['存档', '丢', '丢失', '消失']
                elif '屏蔽' in reason or '聊天' in reason:
                    keywords = ['屏蔽', '聊天', '屏蔽词']
                else:
                    # 用原因文本中的关键字
                    keywords = [reason[:2]] if len(reason) >= 2 else [reason]

                # 在文本池中搜索匹配
                for text in all_texts:
                    if any(kw in text for kw in keywords):
                        # 找到匹配，截取合适长度
                        quote = text[:80] + ('...' if len(text) > 80 else '')
                        item['quote'] = f'"{quote}"'
                        break

    return stats


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
#                  自动结论生成
# ========================================================================= #

def _generate_macro_conclusion(overall, nps_data):
    """生成宏观大盘结论"""
    parts = []
    if overall:
        mean = overall.get('mean', 0)
        top2 = overall.get('top2', 0)
        if mean >= 4.0:
            parts.append(f"整体满意度 {mean} 分，玩家评价较好，满意率达 {top2}%。")
        elif mean >= 3.5:
            parts.append(f"整体满意度 {mean} 分，处于中等水平，满意率 {top2}%。")
        else:
            parts.append(f"整体满意度仅 {mean} 分，低于健康线，需重点关注。")
    if nps_data:
        nps = nps_data.get('value', 0)
        if nps > 30:
            parts.append(f"NPS 得分 {nps}%，保持在优秀区间。")
        elif nps >= 0:
            parts.append(f"NPS 得分 {nps}%，尚有提升空间。")
        else:
            parts.append(f"NPS 得分 {nps}%，贬损者多于推荐者，口碑承压。")
    return ' '.join(parts) if parts else ''


def _generate_reason_conclusion(reasons):
    """生成不满原因结论"""
    if not reasons:
        return ''
    top3 = reasons[:3]
    names = '、'.join([r['reason'] for r in top3])
    return f"玩家不满原因集中在{names}，占比分别为{'、'.join([str(r['pct'])+'%' for r in top3])}。"


def _generate_dim_conclusion(group_title, dims):
    """根据维度数据自动生成一句结论"""
    if not dims:
        return ''
    best = dims[0]  # 已排序，第一个是最高分
    worst = dims[-1]  # 最后一个是最低分
    avg = round(sum(d['mean'] for d in dims) / len(dims), 2)

    # 检查是否有预警项
    alerts = [d for d in dims if d['mean'] < 3.5 or d['bot2'] > 20]

    if len(dims) == 1:
        d = dims[0]
        if d['mean'] >= 4.0:
            return f"均值 {d['mean']} 分，满意率 {d['top2']}%，整体表现良好。"
        elif d['mean'] >= 3.5:
            return f"均值 {d['mean']} 分，表现中等，不满率 {d['bot2']}% 需留意。"
        else:
            return f"均值仅 {d['mean']} 分，不满率 {d['bot2']}%，需重点关注。"

    if not alerts:
        return f"整体均值 {avg} 分，其中「{best['name']}」得分最高（{best['mean']}），各项均在健康区间。"
    else:
        alert_names = '、'.join([f"「{d['name']}」" for d in alerts[:3]])
        return f"整体均值 {avg} 分，「{best['name']}」表现最佳（{best['mean']}），但 {alert_names} 需关注（不满率超 {alerts[0]['bot2']}%）。"


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
        macro_conclusion=report_data.get('macro_conclusion', ''),
        cross_overall=report_data.get('cross_overall'),
        dimensions=report_data['dimensions'],
        dim_groups=report_data.get('dim_groups', []),
        reasons=report_data.get('reasons', []),
        reason_conclusion=report_data.get('reason_conclusion', ''),
        gaming_genres=report_data.get('gaming_genres', []),
        gaming_titles=report_data.get('gaming_titles', []),
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

    # 计算按题目分组的维度数据 + 自动结论 + 展示类型
    dim_groups_data = []
    for group in questions_map['dim_groups']:
        group_dims = []
        for d in group['dims']:
            s = _to_numeric_series(df, d['col'])
            if len(s) == 0:
                continue
            group_dims.append({
                'name': d['label'],
                'mean': round(float(s.mean()), 2),
                'top2': round(float((s >= 4).mean()) * 100, 1),
                'bot2': round(float((s <= 2).mean()) * 100, 1),
                'n': int(len(s)),
            })
        if group_dims:
            sorted_dims = sorted(group_dims, key=lambda x: x['mean'], reverse=True)
            # 展示类型: 少于4项用 KPI 卡片，>=4项用表格+柱状图
            display_type = 'kpi' if len(sorted_dims) < 4 else 'chart'
            # 自动生成结论
            conclusion = _generate_dim_conclusion(group['group_title'], sorted_dims)
            dim_groups_data.append({
                'group_title': group['group_title'],
                'dims': sorted_dims,
                'display_type': display_type,
                'conclusion': conclusion,
            })

    # 不满原因（传入文本列用于搜索玩家原话）
    text_cols = classification.get('text', [])
    reasons = _calc_reasons(df, questions_map['reason_cols'], total_n, text_cols=text_cols) if questions_map['reason_cols'] else []

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

    # 6. 玩家画像 — 手游类型 & 具体手游
    print(f"[html_report] Extracting gaming profile...", file=sys.stderr)
    gaming_genres = []  # 手游类型 (TreeMap)
    gaming_titles = []  # 具体手游 (Top20 排行榜)

    for prefix, sub_cols in questions_map['multi_choice'].items():
        first_col = sub_cols[0] if sub_cols else ''
        # 手游类型 (Q42)
        if '常玩' in first_col and '手机游戏' in first_col and '类' in first_col:
            for col in sub_cols:
                label = col.split('？:')[-1] if '？:' in col else col.split(':')[-1]
                label = label.strip()
                if '没在玩' in label or '其他' in label:
                    continue
                cnt = int(pd.to_numeric(df[col], errors='coerce').fillna(0).astype(bool).sum())
                if cnt > 0:
                    gaming_genres.append({'name': label, 'value': cnt, 'pct': round(cnt / total_n * 100, 1)})
            gaming_genres = sorted(gaming_genres, key=lambda x: x['value'], reverse=True)
        # 具体手游 (Q43)
        elif '目前在玩' in first_col and '手机游戏' in first_col:
            for col in sub_cols:
                label = col.split(':')[-1].strip()
                if not label or '其他' in label:
                    continue
                cnt = int(pd.to_numeric(df[col], errors='coerce').fillna(0).astype(bool).sum())
                if cnt > 0:
                    gaming_titles.append({'name': label, 'value': cnt, 'pct': round(cnt / total_n * 100, 1)})
            gaming_titles = sorted(gaming_titles, key=lambda x: x['value'], reverse=True)[:20]

    # 7. 预警
    alerts = _check_alerts(overall, nps_data, dimensions)

    # 7.5 自动生成结论
    macro_conclusion = _generate_macro_conclusion(overall, nps_data)
    reason_conclusion = _generate_reason_conclusion(reasons)

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
        'macro_conclusion': macro_conclusion,
        'dimensions': dimensions,
        'dim_groups': dim_groups_data,
        'cross_overall': dict(cross_overall) if cross_overall else None,
        'reasons': reasons,
        'reason_conclusion': reason_conclusion,
        'gaming_genres': gaming_genres,
        'gaming_titles': gaming_titles,
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
