"""句式触发、材料证据识别及保守能力分流。"""
import math
import re
from pathlib import Path

import pandas as pd


def trigger(text):
    """命中Excel月度生成/更新意图，明确否定优先。"""
    text = text.lower()
    if re.search(r'(不要|不需要|不|勿)(生成|更新|做|制作|输出).{0,12}(excel|汇总)', text):
        return False
    if re.search(r'月度汇总\s*excel', text):
        return True
    return bool((re.search(r'更新|生成|制作|补上|补充|刷新|汇总', text)) and
                re.search(r'月度|\d{1,2}月|各月|上个月|本月|逐月', text) and
                re.search(r'excel|汇总表|满意度汇总|上个月汇总', text))


def classify_frame(frame):
    """根据行粒度和内容识别，不把.xlsx一概当模板。"""
    columns = [str(c) for c in frame.columns]
    cells = [str(v) for v in frame.head(120).to_numpy().ravel() if pd.notna(v)]
    has_question = any(re.match(r'^[QY]\d+[.．]', c) for c in columns)
    identifiers = any(c in columns for c in ['序号', '答案唯一标识', '开始答题时间', '结束答题时间'])
    month = any(re.search(r'\d{2,4}年-?\d{1,2}月|20\d{2}[-/]\d{1,2}', v) for v in cells + columns)
    aggregate = any(re.search(r'指标|均值|得分|样本量|满意度|NPS|占比', v, re.I) for v in cells + columns)
    if has_question and identifiers:
        return 'raw'
    if month and aggregate and not has_question:
        return 'summary'
    return 'unknown'


def inspect_material(path):
    """逐sheet检查材料；同文件允许同时有明细与汇总，记录证据。"""
    path = Path(path)
    if path.suffix.lower() == '.csv':
        frames = {'CSV': pd.read_csv(path, nrows=120, encoding='utf-8-sig')}
    elif path.suffix.lower() == '.xlsx':
        frames = pd.read_excel(path, sheet_name=None, nrows=120, header=None)
        # 原始XLSX的首行才是列名，汇总则保留全部单元格扫描。
        for name, frame in frames.items():
            if not frame.empty and any(re.match(r'^[QY]\d+[.．]',str(c)) for c in frame.iloc[0]):
                frame.columns = frame.iloc[0].tolist()
                frames[name] = frame.iloc[1:].copy()
    else:
        raise ValueError('材料只支持CSV/XLSX')
    sheets = [{'sheet': name, 'kind': classify_frame(frame), 'sampled_rows': len(frame)} for name, frame in frames.items()]
    kinds = sorted({s['kind'] for s in sheets if s['kind'] != 'unknown'})
    return {'path': str(path), 'kinds': kinds or ['unknown'], 'sheets': sheets}


def choose_flow(kinds, has_template):
    """明确材料对应流程，不拿汇总均值反推个人原声。"""
    kinds = set(kinds)
    if not kinds or 'unknown' in kinds or not kinds <= {'raw','summary'}:
        raise ValueError('材料身份不明确，必须确认后再计算')
    if 'raw' in kinds:
        return 'update_from_raw_and_summary' if has_template else 'new_from_raw'
    return 'update_from_summary' if has_template else 'assemble_summaries'


def validate_summary_record(record, require_pooling=False):
    """已计算数据仅继承已有精度；合并/加权必须有真实分母。"""
    result = dict(record)
    kind, value, n = result['kind'], result['value'], result.get('n')
    if kind not in ['mean','nps','proportion']:
        raise ValueError('汇总记录统计类型未知')
    if value is not None and (not isinstance(value,(int,float)) or not math.isfinite(value)):
        raise ValueError('汇总值必须为有效数值或null')
    if kind == 'nps' and value is not None and not -1 <= value <= 1:
        raise ValueError('NPS汇总值需明确转换为-1至1比例，不能猜测百分数单位')
    if kind == 'proportion' and value is not None and not 0 <= value <= 1:
        raise ValueError('占比需0至1比例')
    if n is not None and (not isinstance(n,int) or n < 0):
        raise ValueError('有效分母必须为非负整数')
    if require_pooling and (not n or value is None):
        raise ValueError('缺少真实有效分母，不能合并或加权')
    result['n'] = n
    return result
