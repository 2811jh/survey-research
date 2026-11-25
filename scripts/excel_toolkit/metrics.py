"""问卷字段映射和逐指标计算，不隐式清洗或提前舍入。"""
import re
from pathlib import Path

import pandas as pd


def validate_months(months):
    """校验明确的年月，拒绝重复月份。"""
    if not months or len(months) != len(set(months)):
        raise ValueError('月份为空或重复')
    for month in months:
        if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])', month):
            raise ValueError(f'月份必须为YYYY-MM: {month}')


def matches(column, spec):
    """匹配所有正向条件和排除词，不从题号猜测题意。"""
    if not spec or not any(k in spec for k in ['exact', 'contains', 'regex', 'suffix']):
        raise ValueError('字段匹配必须指定exact/contains/regex/suffix')
    return (('exact' not in spec or column == spec['exact'])
            and all(term in column for term in spec.get('contains', []))
            and not any(term in column for term in spec.get('exclude', []))
            and ('suffix' not in spec or column.endswith(spec['suffix']))
            and ('regex' not in spec or re.search(spec['regex'], column) is not None))


def resolve_columns(columns, config):
    """将配置映射到唯一列，缺失可选题显式为空。"""
    if len(columns) != len(set(columns)):
        raise ValueError('原始数据含重复列名')
    result = {}
    for item in config.get('metrics', []) + config.get('groups', []):
        candidates = [str(c) for c in columns if matches(str(c), item['match'])]
        if len(candidates) > 1:
            raise ValueError(f'字段映射歧义 {item["id"]}: {candidates}')
        if not candidates and not item.get('optional', False):
            raise ValueError(f'缺失必需字段: {item["id"]}')
        result[item['id']] = candidates[0] if candidates else None
    return result


def numeric(series, allowed, label):
    """仅允许配置列出的编码，非数值非空内容阻断计算。"""
    values = pd.to_numeric(series, errors='coerce')
    invalid = series.notna() & (values.isna() | ~values.isin(allowed))
    if invalid.any():
        raise ValueError(f'未知编码/非数值 {label}，异常行数={int(invalid.sum())}')
    return values


def load_frame(path, sheet=0):
    """读取原始CSV/XLSX，不尝试修复未知编码。"""
    path = Path(path)
    if path.suffix.lower() == '.csv':
        return pd.read_csv(path, encoding='utf-8-sig', low_memory=False)
    if path.suffix.lower() == '.xlsx':
        return pd.read_excel(path, sheet_name=sheet)
    raise ValueError('仅支持原始CSV或XLSX')


def validate_config(config):
    """验证身份、显示名及统计类型，防止同名覆盖和错误口径。"""
    metrics = config.get('metrics', [])
    groups = config.get('groups', [])
    if not metrics:
        raise ValueError('必须配置至少一个指标')
    ids = [x['id'] for x in metrics + groups]
    labels = [x['label'] for x in metrics]
    glabels = [x['label'] for x in groups] + ['全体']
    if len(ids) != len(set(ids)) or len(labels) != len(set(labels)) or len(glabels) != len(set(glabels)):
        raise ValueError('指标/人群身份或显示名重复')
    for item in metrics:
        if item['kind'] not in ['mean', 'nps', 'proportion']:
            raise ValueError(f'未知统计类型: {item["kind"]}')
        if item.get('direction', 'neutral') not in ['higher', 'lower', 'neutral']:
            raise ValueError('未知指标方向')
        if item['kind'] == 'nps' and set(item['allowed']) != set(range(11)):
            raise ValueError('NPS必须明确0–10编码')
        if item['kind'] == 'proportion':
            if set(item['selected']) & set(item['unselected']):
                raise ValueError('选中和未选中编码重叠')
            if item['denominator']['mode'] not in ['all', 'answered', 'eligible']:
                raise ValueError('未知多选分母口径')


def compute_month(frame, config):
    """返回全体和分组统计、映射、逐指标样本分母审计。"""
    validate_config(config)
    data = frame.copy()
    input_rows = len(data)
    if 'response_id' in config:
        spec = config['response_id']
        valid = data[spec['column']].astype('string').str.fullmatch(spec['regex'], na=False)
        data = data.loc[valid].copy()
    mapping = resolve_columns(list(data.columns), config)
    masks = {'全体': pd.Series(True, index=data.index)}
    for group in config.get('groups', []):
        col = mapping[group['id']]
        values = numeric(data[col], group['allowed'], group['id'])
        masks[group['label']] = values.isin(group['values'])
    result = {'rows': len(data), 'input_rows': input_rows, 'excluded_rows': input_rows - len(data),
              'mapping': mapping, 'metrics': {}}
    for item in config['metrics']:
        col = mapping[item['id']]
        allowed = item['allowed'] if item['kind'] != 'proportion' else item['selected'] + item['unselected']
        values = numeric(data[col], allowed, item['id']) if col else pd.Series(float('nan'), index=data.index)
        eligible = values.notna()
        if col and item['kind'] == 'proportion':
            denom = item['denominator']
            if denom['mode'] == 'all':
                eligible = pd.Series(True, index=data.index)
            elif denom['mode'] == 'eligible':
                cols = [c for c in data.columns if matches(c, denom['match'])]
                if len(cols) != 1:
                    raise ValueError(f'资格分母列缺失/歧义: {item["id"]}')
                eligible = data[cols[0]].isin(denom['values'])
                if (values.isin(item['selected']) & ~eligible).any():
                    raise ValueError(f'资格外存在选中值: {item["id"]}')
            else:
                cols = [c for c in data.columns if matches(c, denom['match'])]
                if not cols:
                    raise ValueError('多选分母未匹配到选项列')
                observed = pd.concat([numeric(data[c], allowed, c).notna() for c in cols], axis=1)
                eligible = observed.any(axis=1)
        result['metrics'][item['id']] = {}
        for label, mask in masks.items():
            valid = mask & eligible
            n = int(valid.sum())
            sample = values.loc[valid]
            if not n:
                value = None
            elif item['kind'] == 'mean':
                value = float(sample.mean())
            elif item['kind'] == 'nps':
                value = (int((sample >= 9).sum()) - int((sample <= 6).sum())) / n
            else:
                value = int(sample.isin(item['selected']).sum()) / n
            result['metrics'][item['id']][label] = {
                'value': value, 'n': n, 'group_rows': int(mask.sum()),
                'missing': int((mask & values.isna()).sum()),
                'denominator': item.get('denominator', {'mode': 'valid_answers'})['mode']}
    return result
