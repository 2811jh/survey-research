"""事务交付、全量写入回读与独立verify入口。"""
from datetime import datetime
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import stat
import tempfile
import warnings
import xml.etree.ElementTree as ET
import zipfile

import openpyxl
import pandas as pd

from .metrics import compute_month, load_frame, validate_months
from .template import apply_template, plan_template, read_parts, sync_chart_caches
from .workbook import render_new


def digest(path):
    """返回输入/产物哈希，不记录原始答卷内容。"""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def ensure_unlocked(path):
    """检测Excel/WPS常见锁及只读，不尝试解除锁定。"""
    path = Path(path)
    for prefix in ['~$' + path.name, '~$' + path.name[1:], '.~lock.' + path.name + '#']:
        if path.with_name(prefix).exists():
            raise PermissionError(f'文件占用，先手动关闭: {path}')
    if path.exists() and not path.stat().st_mode & stat.S_IWRITE:
        raise PermissionError(f'只读文件: {path}')


def calculate(job):
    """校验输入年月并返回逐月完整计算及可预览映射。"""
    raw_inputs = job.get('inputs', [])
    if raw_inputs:
        validate_months([x['month'] for x in raw_inputs])
    results, sources = {}, []
    for entry in raw_inputs:
        path = Path(entry['path']).resolve()
        result = compute_month(load_frame(path, entry.get('sheet', 0)), job['config'])
        result['cleaning_note'] = entry.get('cleaning_note', '未声明清洗口径，月间比较需谨慎')
        results[entry['month']] = result
        sources.append({'month': entry['month'], 'path': str(path), 'sha256': digest(path)})
    from .routing import validate_summary_record
    records = list(job.get('summary_records', []))
    for source in job.get('summary_inputs', []):
        path = Path(source['path']).resolve()
        frame = load_frame(path, source.get('sheet', 0))
        columns = source['columns']
        required = ['month','metric','value','kind']
        if not all(k in columns for k in required):
            raise ValueError('汇总输入需明确month/metric/value/kind列映射')
        for _, row in frame.iterrows():
            record = {key: (None if pd.isna(row[col]) else row[col]) for key,col in columns.items()}
            if record.get('n') is not None:
                if float(record['n']) != int(record['n']):raise ValueError('汇总分母不是整数')
                record['n'] = int(record['n'])
            records.append(record)
        sources.append({'type':'summary','path':str(path),'sha256':digest(path)})
    metric_config = {x['id']:x for x in job['config']['metrics']}
    seen = set()
    for record in records:
        record = validate_summary_record(record)
        month, mid = record['month'], record['metric']
        validate_months([month])
        group = record.get('group','全体')
        key = month,mid,group
        if key in seen:raise ValueError('同月同指标同人群汇总重复，禁止平均均值')
        seen.add(key)
        if mid not in metric_config or record['kind'] != metric_config[mid]['kind']:
            raise ValueError('汇总统计类型/指标未匹配配置')
        if month in results and mid in results[month]['metrics'] and group in results[month]['metrics'][mid]:
            raise ValueError('原始答卷与已有汇总存在同月同指标冲突，需明确选用来源')
        result = results.setdefault(month, {'rows':None,'input_rows':None,'excluded_rows':0,'mapping':{},'metrics':{},'source_type':'summary'})
        result['mapping'][mid] = '已有汇总继承（不恢复缺失精度/个人分布）'
        result['metrics'].setdefault(mid,{})[group] = {'value':record['value'],'n':record.get('n'),
            'group_rows':None,'missing':None,'denominator':'inherited_summary'}
    validate_months(list(results))
    for month,result in results.items():
        for item in job['config']['metrics']:
            result['mapping'].setdefault(item['id'],'材料未提供，不推测')
            for group in ['全体']+[g['label'] for g in job['config'].get('groups',[])]:
                result['metrics'].setdefault(item['id'],{}).setdefault(group,{
                    'value':None,'n':None,'group_rows':None,'missing':None,'denominator':'not_available'})
    return results, sources


def verify(path, manifest=None):
    """全量比较预期单元格与显示格式，独立模式同时检查产物哈希。"""
    path = Path(path)
    if manifest is None:
        audit = json.loads(Path(str(path) + '.audit.json').read_text(encoding='utf-8'))
        if digest(path) != audit['sha256']:
            raise ValueError('Excel文件在交付后被改动，哈希不一致')
        manifest = audit['manifest']
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None:
            raise ValueError('Excel ZIP完整性失败')
        for n in z.namelist():
            if n.endswith('.xml'):
                ET.fromstring(z.read(n))
    with warnings.catch_warnings():
        warnings.simplefilter('ignore', UserWarning)
        wb = openpyxl.load_workbook(path, data_only=False)
    for expected in manifest:
        c = wb[expected['sheet']][expected['cell']]
        value, actual = expected['value'], c.value
        if value is None:
            valid = actual is None
        elif isinstance(value, (int, float)):
            valid = isinstance(actual, (int, float)) and math.isclose(value, actual, rel_tol=2e-15, abs_tol=2e-15)
        else:
            valid = value == actual and c.data_type != 'f'
        if not valid or c.number_format != expected['format']:
            raise ValueError(f'回读不符: {c.parent.title}!{c.coordinate}')
    return {'status': 'verified', 'checked_cells': len(manifest), 'sheets': wb.sheetnames}


def build(job, overwrite=False):
    """先验证临时产物再原子交付；失败不改用户Excel。"""
    output = Path(job['output']).resolve()
    if output.suffix.lower() != '.xlsx':
        raise ValueError('输出必须为.xlsx')
    if not output.parent.is_dir():
        raise ValueError('输出目录不存在，请先确认并创建目录')
    ensure_unlocked(output)
    if output.exists() and not overwrite:
        raise FileExistsError('默认另存；覆盖需明确--overwrite')
    sources = {Path(x['path']).resolve() for x in job.get('inputs', []) + job.get('summary_inputs', [])}
    if output in sources:
        raise ValueError('禁止覆盖原始答卷')
    template = Path(job['template']).resolve() if job.get('template') else None
    if template:
        ensure_unlocked(template)
    results, hashes = calculate(job)
    plan = plan_template(template, job.get('bindings', []), results, job['config'], job.get('append_only', False)) if template else []
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    with tempfile.TemporaryDirectory(prefix='.survey_excel_', dir=output.parent) as temp:
        generated = Path(temp) / 'generated.xlsx'
        include_detail = job.get('include_detail_sheets', not template)
        if include_detail:
            manifest = render_new(results, job['config'], generated)
        else:
            generated = None
            manifest = []
        candidate = Path(temp) / 'verified.xlsx'
        preservation = {'preserved_parts': 0, 'changed_parts': []}
        if template:
            preservation = apply_template(template, plan, generated, candidate, results)
            for patch in plan:
                manifest.append({k: patch[k] for k in ['sheet', 'cell', 'value', 'format']})
                manifest.append({'sheet': patch['sheet'], 'cell': patch['month_cell'],
                                 'value': patch['month_label'], 'format': '@'})
            old = read_parts(template)
            new = read_parts(candidate)
            for name, data in old.items():
                if name not in preservation['changed_parts'] and new.get(name) != data:
                    raise ValueError(f'未授权部件变化: {name}')
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', UserWarning)
                a = openpyxl.load_workbook(template)
                b = openpyxl.load_workbook(candidate)
            for name in a.sheetnames:
                if name in ['月度指标', '有效样本', '环比变化', '口径与映射']:
                    continue
                for key, dim in a[name].column_dimensions.items():
                    actual = b[name].column_dimensions[key]
                    if (dim.width, dim.hidden) != (actual.width, actual.hidden):
                        raise ValueError(f'模板列宽/隐藏属性变化: {name}/{key}')
                for key, dim in a[name].row_dimensions.items():
                    actual = b[name].row_dimensions[key]
                    if (dim.height, dim.hidden) != (actual.height, actual.hidden):
                        raise ValueError(f'模板行高/隐藏属性变化: {name}/{key}')
        else:
            parts = sync_chart_caches(read_parts(generated))
            with zipfile.ZipFile(candidate, 'w', zipfile.ZIP_DEFLATED) as z:
                for name, data in parts.items():
                    z.writestr(name, data)
        report = verify(candidate, manifest)
        report.update(preservation)
        report.update({'output': str(output), 'sha256': digest(candidate), 'manifest': manifest,
                       'inputs': hashes, 'results': results, 'backup': None,
                       'config_sha256': hashlib.sha256(json.dumps(job['config'], ensure_ascii=False, sort_keys=True).encode()).hexdigest(),
                       'warnings': ['月间样本/清洗口径不同不等同于版本因果；颜色不代表显著性。']})
        audit_temp = Path(temp) / 'audit.json'
        audit_temp.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
        ensure_unlocked(output)
        if output.exists():
            backup = output.with_name(output.stem + f'.backup_{stamp}.xlsx')
            shutil.copy2(output, backup)
            report['backup'] = str(backup)
            audit_temp.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
        os.replace(candidate, output)
        os.replace(audit_temp, str(output) + '.audit.json')
    return {k: v for k, v in report.items() if k not in ['manifest', 'results']}
