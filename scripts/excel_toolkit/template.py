"""OOXML定点更新：不整本重存用户模板，未触碰部件逐字节保留。"""
import copy
from datetime import datetime
from pathlib import PurePosixPath
import posixpath
import re
import warnings
import xml.etree.ElementTree as ET
import zipfile

import openpyxl

from .workbook import REPORT_SHEETS, metric_format, month_label

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
T = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'s': S, 'c': C}


_PREFIXES = {}


def xml(data):
    """解析XML且记录原前缀，保留未知扩展的命名空间前缀。"""
    import io
    for _, pair in ET.iterparse(io.BytesIO(data), events=['start-ns']):
        prefix, uri = pair
        if not re.fullmatch(r'ns\d+', prefix):
            ET.register_namespace(prefix, uri)
            _PREFIXES[prefix] = uri
    return ET.fromstring(data)


def encoded(root):
    """保留mc:Ignorable等属性引用的命名空间声明。"""
    result = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    prefixes = set()
    for element in root.iter():
        for attr, value in element.attrib.items():
            if attr.endswith('}Ignorable'):
                prefixes.update(value.split())
    for prefix in prefixes:
        declaration = ('xmlns:' + prefix + '=').encode()
        if declaration not in result:
            if prefix not in _PREFIXES:
                raise ValueError(f'未知扩展前缀: {prefix}')
            from xml.sax.saxutils import quoteattr
            root.set('xmlns:' + prefix, _PREFIXES[prefix])
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def read_parts(path):
    """加载包的原始部件字节。"""
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def sheet_paths(parts):
    """通过关系定位工作表，不假定sheet1顺序。"""
    rels = xml(parts['xl/_rels/workbook.xml.rels'])
    lookup = {r.get('Id'): r.get('Target') for r in rels}
    root = xml(parts['xl/workbook.xml'])
    return {s.get('name'): ('xl/' + lookup[s.get('{' + R + '}id')]).replace('xl//xl/', 'xl/')
            if not lookup[s.get('{' + R + '}id')].startswith('/') else lookup[s.get('{' + R + '}id')].lstrip('/')
            for s in root.find('s:sheets', NS)}


def value_at(parts, sheet, address):
    """读取内联、共享、数字和公式缓存值。"""
    root = xml(parts[sheet_paths(parts)[sheet]])
    cell = root.find(f'.//s:c[@r="{address}"]', NS)
    if cell is None:
        return None
    if cell.get('t') == 'inlineStr':
        return ''.join(cell.itertext())
    node = cell.find('s:v', NS)
    if node is None:
        return None
    if cell.get('t') == 's':
        strings = xml(parts['xl/sharedStrings.xml'])
        return ''.join(strings[int(node.text)].itertext())
    if cell.get('t') in ['str', 'e']:
        return node.text
    return float(node.text)


def plan_template(template, bindings, results, config, append_only=False):
    """唯一语义定位行标题与月份表头，缺失/歧义拒绝更新。"""
    if not bindings and not append_only:
        raise ValueError('模板模式必须提供bindings或明确append_only=true')
    with warnings.catch_warnings():
        warnings.simplefilter('ignore', UserWarning)
        wb = openpyxl.load_workbook(template, data_only=False)
    items = {m['id']: m for m in config['metrics']}
    plan = []
    used = set()
    for binding in bindings:
        ws = wb[binding['sheet']]
        item = items[binding['metric']]
        labels = binding.get('aliases', []) + [binding['label']]
        def head(v):
            return re.split(r'[\n（(]', str(v or ''))[0].strip()
        rows = [c.row for c in ws[binding['label_column']] if head(c.value) in labels
                and binding.get('row_range', [1, ws.max_row])[0] <= c.row <= binding.get('row_range', [1, ws.max_row])[1]]
        if len(rows) != 1:
            raise ValueError(f'模板行标题缺失/歧义: {binding["label"]}')
        month = binding.get('month', max(results))
        columns = []
        for cell in ws[binding['header_row']]:
            v = cell.value
            expected = {month, month_label(month), f'{int(month[:4]) % 100}年-{int(month[5:])}月',
                        f'{int(month[:4]) % 100}年{int(month[5:])}月'}
            if isinstance(v, datetime):
                v = v.strftime('%Y-%m')
            if v in expected:
                columns.append(cell.column)
        if binding.get('column'):
            col = openpyxl.utils.column_index_from_string(binding['column'])
            columns = [c for c in columns if c == col]
        if len(columns) != 1:
            raise ValueError(f'模板月份表头缺失/歧义: {month}；不猜测新增月份空位')
        address = ws.cell(rows[0], columns[0]).coordinate
        identity = (ws.title, address)
        if identity in used:
            raise ValueError('模板单元格重复绑定')
        used.add(identity)
        if ws[address].data_type == 'f':
            raise ValueError('绑定位置包含公式，禁止无提示覆盖')
        group = binding.get('group', '全体')
        entry = results[month]['metrics'][item['id']][group]
        field = binding.get('field', 'value')
        if field not in ['value', 'n']:
            raise ValueError('绑定仅支持value或n')
        plan.append({'sheet': ws.title, 'cell': address, 'value': entry[field],
                     'format': metric_format(item) if field == 'value' else '0',
                     'month_cell': ws.cell(binding['header_row'], columns[0]).coordinate,
                     'month_label': month_label(month)})
    return plan


def merge_styles(base, donor):
    """合并生成表样式并重映射所有全局样式索引。"""
    a, b = xml(base['xl/styles.xml']), xml(donor['xl/styles.xml'])
    offsets = {}
    for name in ['fonts', 'fills', 'borders', 'cellStyleXfs', 'dxfs']:
        target, source = a.find('s:' + name, NS), b.find('s:' + name, NS)
        if target is None:
            target = ET.SubElement(a, '{' + S + '}' + name, count='0')
        offsets[name] = len(target)
        if source is not None:
            for child in source:
                target.append(copy.deepcopy(child))
        target.set('count', str(len(target)))
    nums = a.find('s:numFmts', NS)
    if nums is None:
        nums = ET.Element('{' + S + '}numFmts', count='0')
        a.insert(0, nums)
    num_map = {}
    next_id = max([163] + [int(x.get('numFmtId')) for x in nums]) + 1
    for n in b.findall('s:numFmts/s:numFmt', NS):
        num_map[int(n.get('numFmtId'))] = next_id
        new = copy.deepcopy(n)
        new.set('numFmtId', str(next_id))
        nums.append(new)
        next_id += 1
    nums.set('count', str(len(nums)))
    target, source = a.find('s:cellXfs', NS), b.find('s:cellXfs', NS)
    offsets['cellXfs'] = len(target)
    for source_xf in source:
        xf = copy.deepcopy(source_xf)
        for attr, collection in [('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders'), ('xfId', 'cellStyleXfs')]:
            if attr in xf.attrib:
                xf.set(attr, str(int(xf.get(attr)) + offsets[collection]))
        old = int(xf.get('numFmtId', '0'))
        if old in num_map:
            xf.set('numFmtId', str(num_map[old]))
        target.append(xf)
    target.set('count', str(len(target)))
    base['xl/styles.xml'] = encoded(a)
    return offsets


def write_cell(root, address, value):
    """替换单元格值，保留其原样式，不删除其他单元格。"""
    cell = root.find(f'.//s:c[@r="{address}"]', NS)
    if cell is None:
        raise ValueError(f'模板目标单元格不存在: {address}')
    for child in list(cell):
        if child.tag in ['{' + S + '}' + n for n in ['f', 'v', 'is']]:
            cell.remove(child)
    cell.attrib.pop('t', None)
    if value is not None:
        if isinstance(value, str):
            cell.set('t', 'inlineStr')
            ET.SubElement(ET.SubElement(cell, '{' + S + '}is'), '{' + S + '}t').text = value
        else:
            ET.SubElement(cell, '{' + S + '}v').text = repr(value)
    return cell


def apply_template(template, plan, generated, output, results):
    """追加完整月度明细，原模板历史及未知部件保留。"""
    base, donor = read_parts(template), read_parts(generated)
    original = dict(base)
    paths = sheet_paths(base)
    changes = {}
    styles = xml(base['xl/styles.xml'])
    xfs = styles.find('s:cellXfs', NS)
    for patch in plan:
        path = paths[patch['sheet']]
        root = changes.setdefault(path, xml(base[path]))
        target = write_cell(root, patch['cell'], patch['value'])
        fmt_id = 2 if patch['format'] == '0.00' else (10 if patch['format'] == '0.00%' else 1)
        xf = copy.deepcopy(xfs[int(target.get('s', '0'))])
        xf.set('numFmtId', str(fmt_id))
        xf.set('applyNumberFormat', '1')
        target.set('s', str(len(xfs)))
        xfs.append(xf)
        target = write_cell(root, patch['month_cell'], patch['month_label'])
        xf = copy.deepcopy(xfs[int(target.get('s', '0'))])
        xf.set('numFmtId', '49')
        target.set('s', str(len(xfs)))
        xfs.append(xf)
    xfs.set('count', str(len(xfs)))
    base['xl/styles.xml'] = encoded(styles)
    for path, root in changes.items():
        base[path] = encoded(root)
    offsets = merge_styles(base, donor)
    book = xml(base['xl/workbook.xml'])
    rels = xml(base['xl/_rels/workbook.xml.rels'])
    types = xml(base['[Content_Types].xml'])
    sheets = book.find('s:sheets', NS)
    existing = {s.get('name'): s for s in sheets}
    original_paths = sheet_paths(base)
    donor_paths = sheet_paths(donor)
    renames = {}
    for name, path in donor_paths.items():
        if name in existing:
            old = xml(base[original_paths[name]])
            months = []
            for cell in old.findall('.//s:row[@r="2"]/s:c', NS):
                text = ''.join(cell.itertext()) if cell.get('t') == 'inlineStr' else ''
                if re.fullmatch(r'\d{4}年\d{1,2}月', text):
                    months.append(text)
            if any(m not in [month_label(x) for x in results] for m in months):
                raise ValueError('更新数据未包含已生成表中的历史月份，请补齐输入，禁止丢弃历史')
        unique = 'monthly_' + str(len(rels)) + '_' + PurePosixPath(path).name
        renames[path] = 'xl/worksheets/' + unique
    for path in donor:
        if path.startswith(('xl/charts/', 'xl/drawings/')) and not path.endswith('/'):
            renames[path] = str(PurePosixPath(path).parent / ('monthly_' + str(len(rels)) + '_' + PurePosixPath(path).name))
    for path in donor:
        if '/_rels/' in path and path.endswith('.rels') and path != 'xl/_rels/workbook.xml.rels':
            parent, filename = path.split('/_rels/')
            origin = parent + '/' + filename[:-5]
            if origin in renames:
                dest = PurePosixPath(renames[origin])
                renames[path] = str(dest.parent / '_rels' / (dest.name + '.rels'))
    for source, target in renames.items():
        root = xml(donor[source])
        if source in donor_paths.values():
            for c in root.findall('.//s:c', NS):
                c.set('s', str(int(c.get('s', '0')) + offsets['cellXfs']))
            for rule in root.findall('.//s:cfRule', NS):
                if 'dxfId' in rule.attrib:
                    rule.set('dxfId', str(int(rule.get('dxfId')) + offsets['dxfs']))
        if source.endswith('.rels'):
            owner = source.replace('/_rels/', '/').removesuffix('.rels')
            new_owner = renames[owner]
            for rel in root:
                if rel.get('TargetMode') == 'External':
                    raise ValueError('生成报表不允许外部关系')
                absolute = posixpath.normpath(posixpath.join(posixpath.dirname(owner), rel.get('Target')))
                if rel.get('Target').startswith('/'):
                    absolute = rel.get('Target').lstrip('/')
                destination = renames.get(absolute, absolute)
                rel.set('Target', posixpath.relpath(destination, posixpath.dirname(new_owner)))
        base[target] = encoded(root)
    for name, path in donor_paths.items():
        if name in existing:
            rel_id = existing[name].get('{' + R + '}id')
            rel = next(x for x in rels if x.get('Id') == rel_id)
            rel.set('Target', '/' + renames[path])
        else:
            rel_id = 'rIdMonthly' + str(len(rels))
            ET.SubElement(rels, '{' + P + '}Relationship', Id=rel_id,
                          Type=R + '/worksheet', Target='/' + renames[path])
            sid = max([0] + [int(s.get('sheetId')) for s in sheets]) + 1
            ET.SubElement(sheets, '{' + S + '}sheet', name=name, sheetId=str(sid), attrib={'{' + R + '}id': rel_id})
    donor_types = xml(donor['[Content_Types].xml'])
    for override in donor_types:
        source = override.get('PartName', '').lstrip('/')
        if source in renames:
            new = copy.deepcopy(override)
            new.set('PartName', '/' + renames[source])
            types.append(new)
    base['xl/workbook.xml'] = encoded(book)
    base['xl/_rels/workbook.xml.rels'] = encoded(rels)
    base['[Content_Types].xml'] = encoded(types)
    base = sync_chart_caches(base)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in base.items():
            z.writestr(name, data)
    return {'preserved_parts': sum(base.get(k) == v for k, v in original.items()),
            'changed_parts': [k for k, v in original.items() if base.get(k) != v]}


def sync_chart_caches(parts):
    """刷新本地单元格引用的图表缓存；外部/复杂引用原样保留。"""
    for path, raw in list(parts.items()):
        if not path.startswith('xl/charts/') or not path.endswith('.xml'):
            continue
        root = xml(raw)
        changed = False
        for ref in list(root.iter()):
            if ref.tag not in ['{' + C + '}numRef', '{' + C + '}strRef']:
                continue
            formula = ref.find('c:f', NS)
            if formula is None:
                continue
            match = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^!]+))!([$A-Z]+\$?\d+(?::[$A-Z]+\$?\d+)?)", formula.text or '')
            if not match:
                continue
            name = (match[1] or match[2]).replace("''", "'")
            if name not in sheet_paths(parts):
                continue
            values = [value_at(parts, name, a) for row in openpyxl.utils.cell.rows_from_range(match[3].replace('$', '')) for a in row]
            string = any(isinstance(v, str) for v in values)
            ref.tag = '{' + C + '}' + ('strRef' if string else 'numRef')
            for child in list(ref):
                if child is not formula:
                    ref.remove(child)
            cache = ET.SubElement(ref, '{' + C + '}' + ('strCache' if string else 'numCache'))
            if not string:
                ET.SubElement(cache, '{' + C + '}formatCode').text = '0.00'
            ET.SubElement(cache, '{' + C + '}ptCount', val=str(len(values)))
            for i, value in enumerate(values):
                if value is not None:
                    ET.SubElement(ET.SubElement(cache, '{' + C + '}pt', idx=str(i)), '{' + C + '}v').text = str(value)
            changed = True
        if changed:
            parts[path] = encoded(root)
    return parts
