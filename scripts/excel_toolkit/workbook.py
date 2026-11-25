"""标准月报渲染：数值与显示分离，规则只覆盖数据。"""
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


REPORT_SHEETS = ['月度指标', '有效样本', '环比变化', '口径与映射']


def month_label(month):
    """返回不被Excel推断成日期的明确月份文字。"""
    year, mon = month.split('-')
    return f'{year}年{int(mon)}月'


def metric_format(item):
    """按统计类型返回两位显示格式，不舍入数值。"""
    return '0.00' if item['kind'] == 'mean' else '0.00%'


def color_rules(ws, area, item):
    """固定量尺，避免月间自适应颜色造成误读。"""
    if item['kind'] == 'proportion':
        rule = DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='7F9DB9')
    else:
        lo, mid, hi = (-1, 0, 1) if item['kind'] == 'nps' else (min(item['allowed']),
                         (min(item['allowed']) + max(item['allowed'])) / 2, max(item['allowed']))
        rule = ColorScaleRule(start_type='num', start_value=lo, start_color='F2CECE',
                              mid_type='num', mid_value=mid, mid_color='FFF2CC',
                              end_type='num', end_value=hi, end_color='D9EAD3')
    ws.conditional_formatting.add(area, rule)


def sign_rules(ws, area, item):
    """方向已知才评价好坏；问题占比上升用红色。"""
    direction = item.get('direction', 'neutral')
    if direction == 'neutral':
        return
    first = area.split(':')[0]
    positive, negative = ('008000', 'C00000') if direction == 'higher' else ('C00000', '008000')
    for op, color in [('>', positive), ('<', negative)]:
        ws.conditional_formatting.add(area, FormulaRule(
            formula=[f'AND(ISNUMBER({first}),{first}{op}0)'], font=Font(color=color)))


def record(ws, address, value, fmt, manifest):
    """写单元格及可独立回读的预期清单，字符串强制文本。"""
    cell = ws[address]
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'
    cell.number_format = fmt
    manifest.append({'sheet': ws.title, 'cell': address, 'value': value, 'format': fmt})


def render_new(results, config, path):
    """生成完整月度明细与有效分母，不写入任何个人答卷。"""
    book = Workbook()
    book.remove(book.active)
    for name in REPORT_SHEETS:
        book.create_sheet(name)
    manifest = []
    months = sorted(results)
    groups = ['全体'] + [g['label'] for g in config.get('groups', [])]
    for name in REPORT_SHEETS[:3]:
        ws = book[name]
        record(ws, 'A1', name + '（完整精度，显示两位；颜色不代表显著性）', '@', manifest)
        record(ws, 'A2', '指标', '@', manifest)
        record(ws, 'B2', '人群', '@', manifest)
        for c, month in enumerate(months, 3):
            record(ws, ws.cell(2, c).coordinate, month_label(month), '@', manifest)
    row = 3
    for item in config['metrics']:
        for group in groups:
            for name in REPORT_SHEETS[:3]:
                record(book[name], f'A{row}', item['label'], '@', manifest)
                record(book[name], f'B{row}', group, '@', manifest)
            for c, month in enumerate(months, 3):
                entry = results[month]['metrics'][item['id']][group]
                addr = book[REPORT_SHEETS[0]].cell(row, c).coordinate
                record(book['月度指标'], addr, entry['value'], metric_format(item), manifest)
                record(book['有效样本'], addr, entry['n'], '0', manifest)
                prev_month = months[c - 4] if c > 3 else None
                previous = results[prev_month]['metrics'][item['id']][group]['value'] if prev_month else None
                adjacent = prev_month and (int(month[:4])*12+int(month[5:]) - int(prev_month[:4])*12-int(prev_month[5:]) == 1)
                delta = None
                if adjacent and entry['value'] is not None and previous is not None:
                    delta = entry['value'] - previous
                    if item['kind'] != 'mean':
                        delta *= 100
                record(book['环比变化'], addr, delta, '0.00' if item['kind'] == 'mean' else '0.00"pp"', manifest)
            area = f'C{row}:{book["月度指标"].cell(row, len(months)+2).coordinate}'
            color_rules(book['月度指标'], area, item)
            sign_rules(book['环比变化'], area, item)
            row += 1
    ws = book['口径与映射']
    headers = ['月份', '指标ID', '指标', '人群', '原始列', '有效分母', '人群人数', '缺失数', '分母口径']
    for c, value in enumerate(headers, 1):
        record(ws, ws.cell(1, c).coordinate, value, '@', manifest)
    r = 2
    for month in months:
        for item in config['metrics']:
            for group in groups:
                value = results[month]['metrics'][item['id']][group]
                values = [month_label(month), item['id'], item['label'], group,
                          results[month]['mapping'][item['id']] or '未出现的可选题', value['n'],
                          value['group_rows'], value['missing'], value['denominator']]
                for c, value in enumerate(values, 1):
                    record(ws, ws.cell(r, c).coordinate, value, '0' if c in [6, 7, 8] else '@', manifest)
                r += 1
    for ws in book:
        ws.freeze_panes = 'C3' if ws.title != '口径与映射' else 'F2'
        ws.sheet_view.showGridLines = False
        ws.column_dimensions['A'].width = 32
        ws.column_dimensions['B'].width = 20
        for row_cells in ws:
            for cell in row_cells:
                cell.font = Font(name='微软雅黑', size=11)
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.border = Border(bottom=Side(style='hair', color='DCE4E8'))
                if cell.row == (1 if ws.title == '口径与映射' else 2):
                    cell.fill = PatternFill('solid', fgColor='D9E5EC')
                    cell.font = Font(name='微软雅黑', size=11, bold=True, color='17354A')
                if cell.column >= 3:
                    ws.column_dimensions[cell.column_letter].width = 18
        ws.auto_filter.ref = f'A{1 if ws.title == "口径与映射" else 2}:{ws.cell(ws.max_row, ws.max_column).coordinate}'
    if len(months) >= 2:
        chart = LineChart()
        chart.title = '整体指标趋势（首个指标）'
        chart.y_axis.numFmt = metric_format(config['metrics'][0])
        chart.add_data(Reference(book['月度指标'], min_col=3, max_col=len(months)+2, min_row=3, max_row=3), from_rows=True)
        chart.set_categories(Reference(book['月度指标'], min_col=3, max_col=len(months)+2, min_row=2, max_row=2))
        book['月度指标'].add_chart(chart, f'A{row+3}')
    book.save(path)
    return manifest
