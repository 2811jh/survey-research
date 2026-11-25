import importlib
import json
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from test_survey_excel import config, frame


class DeliveryTests(unittest.TestCase):
    def setUp(self):
        try:
            self.api = importlib.import_module('excel_toolkit.validation')
        except ModuleNotFoundError:
            self.fail('尚未实现可验收交付入口')
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        source = self.base / 'answers.csv'
        frame().to_csv(source, index=False, encoding='utf-8-sig')
        self.job = {'config': config(), 'inputs': [{'month': '2026-07', 'path': str(source)}],
                    'output': str(self.base / 'result.xlsx')}

    def test_new_report_precision_month_text_and_audit(self):
        report = self.api.build(self.job)
        self.assertEqual(report['status'], 'verified')
        w = openpyxl.load_workbook(self.job['output'])
        self.assertEqual(w['月度指标']['C2'].value, '2026年7月')
        self.assertAlmostEqual(w['月度指标']['C3'].value, 8 / 3, places=14)
        self.assertEqual(w['月度指标']['C3'].number_format, '0.00')
        self.assertEqual(w['有效样本']['C3'].value, 3)
        self.assertTrue(Path(self.job['output'] + '.audit.json').exists())
        self.assertEqual(self.api.verify(self.job['output'])['status'], 'verified')

    def test_default_no_overwrite_and_explicit_backup(self):
        self.api.build(self.job)
        old = Path(self.job['output']).read_bytes()
        with self.assertRaises(FileExistsError):
            self.api.build(self.job)
        self.assertEqual(old, Path(self.job['output']).read_bytes())
        report = self.api.build(self.job, overwrite=True)
        self.assertEqual(Path(report['backup']).read_bytes(), old)

    def test_lock_duplicate_months_and_invalid_code_leave_target_unchanged(self):
        path = Path(self.job['output'])
        path.write_bytes(b'original')
        lock = path.with_name('~$' + path.name)
        lock.write_text('lock')
        with self.assertRaises(PermissionError):
            self.api.build(self.job, overwrite=True)
        lock.unlink()
        self.job['inputs'] *= 2
        with self.assertRaises(ValueError):
            self.api.build(self.job, overwrite=True)
        self.assertEqual(path.read_bytes(), b'original')

    def test_corrupted_output_rejected_by_verify(self):
        self.api.build(self.job)
        w = openpyxl.load_workbook(self.job['output'])
        w['月度指标']['C3'] = 99
        w.save(self.job['output'])
        with self.assertRaises(ValueError):
            self.api.verify(self.job['output'])

    def template(self):
        p = self.base / 'template.xlsx'
        w = openpyxl.Workbook()
        s = w.active
        s.title = '原表'
        s['B2'] = '整体满意度'
        s['C1'] = '2026年7月'
        s['C2'] = 4.2
        s['C2'].font = openpyxl.styles.Font(name='微软雅黑', bold=True)
        s['C2'].number_format = '0.0000'
        s['A10'] = '历史必须保留'
        s.column_dimensions['C'].width = 23
        s.row_dimensions[2].height = 34
        w.save(p)
        with zipfile.ZipFile(p, 'a') as z:
            z.writestr('customXml/private-preserve.xml', b'<preserve data="unchanged"/>')
        self.job['template'] = str(p)
        self.job['bindings'] = [{'sheet': '原表', 'metric': 'overall', 'label_column': 'B',
                                 'label': '整体满意度', 'header_row': 1, 'month': '2026-07'}]
        return p

    def test_template_preserves_unknown_parts_layout_and_binds_semantically(self):
        template = self.template()
        old = openpyxl.load_workbook(template)
        report = self.api.build(self.job)
        new = openpyxl.load_workbook(self.job['output'])
        self.assertAlmostEqual(new['原表']['C2'].value, 8 / 3, places=14)
        import copy
        self.assertEqual(copy.copy(new['原表']['C2'].font), copy.copy(old['原表']['C2'].font))
        self.assertEqual(new['原表'].row_dimensions[2].height, 34)
        self.assertEqual(new['原表'].column_dimensions['C'].width, 23)
        self.assertEqual(new['原表']['A10'].value, '历史必须保留')
        self.assertIn('月度指标', new.sheetnames)
        with zipfile.ZipFile(self.job['output']) as z:
            self.assertEqual(z.read('customXml/private-preserve.xml'), b'<preserve data="unchanged"/>')
        self.assertGreater(report['preserved_parts'], 0)

    def test_template_ambiguous_label_stops_without_output(self):
        p = self.template()
        w = openpyxl.load_workbook(p)
        w['原表']['B3'] = '整体满意度'
        w.save(p)
        with self.assertRaisesRegex(ValueError, '歧义'):
            self.api.build(self.job)
        self.assertFalse(Path(self.job['output']).exists())

    def test_template_requires_bindings_or_explicit_append_only(self):
        self.template()
        self.job.pop('bindings')
        with self.assertRaises(ValueError):
            self.api.build(self.job)
        self.job['append_only'] = True
        self.api.build(self.job)
        self.assertEqual(openpyxl.load_workbook(self.job['output'])['原表']['C2'].value, 4.2)

    def test_schema_prefix_in_ignorable_extension_survives(self):
        p = self.template()
        from excel_toolkit.template import xml, encoded
        raw = b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" mc:Ignorable="x14"><sheetData/></worksheet>'
        result = encoded(xml(raw))
        self.assertIn(b'xmlns:x14=', result)

    def test_template_row_scope_disambiguates_repeated_blocks(self):
        p = self.template()
        w = openpyxl.load_workbook(p)
        w['原表']['B8'] = '整体满意度'
        w.save(p)
        self.job['bindings'][0]['row_range'] = [2, 4]
        self.api.build(self.job)
        self.assertAlmostEqual(openpyxl.load_workbook(self.job['output'])['原表']['C2'].value, 8 / 3)

    def test_failed_code_validation_never_replaces_target(self):
        output = Path(self.job['output'])
        output.write_bytes(b'must survive')
        data = frame()
        data['Q1.整体满意度'] = 99
        data.to_csv(self.job['inputs'][0]['path'], index=False)
        with self.assertRaises(ValueError):
            self.api.build(self.job, overwrite=True)
        self.assertEqual(output.read_bytes(), b'must survive')

    def test_same_raw_input_as_output_and_readonly_stop(self):
        raw = self.base / 'raw.xlsx'
        frame().to_excel(raw, index=False)
        self.job['inputs'][0]['path'] = str(raw)
        self.job['output'] = str(raw)
        old = raw.read_bytes()
        with self.assertRaises(ValueError):
            self.api.build(self.job, overwrite=True)
        self.assertEqual(raw.read_bytes(), old)

    def test_chart_caches_and_rules_are_valid_and_nonduplicated(self):
        self.job['inputs'].append(dict(self.job['inputs'][0], month='2026-08'))
        self.api.build(self.job)
        from excel_toolkit.template import C, NS, read_parts, xml
        parts = read_parts(self.job['output'])
        chart = xml(parts['xl/charts/chart1.xml'])
        self.assertIn('2026年7月', [x.text for x in chart.findall('.//c:strCache/c:pt/c:v', NS)])
        w = openpyxl.load_workbook(self.job['output'])
        for sheet in w:
            for cf, rules in sheet.conditional_formatting._cf_rules.items():
                self.assertNotIn('A2', cf.sqref)
                self.assertEqual(len(rules), len({(r.type, tuple(r.formula)) for r in rules}))

    def test_repeat_rejects_incomplete_month_history(self):
        self.template()
        self.job['inputs'].append(dict(self.job['inputs'][0], month='2026-08'))
        self.api.build(self.job)
        self.job['template'] = self.job['output']
        self.job['output'] = str(self.base / 'missing_history.xlsx')
        self.job['inputs'].pop()
        with self.assertRaisesRegex(ValueError, '历史月份'):
            self.api.build(self.job)
        self.assertFalse(Path(self.job['output']).exists())

    def test_repeat_template_run_preserves_history_and_rule_count(self):
        self.template()
        self.api.build(self.job)
        self.job['template'] = self.job['output']
        self.job['output'] = str(self.base / 'second.xlsx')
        self.api.build(self.job)
        w = openpyxl.load_workbook(self.job['output'])
        self.assertEqual(len([n for n in w.sheetnames if n.startswith('月度指标')]), 1)
        self.assertEqual(w['原表']['A10'].value, '历史必须保留')


if __name__ == '__main__':
    unittest.main()
