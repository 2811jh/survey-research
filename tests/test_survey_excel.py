import copy
import importlib
import json
from pathlib import Path
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))


def config():
    return {
        'metrics': [
            {'id': 'overall', 'label': '整体满意度', 'kind': 'mean',
             'match': {'contains': ['整体满意度']}, 'allowed': [1, 2, 3, 4, 5], 'direction': 'higher'},
            {'id': 'nps', 'label': 'NPS', 'kind': 'nps',
             'match': {'contains': ['推荐给其他人']}, 'allowed': list(range(11)), 'direction': 'higher'},
            {'id': 'issue', 'label': '性能问题', 'kind': 'proportion',
             'match': {'contains': ['多选:性能']}, 'selected': [1], 'unselected': [0],
             'denominator': {'mode': 'answered', 'match': {'contains': ['多选:']}}, 'direction': 'lower'}],
        'groups': [{'id': 'male', 'label': '男', 'match': {'contains': ['性别']},
                    'allowed': [1, 2, 3], 'values': [1]}]}


def frame():
    return pd.DataFrame({'Q1.整体满意度': [1, 2, 5], 'Q9.推荐给其他人': [0, 8, 10],
                         'Q4.多选:性能': [1, 0, None], 'Q4.多选:其他': [0, 1, None],
                         'Q7.性别': [1, 1, 2]})


class MetricsTests(unittest.TestCase):
    def setUp(self):
        try:
            self.api = importlib.import_module('excel_toolkit.metrics')
        except ModuleNotFoundError:
            self.fail('尚未实现 excel_toolkit.metrics')

    def test_full_precision_and_metric_specific_denominators(self):
        result = self.api.compute_month(frame(), config())
        self.assertEqual(result['metrics']['overall']['全体']['value'], 8 / 3)
        self.assertEqual(result['metrics']['nps']['全体']['value'], 0)
        self.assertEqual(result['metrics']['issue']['全体']['value'], 0.5)
        self.assertEqual(result['metrics']['issue']['全体']['n'], 2)
        self.assertEqual(result['metrics']['overall']['男']['value'], 1.5)

    def test_q_number_changes_not_identity(self):
        data = frame().rename(columns={'Q1.整体满意度': 'Q88.整体满意度'})
        self.assertEqual(self.api.compute_month(data, config())['metrics']['overall']['全体']['value'], 8 / 3)

    def test_ambiguous_match_stops(self):
        data = frame().assign(**{'Q99.整体满意度': [1, 2, 3]})
        with self.assertRaisesRegex(ValueError, '歧义'):
            self.api.compute_month(data, config())

    def test_unknown_code_and_non_numeric_stop(self):
        for value in [99, '坏数据']:
            data = frame()
            data['Q1.整体满意度'] = data['Q1.整体满意度'].astype(object)
            data.loc[0, 'Q1.整体满意度'] = value
            with self.assertRaises(ValueError):
                self.api.compute_month(data, config())

    def test_missing_metric_does_not_drop_entire_response(self):
        data = frame()
        data.loc[0, 'Q1.整体满意度'] = None
        result = self.api.compute_month(data, config())
        self.assertEqual(result['rows'], 3)
        self.assertEqual(result['metrics']['overall']['全体']['n'], 2)
        self.assertEqual(result['metrics']['nps']['全体']['n'], 3)

    def test_optional_missing_returns_null_not_zero(self):
        cfg = config()
        cfg['metrics'][0]['optional'] = True
        result = self.api.compute_month(frame().drop(columns=['Q1.整体满意度']), cfg)
        self.assertIsNone(result['metrics']['overall']['全体']['value'])
        self.assertEqual(result['metrics']['overall']['全体']['n'], 0)

    def test_nps_boundaries(self):
        data = pd.DataFrame({'推荐给其他人': [0, 6, 7, 8, 9, 10]})
        cfg = {'metrics': [config()['metrics'][1]], 'groups': []}
        result = self.api.compute_month(data, cfg)['metrics']['nps']['全体']
        self.assertEqual(result['value'], 0)
        self.assertEqual(result['n'], 6)

    def test_one_blank_multi_select_and_eligibility(self):
        data = frame().drop(columns=['Q4.多选:其他'])
        data['Q4.多选:性能'] = [1, None, None]
        cfg = config()
        cfg['metrics'][2]['unselected'] = []
        cfg['metrics'][2]['denominator'] = {'mode': 'all'}
        self.assertEqual(self.api.compute_month(data, cfg)['metrics']['issue']['全体']['value'], 1 / 3)
        cfg['metrics'][2]['denominator'] = {'mode': 'eligible', 'match': {'contains': ['性别']}, 'values': [1]}
        self.assertEqual(self.api.compute_month(data, cfg)['metrics']['issue']['全体']['value'], 0.5)

    def test_unknown_group_code_stops(self):
        data = frame()
        data.loc[0, 'Q7.性别'] = 4
        with self.assertRaises(ValueError):
            self.api.compute_month(data, config())

    def test_missing_required_and_empty_group(self):
        with self.assertRaises(ValueError):
            self.api.compute_month(frame().drop(columns=['Q1.整体满意度']), config())
        data = frame()
        data['Q7.性别'] = 2
        result = self.api.compute_month(data, config())
        self.assertIsNone(result['metrics']['overall']['男']['value'])

    def test_month_validation_and_duplicate(self):
        for values in [['2026-7'], ['2026-13'], ['2026-07', '2026-07']]:
            with self.assertRaises(ValueError):
                self.api.validate_months(values)

    def test_footer_filter_requires_explicit_identifier(self):
        data = frame()
        data['序号'] = [1, 2, '网易内部文件']
        cfg = config()
        cfg['response_id'] = {'column': '序号', 'regex': r'^\d+$'}
        result = self.api.compute_month(data, cfg)
        self.assertEqual(result['rows'], 2)
        self.assertEqual(result['excluded_rows'], 1)

    def test_duplicate_labels_and_unknown_direction_stop(self):
        for change in ['label', 'direction']:
            cfg = config()
            if change == 'label':
                cfg['metrics'][1]['label'] = '整体满意度'
            else:
                cfg['metrics'][0]['direction'] = 'sometimes'
            with self.assertRaises(ValueError):
                self.api.compute_month(frame(), cfg)


if __name__ == '__main__':
    unittest.main()
