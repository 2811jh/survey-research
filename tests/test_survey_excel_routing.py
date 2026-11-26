import importlib
from pathlib import Path
import sys
import unittest

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))


class RoutingTests(unittest.TestCase):
    def api(self):
        try:
            return importlib.import_module('excel_toolkit.routing')
        except ModuleNotFoundError:
            self.fail('缺少材料识别及意图分流工具')

    def test_trigger_and_negative_intent(self):
        route = self.api().trigger
        for text in ['月度汇总 Excel', '更新11月满意度汇总表', '用原始问卷生成月度Excel', '在上个月汇总基础上补上本月数据']:
            self.assertTrue(route(text), text)
        for text in ['只下载问卷，不生成Excel', '分析开放题', '解释NPS算法', '不要更新月度汇总 Excel']:
            self.assertFalse(route(text), text)

    def test_material_shape_not_extension(self):
        detect = self.api().classify_frame
        self.assertEqual(detect(pd.DataFrame({'序号':[1,2], 'Q1.满意度':[4,5]})), 'raw')
        self.assertEqual(detect(pd.DataFrame([['指标','2026年10月'],['满意度',4.1]])), 'summary')
        self.assertEqual(detect(pd.DataFrame([['指标','2026年10月'],['均值',4.1],['样本量',10]])), 'summary')
        self.assertEqual(detect(pd.DataFrame({'姓名':['小王']})), 'unknown')

    def test_route_material_combinations(self):
        route = self.api().choose_flow
        self.assertEqual(route(['raw'], False), 'new_from_raw')
        self.assertEqual(route(['raw','summary'], True), 'update_from_raw_and_summary')
        self.assertEqual(route(['summary'], True), 'update_from_summary')
        self.assertEqual(route(['summary'], False), 'assemble_summaries')
        with self.assertRaises(ValueError):route(['unknown'], False)

    def test_summary_missing_denominator_not_invented(self):
        merge = self.api().validate_summary_record
        self.assertEqual(merge({'kind':'mean','value':4.1,'n':None})['n'], None)
        with self.assertRaisesRegex(ValueError, '分母'):
            merge({'kind':'mean','value':4.1,'n':None}, require_pooling=True)
        with self.assertRaisesRegex(ValueError, 'NPS'):
            merge({'kind':'nps','value':29.5,'n':100})


if __name__ == '__main__':
    unittest.main()
