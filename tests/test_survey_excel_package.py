import importlib.util
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]


class PackagingTests(unittest.TestCase):
    def test_fixed_trigger_routed_before_generic_analysis(self):
        skill = (ROOT / 'SKILL.md').read_text(encoding='utf-8')
        self.assertIn('固定触发词 **「月度汇总 Excel」**', skill)
        self.assertIn('references/19-excel-workflow.md', skill)
        self.assertLess(skill.index('## 优先路由'), skill.index('## 数据来源路由'))

    def test_standalone_archive_cli_and_no_secrets(self):
        spec = importlib.util.spec_from_file_location('pack', ROOT / 'scripts/package_excel_toolkit.py')
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'toolkit.zip'
            module.package(output)
            with zipfile.ZipFile(output) as z:
                names = z.namelist()
                self.assertFalse(any(n.endswith('config.json') or '__pycache__' in n or n.endswith('.csv') for n in names))
                self.assertIn('survey-research/scripts/survey_excel.py', names)
                z.extractall(Path(directory) / 'unpacked')
            cli = Path(directory) / 'unpacked/survey-research/scripts/survey_excel.py'
            result = subprocess.run([sys.executable, '-X', 'utf8', str(cli), '--help'], capture_output=True, text=True, encoding='utf-8')
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn('preview', result.stdout)


if __name__ == '__main__':
    unittest.main()
