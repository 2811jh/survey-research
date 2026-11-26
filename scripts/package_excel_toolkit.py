"""白名单分发月度汇总工具包，不包含凭据和真实答卷。"""
import argparse
from pathlib import Path
import zipfile


def package(output):
    """打包独立入口、配置、测试及扩展skill操作说明。"""
    root = Path(__file__).resolve().parents[1]
    files = [root / 'SKILL.md', root / 'references/19-excel-workflow.md',
             root / 'scripts/survey_excel.py', root / 'scripts/package_excel_toolkit.py',
             root / 'tests/test_survey_excel_routing.py',
             root / 'assets/excel/minecraft-monthly.json', root / 'assets/excel/job.example.json',
             root / 'tests/test_survey_excel.py', root / 'tests/test_survey_excel_template.py',
             root / 'tests/test_survey_excel_package.py']
    files += sorted((root / 'scripts/excel_toolkit').glob('*.py'))
    output = Path(output)
    if output.exists():
        raise FileExistsError('分发文件已存在，请指定新文件名')
    if not output.parent.is_dir():
        raise ValueError('输出目录不存在')
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for path in files:
            content = path.read_bytes()
            if path.name == 'SKILL.md':
                content = content.replace(b'C:\\Users\\xxx\\Desktop\\', b'{DESKTOP}/')
            if path.name != 'package_excel_toolkit.py' and (b'C:\\Users\\' in content or b'C:/Users/' in content):
                raise ValueError(f'分发内容含本机绝对路径: {path.name}')
            z.writestr('survey-research/' + path.relative_to(root).as_posix(), content)
        z.writestr('survey-research/requirements-excel.txt', 'pandas>=1.5.0\nopenpyxl>=3.0.10\n')
        z.writestr('INSTALL.txt', '月度汇总 Excel 工具包\n将survey-research目录合并到已有同名skill，或直接运行其中scripts/survey_excel.py。\n安装依赖：py -m pip install -r survey-research/requirements-excel.txt\n说明：survey-research/references/19-excel-workflow.md\n离线独立使用支持月度Excel功能；下载/文本等既有功能需完整原skill。\n测试：py -m unittest discover -s survey-research/tests -p test_survey_excel*.py\n')
    return {'output': str(output), 'files': len(files) + 2}


if __name__ == '__main__':
    import json
    p = argparse.ArgumentParser()
    p.add_argument('--output', required=True)
    print(json.dumps(package(p.parse_args().output), ensure_ascii=False))
