"""月度汇总 Excel 命令行入口。"""
import argparse
import json
from pathlib import Path
import sys

from excel_toolkit.metrics import load_frame
from excel_toolkit.routing import classify_frame, inspect_material, choose_flow, trigger
from excel_toolkit.validation import build, calculate, verify


def read_job(path):
    """配置和输入路径相对job文件解析，工具包可移机。"""
    path = Path(path).resolve()
    job = json.loads(path.read_text(encoding='utf-8-sig'))
    def resolve(value):
        p = Path(value)
        return str(p if p.is_absolute() else path.parent / p)
    if isinstance(job['config'], str):
        job['config'] = json.loads(Path(resolve(job['config'])).read_text(encoding='utf-8-sig'))
    for item in job['inputs']:
        item['path'] = resolve(item['path'])
    job['output'] = resolve(job['output'])
    if job.get('template'):
        job['template'] = resolve(job['template'])
    return job


def main():
    """提供inspect/preview/build/verify，错误JSON及非零退出码。"""
    parser = argparse.ArgumentParser(description='月度汇总 Excel：原始答卷到可验收工作簿')
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('route', help='识别句式及材料类型，决定处理流程')
    p.add_argument('--text', required=True)
    p.add_argument('--input', action='append', default=[])
    p.add_argument('--template', action='store_true')
    p = sub.add_parser('inspect', help='查看原始数据列名，不输出个人答卷')
    p.add_argument('--input', required=True)
    for command in ['preview', 'build']:
        p = sub.add_parser(command)
        p.add_argument('--job', required=True)
        if command == 'build':
            p.add_argument('--overwrite', action='store_true', help='明确授权备份后覆盖')
    p = sub.add_parser('verify')
    p.add_argument('--file', required=True)
    args = parser.parse_args()
    try:
        if args.command == 'inspect':
            data = load_frame(args.input)
            result = {'status': 'success', 'rows': len(data), 'columns': list(data.columns)}
        elif args.command == 'route':
            if not trigger(args.text):
                result = {'status': 'not_matched', 'flow': None}
            else:
                details = [inspect_material(path) for path in args.input]
                kinds = sorted({kind for detail in details for kind in detail['kinds']})
                result = {'status': 'matched', 'flow': choose_flow(kinds, args.template), 'materials': details}
        elif args.command == 'verify':
            result = verify(args.file)
        elif args.command == 'preview':
            job = read_job(args.job)
            results, sources = calculate(job)
            result = {'status': 'preview', 'inputs': sources, 'results': results}
        else:
            result = build(read_job(args.job), overwrite=args.overwrite)
        print(json.dumps(result, ensure_ascii=False, allow_nan=False, indent=2))
    except (ValueError, OSError, KeyError) as error:
        print(json.dumps({'status': 'error', 'message': str(error)}, ensure_ascii=False))
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
