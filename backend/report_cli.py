# -*- coding: utf-8 -*-
"""报表组命令行处理入口，不依赖浏览器。"""
import argparse
import os
import sys
import re

from openpyxl import load_workbook

try:
    from . import report_processor
except ImportError:
    import report_processor


def _progress(current, total, message):
    print(f"[进度] {message}", flush=True)


def _clean_path(value):
    value = value.strip()
    # 支持直接拖入文件，也支持误粘贴 PowerShell 的 & '路径' / & "路径"。
    if value.startswith('&'):
        value = value[1:].strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
        value = value[1:-1]
    return value.strip()


def _choose_sheet(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = wb.sheetnames
    finally:
        wb.close()
    if not names:
        raise ValueError('工作簿中没有工作表')
    print('检测到工作表：')
    for index, name in enumerate(names, 1):
        print(f"  {index}. {name}")
    while True:
        value = input('请输入要处理的工作表序号（直接输入名称也可以）：').strip()
        if value.isdigit() and 1 <= int(value) <= len(names):
            return names[int(value) - 1]
        if value in names:
            return value
        print('输入无效，请重新选择。')


def main():
    parser = argparse.ArgumentParser(description='报表组分步处理（终端版）')
    parser.add_argument('input', nargs='?', help='输入 Excel 文件路径')
    parser.add_argument('-s', '--sheet', help='工作表名称；不填则交互选择')
    parser.add_argument('-o', '--output-dir', help='输出目录；默认与输入文件同目录')
    parser.add_argument('--no-pause', action='store_true', help='不在每一步等待回车')
    args = parser.parse_args()

    path = _clean_path(args.input or input('请输入 Excel 文件路径：'))
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        print(f'文件不存在：{path}', file=sys.stderr)
        return 2
    sheet = args.sheet or _choose_sheet(path)
    output_dir = os.path.abspath(args.output_dir or os.path.dirname(path))
    os.makedirs(output_dir, exist_ok=True)

    current = path
    total_steps = len(report_processor.STEP_NAMES)
    print(f'开始处理：{path}')
    print(f'工作表：{sheet}')
    print(f'共 {total_steps} 步；每步完成后会暂停。')
    for step in range(1, total_steps + 1):
        output = os.path.join(output_dir, f'无应收明细-步骤{step}.xlsx')
        print(f'\n===== 步骤 {step}/{total_steps}：{report_processor.STEP_NAMES[step - 1]} =====')
        try:
            report_processor.process_report_step(current, sheet, output, step, _progress)
        except Exception as exc:
            print(f'处理失败：{exc}', file=sys.stderr)
            return 1
        print(f'本步骤完成，结果文件：{output}')
        if step < total_steps and not args.no_pause:
            answer = input('按回车继续下一步，输入 q 退出：').strip().lower()
            if answer == 'q':
                print('已暂停。下次可从该步骤结果文件继续处理。')
                return 0
        current = output
    print(f'全部处理完成：{current}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
