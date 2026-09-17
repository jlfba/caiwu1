# -*- coding: utf-8 -*-
"""资金组独立处理入口：不启动网页，直接调用本机 Excel 完成核对。"""
import os
import sys
import traceback
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'web' / 'backend'))

import fund_processor  # noqa: E402


def _progress(current, total, message):
    print(f'[{current}/{total}] {message}', flush=True)


def main():
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        selected = filedialog.askopenfilenames(
            title='选择资金组 Excel 文件（可一次选择全部）',
            initialdir=str(ROOT),
            filetypes=[('Excel 文件', '*.xlsx *.xlsm')],
        )
        if not selected:
            return

        paths = list(selected)
        file1, file2, file3, file4 = fund_processor.identify_fund_files(paths)
        date_suffix = fund_processor._extract_date(next(
            path for path in (file1, file2, file4, file3) if path
        ))
        output_dir = os.path.join(os.path.dirname(file3), f'资金核对-{date_suffix}')
        os.makedirs(output_dir, exist_ok=True)

        print('开始处理资金组文件：', flush=True)
        for path in paths:
            print(' - ' + os.path.basename(path), flush=True)
        print('结果目录：' + output_dir, flush=True)
        result = fund_processor.process_fund(
            file1, file2, file3, file4, output_dir, _progress
        )
        messagebox.showinfo('资金组核对完成', f'已生成结果：\n{result}')
        os.startfile(os.path.dirname(result))
    except Exception as exc:
        traceback.print_exc()
        messagebox.showerror('资金组核对失败', str(exc))
    finally:
        root.destroy()


if __name__ == '__main__':
    main()
