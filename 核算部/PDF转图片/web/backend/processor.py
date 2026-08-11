# -*- coding: utf-8 -*-
"""网页版处理逻辑：复用 pdf转图片.py 的函数，封装为可上报进度的任务函数。

通过 import 复用原脚本（不复制逻辑）：把原脚本所在目录加入 sys.path，
后续对控制台版逻辑的修改会自动同步到网页版。
"""
import os
import sys

# PDF转图片/ 目录（web/backend 的上一级的上一级的上一级）
_BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

import pdf转图片 as tool  # noqa: E402


def sanitize_filename(name):
    """清洗上传文件名，返回安全的保存名（仅文件名，不含路径）。"""
    name = os.path.basename(name or '')
    name = tool.sanitize(name)
    return name or 'file.pdf'


def _blank_workbook(path):
    """新建一个空白 xlsx，供 images_into_excel 装载图片。"""
    from openpyxl import Workbook
    wb = Workbook()
    wb.save(path)


def process_mode1(pdf_paths, out_dir, progress=None):
    """收款组：PDF 转图片 + OCR 识别重命名 + 生成含图 Excel。

    pdf_paths: 已保存到磁盘的 PDF 绝对路径列表。
    out_dir:   本任务输出目录。
    progress(current, total, message): 进度回调。
    返回生成的 Excel 绝对路径。
    """
    def report(cur, tot, msg):
        if progress:
            progress(cur, tot, msg)

    img_dir = os.path.join(out_dir, 'images')
    os.makedirs(img_dir, exist_ok=True)

    # 统计总页数（渲染一页 + 识别一张，整体进度按 页数*2 计算）
    total_pages = 0
    for p in pdf_paths:
        try:
            doc = tool.fitz.open(p)
            total_pages += doc.page_count
            doc.close()
        except Exception:
            pass
    total_units = max(total_pages * 2, 1)

    # ---- 渲染 PDF 为 PNG ----
    images = []
    seq = 0
    pages_done = 0
    for pdf in pdf_paths:
        try:
            imgs, seq = tool.pdf_to_images(
                pdf, img_dir, start_index=seq,
                progress_cb=lambda st, done, tot: report(
                    pages_done + done, total_units,
                    '正在渲染第 %d/%d 页…' % (pages_done + done, total_pages)))
            pages_done += len(imgs)
            images.extend(imgs)
        except Exception as e:
            report(pages_done, total_units,
                   '渲染失败，已跳过：%s（%s）' % (os.path.basename(pdf), e))

    if not images:
        raise RuntimeError('没有成功转换的图片')

    # ---- OCR 识别四字段并重命名 ----
    renamed = []
    for i, img in enumerate(images, 1):
        fields = tool.extract_invoice_fields(img)
        new = tool.rename_with_fields(img, fields)
        renamed.append(new)
        report(total_pages + i, total_units,
               '正在识别发票字段 %d/%d 张…' % (i, len(images)))
    images = renamed

    # ---- 生成含图 Excel ----
    report(total_units, total_units, '正在生成 Excel…')
    xlsx_path = os.path.join(out_dir, '发票图片表.xlsx')
    _blank_workbook(xlsx_path)
    tool.images_into_excel(xlsx_path, images, direction='v')
    return xlsx_path


def process_mode2(pdf_paths, out_dir, inv_type, progress=None):
    """付款组：发票明细识别 → Excel。

    inv_type: '1' canexs | '2' 精准。返回生成的 Excel 绝对路径。
    """
    def report(cur, tot, msg):
        if progress:
            progress(cur, tot, msg)

    n = max(len(pdf_paths), 1)
    all_rows, pages, skipped = [], 0, 0
    for i, pdf in enumerate(pdf_paths, 1):
        if not os.path.isfile(pdf):
            skipped += 1
            continue
        if inv_type == '2':
            rows, pg, sk = tool.extract_jingzhun_from_pdfs([pdf])
        else:
            rows, pg, sk = tool.extract_detail_from_pdfs([pdf])
        pages += pg
        skipped += sk
        all_rows.extend(rows)
        report(i, n, '正在识别第 %d/%d 个文件：%s' % (i, n, os.path.basename(pdf)))

    if not all_rows:
        raise RuntimeError('没有识别到任何明细，未生成 Excel（共 %d 个文件，%d 页）' % (n, pages))

    if inv_type == '2':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            '精准发票明细表.xlsx', tool.JINGZHUN_OUTPUT_HEADERS,
            {4}, {0}, [16, 26, 18, 50, 16])
    else:
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'canexs发票明细表.xlsx', None, None, None, None)

    report(i, n, '正在生成 Excel…')
    xlsx_path = os.path.join(out_dir, name)
    tool.write_detail_excel(all_rows, xlsx_path,
                            headers=headers, numeric_cols=numeric_cols,
                            zero_pad_cols=zero_pad_cols, widths=widths)
    return xlsx_path
