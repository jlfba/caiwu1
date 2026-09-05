# -*- coding: utf-8 -*-
"""网页版处理逻辑：复用 pdf转图片.py 的函数，封装为可上报进度的任务函数。

通过 import 复用原脚本（不复制逻辑）：把原脚本所在目录加入 sys.path，
后续对控制台版逻辑的修改会自动同步到网页版。
"""
import os
import shutil
import sys

# PDF转图片/ 目录（web/backend 的上一级的上一级的上一级）
_BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

import pdf转图片 as tool  # noqa: E402


def _normalize_chuangshi_car_rows(rows):
    """Normalize current and legacy Chuangshi-car rows to six columns."""
    normalized = []
    for row in rows:
        if len(row) >= 7:
            # Legacy format: invoice, reference, description, description(1),
            # quantity, unit price, amount.
            normalized.append([row[0], row[1], '\n'.join(
                value for value in (row[2], row[3]) if value),
                row[4], row[5], row[6]])
        elif len(row) >= 6:
            normalized.append(list(row[:6]))
    return normalized


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


def process_mode1(pdf_paths, out_dir, progress=None, layout='v', start_cell='A1',
                  template_path=None, sheet_name=''):
    """收款组：PDF 转图片 + OCR 识别重命名 + 生成含图 Excel。

    pdf_paths: 已保存到磁盘的 PDF 绝对路径列表。
    out_dir:   本任务输出目录。
    progress(current, total, message): 进度回调。
    layout:    图片排版方向 'v' 纵向（沿列向下）| 'h' 横向（沿行向右）。
    start_cell: 第一张图的起始单元格（如 A1 / C5）。
    template_path: 可选的已有表格模板（.xlsx/.xlsm），图片插入到其中；不传则自动新建表格。
    sheet_name:   插入到模板的哪个工作表（空则用活动工作表）。
    返回生成的 Excel 绝对路径。
    """
    def report(cur, tot, msg):
        if progress:
            progress(cur, tot, msg)

    if layout not in ('v', 'h'):
        layout = 'v'
    if tool.parse_cell(start_cell) is None:
        raise ValueError('无效的起始格位置：%s' % start_cell)

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

    # ---- 读取文字层 + 渲染 PDF 为 PNG ----
    images = []
    native_fields = []
    seq = 0
    pages_done = 0
    for pdf in pdf_paths:
        try:
            try:
                pdf_fields = tool.extract_invoice_fields_from_pdf(pdf)
            except Exception as e:
                pdf_fields = []
                report(pages_done, total_units,
                       '文字层读取失败，将使用 OCR：%s（%s）'
                       % (os.path.basename(pdf), e))
            imgs, seq = tool.pdf_to_images(
                pdf, img_dir, start_index=seq,
                progress_cb=lambda st, done, tot: report(
                    pages_done + done, total_units,
                    '正在渲染第 %d/%d 页…' % (pages_done + done, total_pages)))
            pages_done += len(imgs)
            images.extend(imgs)
            native_fields.extend(pdf_fields[:len(imgs)])
            if len(pdf_fields) < len(imgs):
                native_fields.extend([None] * (len(imgs) - len(pdf_fields)))
        except Exception as e:
            report(pages_done, total_units,
                   '渲染失败，已跳过：%s（%s）' % (os.path.basename(pdf), e))

    if not images:
        raise RuntimeError('没有成功转换的图片')

    # ---- OCR 识别五字段（开票日期/号码/购买方/销售方/金额）并重命名 ----
    # 文字层完整时直接使用；仅缺字段的页面执行 OCR。
    # 4 核 CPU 限制为 2 个并行任务，避免多个 ONNX 推理会话抢占线程。
    from concurrent.futures import ThreadPoolExecutor, as_completed
    renamed = [None] * len(images)
    done = 0
    needs_ocr = any(not fields or not tool._invoice_fields_complete(fields)
                    for fields in native_fields)
    if needs_ocr:
        tool.get_ocr()  # 仅实际需要 OCR 时加载模型
    max_workers = max(1, min(2, (os.cpu_count() or 2) // 2))
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        future_map = {pool.submit(tool.extract_invoice_fields, img, native_fields[i]): i
                      for i, img in enumerate(images)}
        for fut in as_completed(future_map):
            i = future_map[fut]
            fields = fut.result()
            renamed[i] = tool.rename_with_fields(images[i], fields)
            done += 1
            print('  收款组识别 %d/%d：%s，%.2fs'
                  % (done, len(images), fields.get('_method', 'unknown'),
                     fields.get('_elapsed', 0)))
            report(total_pages + done, total_units,
                   '正在识别发票字段 %d/%d 张…' % (done, len(images)))
    images = renamed

    # ---- 生成含图 Excel ----
    report(total_units, total_units, '正在生成 Excel…')
    if template_path and os.path.isfile(template_path):
        # 插入到用户上传的模板（复制一份到输出目录，不破坏原始上传）
        out = os.path.join(out_dir, os.path.basename(template_path))
        shutil.copy2(template_path, out)
        tool.images_into_excel(out, images, sheet_name=sheet_name or None,
                               start_cell=start_cell.upper(), direction=layout,
                               include_date=False)
    else:
        # 未传模板：自动新建表格
        out = os.path.join(out_dir, '发票图片表.xlsx')
        _blank_workbook(out)
        tool.images_into_excel(out, images,
                               start_cell=start_cell.upper(), direction=layout,
                               include_date=False)
    return out


def process_receipt_mode2(pdf_paths, out_dir, progress=None):
    """收款组模式 2：识别发票五字段并生成简洁 Excel。"""
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    def report(cur, tot, msg):
        if progress:
            progress(cur, tot, msg)

    total = max(len(pdf_paths), 1)
    rows = []
    image_dir = os.path.join(out_dir, 'receipt_images')
    os.makedirs(image_dir, exist_ok=True)
    image_seq = 0

    for index, pdf in enumerate(pdf_paths, 1):
        if not os.path.isfile(pdf):
            report(index, total, '跳过不存在的文件：%s' % os.path.basename(pdf))
            continue
        try:
            fields_list = tool.extract_invoice_fields_from_pdf(pdf)
            # 模式 2 最后一列需要发票原图，因此每一页都渲染；
            # 文字层缺字段时再复用渲染图执行 OCR 兜底。
            images, image_seq = tool.pdf_to_images(pdf, image_dir, start_index=image_seq)
            native_missing = any(not tool._invoice_fields_complete(f) for f in fields_list)
            if native_missing:
                merged = []
                for page_index, image in enumerate(images):
                    initial = fields_list[page_index] if page_index < len(fields_list) else None
                    merged.append(tool.extract_invoice_fields(image, initial))
                fields_list = merged
            for page_index, fields in enumerate(fields_list):
                rows.append([fields.get('date', '未知'),
                             fields.get('seller', '未知'),
                             fields.get('amount', '未知'),
                             fields.get('no', '未知'),
                             fields.get('buyer', '未知'),
                             images[page_index] if page_index < len(images) else None])
        except Exception as exc:
            print('收款组模式 2 识别失败：%s：%s' % (os.path.basename(pdf), exc))
        report(index, total, '正在识别第 %d/%d 个文件：%s' %
               (index, total, os.path.basename(pdf)))

    if not rows:
        raise RuntimeError('没有识别到任何发票信息，未生成 Excel')

    output = os.path.join(out_dir, '收款组发票信息.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = '发票信息'
    headers = ('开票日期', '销售方名称', '金额', '发票号码', '购买方信息', '发票图片')
    ws.append(list(headers))
    for excel_row, row in enumerate(rows, 2):
        ws.append(row[:-1] + [''])
        image_path = row[-1]
        if image_path and os.path.isfile(image_path):
            image = XLImage(image_path)
            ratio = min(320 / image.width, 190 / image.height)
            image.width = int(image.width * ratio)
            image.height = int(image.height * ratio)
            ws.add_image(image, 'F%d' % excel_row)
            ws.row_dimensions[excel_row].height = 150
    for cell in ws[1]:
        cell.font = Font(bold=True)
    widths = (16, 34, 16, 24, 34, 46)
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(output)
    report(total, total, '正在生成 Excel')
    return output


def process_shao_meilin(pdf_paths, out_dir, progress=None):
    """邵梅琳：仅识别中文发票字段，生成不含图片的 Excel。"""
    def report(cur, tot, msg):
        if progress:
            progress(cur, tot, msg)

    total = max(len(pdf_paths), 1)
    rows = []
    image_dir = os.path.join(out_dir, 'ocr_cache')
    image_seq = 0

    for index, pdf in enumerate(pdf_paths, 1):
        if not os.path.isfile(pdf):
            report(index, total, '跳过不存在的文件：%s' % os.path.basename(pdf))
            continue
        try:
            native = tool.extract_invoice_fields_with_summary_from_pdf(pdf)
            # 文字层不完整时，用渲染图补齐字段和摘要。
            needs_ocr = any(not tool._invoice_fields_complete(fields)
                            or fields.get('summary', '未知') == '未知'
                            for fields in native)
            if needs_ocr:
                images, image_seq = tool.pdf_to_images(pdf, image_dir, start_index=image_seq)
                completed = []
                for page_index, image in enumerate(images):
                    initial = native[page_index] if page_index < len(native) else None
                    completed.append(tool.extract_invoice_fields_with_summary(image, initial))
                native = completed
            for fields in native:
                rows.append([fields.get('date', '未知'),
                             fields.get('buyer', '未知'),
                             fields.get('seller', '未知'),
                             fields.get('no', '未知'),
                             fields.get('summary', '未知'),
                             fields.get('amount', '未知')])
        except Exception as exc:
            print('邵梅琳识别失败：%s，%s' % (os.path.basename(pdf), exc))
        report(index, total, '正在识别第 %d/%d 个文件：%s' %
               (index, total, os.path.basename(pdf)))

    if not rows:
        raise RuntimeError('没有识别到任何发票信息，未生成 Excel')

    output = os.path.join(out_dir, '邵梅琳发票识别表.xlsx')
    tool.write_detail_excel(
        rows, output,
        headers=('开票日期', '我方发票抬头', '对方发票抬', '发票号', '摘要', '金额'),
        numeric_cols={5}, text_cols={3}, widths=[16, 34, 34, 22, 52, 16])
    # OCR 图片仅作识别中转，结果目录和下载表格均不保留图片。
    if os.path.isdir(image_dir):
        shutil.rmtree(image_dir)
    report(total, total, '正在生成 Excel')
    return output


def process_mode2(pdf_paths, out_dir, inv_type, progress=None):
    """付款组：发票明细识别 → Excel。

    inv_type: '1' canexs | '2' 精准 | '3' 创时亚马逊卡派 | '4' 创时卡派 | '5' 创时清关费 | '6' 创时附加费 | '7' MAX萨凡纳 | '8' MAX纽约 | '9' AA | '10' JCK | '11' MKK | '12' DINO | '13' EYNEX。
    返回生成的 Excel 绝对路径。
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
        if inv_type == '13':
            rows, pg, sk = tool.extract_eynex_from_pdfs([pdf])
        elif inv_type == '12':
            rows, pg, sk = tool.extract_dino_from_pdfs([pdf])
        elif inv_type == '11':
            rows, pg, sk = tool.extract_mkk_from_pdfs([pdf])
        elif inv_type == '10':
            rows, pg, sk = tool.extract_jck_from_pdfs([pdf])
        elif inv_type == '9':
            rows, pg, sk = tool.extract_aa_from_pdfs([pdf])
        elif inv_type == '8':
            rows, pg, sk = tool.extract_max_ny_from_pdfs([pdf])
        elif inv_type == '7':
            rows, pg, sk = tool.extract_max_portlink_from_pdfs([pdf])
        elif inv_type == '6':
            rows, pg, sk = tool.extract_chuangshi_surcharge_from_pdfs([pdf])
        elif inv_type == '5':
            rows, pg, sk = tool.extract_chuangshi_clearance_from_pdfs([pdf])
        elif inv_type == '4':
            rows, pg, sk = tool.extract_chuangshi_car_from_pdfs([pdf])
        elif inv_type == '3':
            rows, pg, sk = tool.extract_chuangshi_from_pdfs([pdf])
        elif inv_type == '2':
            rows, pg, sk = tool.extract_jingzhun_from_pdfs([pdf])
        else:
            rows, pg, sk = tool.extract_detail_from_pdfs([pdf])
        pages += pg
        skipped += sk
        if inv_type == '4':
            rows = _normalize_chuangshi_car_rows(rows)
        all_rows.extend(rows)
        report(i, n, '正在识别第 %d/%d 个文件：%s' % (i, n, os.path.basename(pdf)))

    if not all_rows:
        raise RuntimeError('没有识别到任何明细，未生成 Excel（共 %d 个文件，%d 页）' % (n, pages))

    if inv_type == '13':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'EYNEX发票明细表.xlsx', tool.EYNEX_OUTPUT_HEADERS,
            {5, 6, 7}, set(), [18, 18, 16, 34, 30, 10, 12, 14])
    elif inv_type == '12':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'DINO发票明细表.xlsx', tool.DINO_OUTPUT_HEADERS,
            {4}, set(), [16, 16, 22, 52, 10, 12, 14, 12])
    elif inv_type == '11':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'MKK发票明细表.xlsx', tool.MKK_OUTPUT_HEADERS,
            {5, 7}, set(), [16, 18, 20, 40, 32, 10, 10, 12])
    elif inv_type == '10':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'JCK发票明细表.xlsx', tool.JCK_OUTPUT_HEADERS,
            {6, 7, 9}, set(), [18, 20, 8, 10, 10, 18, 8, 12, 10, 12, 16])
    elif inv_type == '9':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'AA发票明细表.xlsx', tool.MAX_STYLE_OUTPUT_HEADERS,
            {6}, set(), [36, 22, 16, 16, 10, 36, 8, 12, 12])
    elif inv_type == '8':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'MAX纽约发票明细表.xlsx', tool.MAX_STYLE_OUTPUT_HEADERS,
            {6}, set(), [36, 22, 16, 16, 10, 36, 8, 12, 12])
    elif inv_type == '7':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            'MAX萨凡纳发票明细表.xlsx', tool.MAX_STYLE_OUTPUT_HEADERS,
            {6}, set(), [36, 22, 16, 16, 10, 36, 8, 12, 12])
    elif inv_type == '6':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            '创时附加费发票明细表.xlsx', tool.CHUANGSHI_SURCHARGE_OUTPUT_HEADERS,
            {3, 4, 5}, set(), [16, 20, 40, 12, 14, 16])
    elif inv_type == '5':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            '创时清关费发票明细表.xlsx',
            ('Invoice number', 'Reference', 'Description',
             'Quantity', 'Price', 'Amount'),
            {3, 4, 5}, set(), [16, 20, 34, 12, 14, 16])
    elif inv_type == '4':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            '创时卡派发票明细表.xlsx',
            ('Invoice number', 'Reference', 'Description',
             'Quantity', 'Price', 'Amount'),
            {3, 4, 5}, set(), [16, 20, 46, 12, 14, 16])
    elif inv_type == '3':
        name, headers, numeric_cols, zero_pad_cols, widths = (
            '创时亚马逊卡派发票明细表.xlsx',
            ('Invoice number', 'Reference', 'Description',
             'Quantity', 'Price', 'Amount'),
            {3, 4, 5}, set(), [16, 20, 46, 12, 14, 16])
    elif inv_type == '2':
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
