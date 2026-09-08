# -*- coding: utf-8 -*-
"""网页版处理逻辑：复用 pdf转图片.py 的函数，封装为可上报进度的任务函数。

通过 import 复用原脚本（不复制逻辑）：把原脚本所在目录加入 sys.path，
后续对控制台版逻辑的修改会自动同步到网页版。
"""
import os
import re
import datetime
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
            # 普通电子发票优先走一次 PDF 原生文字层；字段和“*”摘要都能从
            # 坐标文字中取得时，直接完成，不再因为摘要字段未被通用规则识别
            # 而重复转图、启动 OCR。
            native = []
            doc = tool.fitz.open(pdf)
            try:
                for page in doc:
                    items = tool.pdf_native_items(pdf, page)
                    fields = tool._extract_invoice_fields_from_items(items)
                    fields = tool._merge_invoice_fields(fields, _ordinary_native_fields(items))
                    summary = tool._normalize_invoice_summary(
                        tool._extract_invoice_summary_from_items(items))
                    # 这类票的表头可能是“货物或应税劳务、服务名称”，
                    # 通用摘要规则只认“项目名称”，因此用坐标列规则补齐。
                    if summary == '未知':
                        summary = _ordinary_service_from_items(items)
                    fields['summary'] = summary
                    native.append(fields)
            finally:
                doc.close()

            # 原生文字层缺失或字段不完整时，才用 OCR 兜底。
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


def _wechat_items_text(items):
    return ' '.join(str(item.get('text', '')) for item in items if item.get('text'))


def _ordinary_service_from_items(items):
    lines = tool._group_detail_lines(items)
    labels = ("\u8d27\u7269\u6216\u5e94\u7a0e\u52b3\u52a1", "\u670d\u52a1\u540d\u79f0")
    stops = ("\u5408\u8ba1", "\u4ef7\u7a0e\u5408\u8ba1", "\u5907\u6ce8", "\u6536\u6b3e\u4eba", "\u5f00\u7968\u4eba")
    for i, line in enumerate(lines):
        text = re.sub(r"\s+", "", line.get("text", ""))
        if not any(label in text for label in labels):
            continue
        header_items = line.get("items", [])
        anchor = next((item for item in header_items if any(label in item.get("text", "") for label in labels)), None)
        left = anchor.get("cx", 0) - anchor.get("w", 0) / 2 if anchor else 0
        # 右边界取服务名称右侧最近表头的左边缘，避免把后面的规格、数量、金额等列带入。
        right = left + 260
        if anchor:
            right_headers = []
            for candidate in lines:
                if abs(candidate.get("cy", 0) - line.get("cy", 0)) > max(4, anchor.get("h", 10) * 1.5):
                    continue
                for item in candidate.get("items", []):
                    candidate_text = re.sub(r"\s+", "", item.get("text", ""))
                    if item.get("cx", 0) <= anchor.get("cx", 0):
                        continue
                    if candidate_text in ("规格型号", "单位", "数量", "单价", "金额", "税率", "税额"):
                        right_headers.append(item.get("cx", 0) - item.get("w", 0) / 2)
            if right_headers:
                right = min(right_headers)
        values = []
        for candidate in lines[i + 1:]:
            value_text = re.sub(r"\s+", "", candidate.get("text", ""))
            if any(stop in value_text for stop in stops):
                break
            selected = [item.get("text", "").strip() for item in candidate.get("items", []) if left - 8 <= item.get("cx", 0) < right and item.get("text", "").strip()]
            if selected:
                values.append(" ".join(selected))
        if values:
            return "\n".join(values).strip()
    flat = sorted(items, key=lambda item: (item.get("cy", 0), item.get("cx", 0)))
    label_words = ("\u8d27\u7269\u6216\u5e94\u7a0e\u52b3\u52a1", "\u670d\u52a1\u540d\u79f0")
    stop_words = ("\u5408\u8ba1", "\u4ef7\u7a0e\u5408\u8ba1", "\u5907\u6ce8", "\u6536\u6b3e\u4eba", "\u5f00\u7968\u4eba")
    header = next((item for item in flat if any(word in item.get("text", "") for word in label_words)), None)
    if header:
        header_y = header.get("cy", 0)
        later = [item for item in flat if item.get("cy", 0) > header_y + max(3, header.get("h", 10))]
        stop_y = next((item.get("cy", 0) for item in later if any(word in item.get("text", "") for word in stop_words)), float("inf"))
        values = [item.get("text", "").strip() for item in later if item.get("cy", 0) < stop_y and item.get("cx", 0) <= header.get("cx", 0) + max(140, header.get("w", 0) / 2) and item.get("text", "").strip()]
        if values:
            return " \n".join(values).strip()
    return "\u672a\u77e5"


def _ordinary_party_fields(items, fields):
    """普通发票专用的购买方/销售方坐标兜底。

    这类电子发票的“购买方/销售方”常被 PDF 文字层拆成左侧竖排单字，
    而公司名称是右侧独立文字块。按明细表头把页面分成上下两个区域，
    分别在“名称：”所在横带右侧取公司名，避免两栏互相串值。
    """
    if not items:
        return fields
    ordered = sorted(items, key=lambda item: (item.get("cy", 0), item.get("cx", 0)))
    compact = lambda value: re.sub(r"\s+", "", str(value or ""))
    header = next((item for item in ordered
                   if "货物或应税劳务" in compact(item.get("text"))
                   or "服务名称" in compact(item.get("text"))), None)
    header_y = header.get("cy", 0) if header else 180

    def find_name(y_min, y_max):
        name_labels = [item for item in ordered
                       if "名称" in compact(item.get("text"))
                       and y_min <= item.get("cy", 0) <= y_max]
        # PDF 文字层经常把“名称：”拆成“名”和“称：”两个条目；
        # 用视觉行重新合并后再找标签，避免因此漏掉公司名称。
        if not name_labels:
            for line in tool._group_detail_lines(ordered):
                if "名称" in compact(line.get("text", "")) and y_min <= line.get("cy", 0) <= y_max:
                    name_labels = [item for item in line.get("items", [])
                                   if "称" in compact(item.get("text", ""))]
                    if name_labels:
                        break
        if not name_labels:
            return "未知"
        label = min(name_labels, key=lambda item: abs(item.get("cy", 0) -
                                                       (y_min + y_max) / 2))
        candidates = []
        for item in ordered:
            text = str(item.get("text", "")).strip()
            if not text or item is label:
                continue
            if not (y_min <= item.get("cy", 0) <= y_max):
                continue
            if item.get("cx", 0) <= label.get("cx", 0) + 12:
                continue
            if compact(text) in ("纳税人识别号：", "纳税人识别号:", "地址、电话：", "地址、电话:",
                                 "电子支付标识：", "电子支付标识:"):
                continue
            if re.fullmatch(r"[0-9A-Za-z]+", compact(text)):
                continue
            candidates.append(item)
        if not candidates:
            # 标签和值在同一视觉行时，直接取标签右侧第一个非标签文字块。
            for line in tool._group_detail_lines(ordered):
                if abs(line.get("cy", 0) - label.get("cy", 0)) > 14:
                    continue
                right_items = [item for item in line.get("items", [])
                               if item.get("cx", 0) > label.get("cx", 0) + 12
                               and str(item.get("text", "")).strip()]
                right_items = [item for item in right_items
                               if compact(item.get("text")) not in ("名称", "名称：", "名称:")]
                if right_items:
                    return " ".join(item.get("text", "").strip() for item in right_items)
            return "未知"
        # 同一横带优先；公司名通常是该带中最靠左、最长的非标签文字块。
        same_band = [item for item in candidates
                     if abs(item.get("cy", 0) - label.get("cy", 0)) <= 14]
        pool = same_band or candidates
        pool.sort(key=lambda item: (abs(item.get("cy", 0) - label.get("cy", 0)),
                                   item.get("cx", 0), -len(str(item.get("text", "")))))
        return str(pool[0].get("text", "")).strip() or "未知"

    if fields.get("buyer", "未知") == "未知":
        fields["buyer"] = find_name(0, header_y - 8)
    if fields.get("seller", "未知") == "未知":
        fields["seller"] = find_name(header_y + 90, max(item.get("cy", 0) for item in ordered) + 10)
    return fields


def _ordinary_native_fields(items):
    """从普通发票 PDF 文字层直接取字段，避免完整电子票重复跑 OCR。"""
    fields = {key: "未知" for key in ("date", "no", "buyer", "seller", "amount")}
    if not items:
        return fields
    ordered = sorted(items, key=lambda item: (item.get("cy", 0), item.get("cx", 0)))
    compact = lambda value: re.sub(r"\s+", "", str(value or ""))
    lines = tool._group_detail_lines(ordered)

    # 发票号码、开票日期：这类电子票的标签和值通常在同一横带。
    for line in lines:
        text = compact(line.get("text", ""))
        if fields["no"] == "未知" and "发票号码" in text:
            match = re.search(r"发票号码[：:]?([0-9A-Za-z]{6,})", text)
            if match:
                fields["no"] = match.group(1)
        if fields["date"] == "未知" and "开票日期" in text:
            match = re.search(r"开票日期[：:]?(20\d{2})年?(\d{1,2})月?(\d{1,2})日?", text)
            if match:
                fields["date"] = "%04d-%02d-%02d" % tuple(map(int, match.groups()))
            else:
                date_parts = [item.get("text", "") for item in line.get("items", [])]
                match = re.search(r"(20\d{2})年?(\d{1,2})月?(\d{1,2})日?", "".join(date_parts))
                if match:
                    fields["date"] = "%04d-%02d-%02d" % tuple(map(int, match.groups()))

    # 总金额优先取“价税合计”后的大写金额附近的数字，兼容小写金额。
    total_mark = next((line for line in lines if "价税合计" in compact(line.get("text", ""))), None)
    amount_candidates = []
    if total_mark:
        amount_candidates.extend(item.get("text", "") for item in total_mark.get("items", []))
        base_y = total_mark.get("cy", 0)
        amount_candidates.extend(item.get("text", "") for item in ordered
                                 if abs(item.get("cy", 0) - base_y) <= 18)
    amount_candidates.extend(item.get("text", "") for item in ordered if "小写" in item.get("text", ""))
    amount_text = " ".join(amount_candidates)
    amounts = re.findall(r"(?:[¥￥]\s*)?([0-9][0-9,]*\.\d{1,2})", amount_text)
    if amounts:
        fields["amount"] = amounts[-1].replace(",", "")

    fields = _ordinary_party_fields(ordered, fields)
    return fields

def process_shao_ordinary_invoice(pdf_paths, out_dir, progress=None):
    def report(cur, total, message):
        if progress:
            progress(cur, total, message)
    total = max(len(pdf_paths), 1)
    rows = []
    for index, pdf in enumerate(pdf_paths, 1):
        try:
            native = tool.extract_invoice_fields_from_pdf(pdf)
            doc = tool.fitz.open(pdf)
            try:
                for page_index, page in enumerate(doc):
                    items = tool.pdf_native_items(pdf, page)
                    native_fields = native[page_index] if page_index < len(native) else {}
                    fields = tool._merge_invoice_fields(native_fields, _ordinary_native_fields(items))
                    summary_items = items
                    if not items or not tool._invoice_fields_complete(fields):
                        cache_dir = os.path.join(out_dir, "ordinary_cache")
                        os.makedirs(cache_dir, exist_ok=True)
                        image_path = os.path.join(cache_dir, "page_%05d.png" % len(rows))
                        pix = page.get_pixmap(matrix=tool.fitz.Matrix(tool.RENDER_DPI / 72.0, tool.RENDER_DPI / 72.0), alpha=False)
                        pix.save(image_path)
                        ocr_items = tool.ocr_lines(image_path)
                        fields = tool._merge_invoice_fields(fields, tool._extract_invoice_fields_from_items(ocr_items, fields))
                        if not summary_items:
                            summary_items = ocr_items
                    fields = _ordinary_party_fields(items, fields)
                    rows.append([fields.get("no", "\u672a\u77e5"), fields.get("date", "\u672a\u77e5"), fields.get("buyer", "\u672a\u77e5"), fields.get("seller", "\u672a\u77e5"), _ordinary_service_from_items(summary_items), fields.get("amount", "\u672a\u77e5")])
            finally:
                doc.close()
        except Exception as exc:
            print("ordinary invoice recognition failed: %s: %s" % (os.path.basename(pdf), exc))
        report(index, total, "ordinary invoice %d/%d: %s" % (index, total, os.path.basename(pdf)))
    if not rows:
        raise RuntimeError("\u6ca1\u6709\u8bc6\u522b\u5230\u666e\u901a\u53d1\u7968\u4fe1\u606f")
    output = os.path.join(out_dir, "\u9093\u6885\u7433\u666e\u901a\u53d1\u7968\u8bc6\u522b\u8868.xlsx")
    tool.write_detail_excel(rows, output, headers=("\u53d1\u7968\u53f7\u7801", "\u5f00\u7968\u65e5\u671f", "\u8d2d\u4e70\u65b9\u540d\u79f0", "\u9500\u552e\u65b9\u540d\u79f0", "\u8d27\u7269\u6216\u5e94\u7a0e\u52b3\u52a1\u3001\u670d\u52a1\u540d\u79f0", "\u603b\u91d1\u989d"), numeric_cols={5}, text_cols={0}, widths=[22, 16, 34, 34, 52, 16])
    return output


def _wechat_date_time(text):
    compact = re.sub(r'\s+', '', text or '')
    date_pattern = r'(20\d{2})[\u5e74./-](\d{1,2})[\u6708./-](\d{1,2})\u65e5?[^0-9]{0,8}(\d{1,2})[:\uFF1A](\d{2})[:\uFF1A](\d{2})'
    compact_pattern = r'(20\d{2})(\d{2})(\d{2})[^0-9]{0,8}(\d{1,2})[:\uFF1A](\d{2})[:\uFF1A](\d{2})'
    patterns = (
        date_pattern,
        compact_pattern,
        r'(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})日?[^0-9]{0,8}(\d{1,2})[:：](\d{2})[:：](\d{2})',
        r'(20\d{2})(\d{2})(\d{2})[^0-9]{0,8}(\d{1,2})[:：](\d{2})[:：](\d{2})',
    )
    for pattern in patterns:
        match = re.search(pattern, compact)
        if match:
            year, month, day, hour, minute, second = map(int, match.groups())
            try:
                date_value = datetime.date(year, month, day).isoformat()
                return date_value, "%02d:%02d:%02d" % (hour, minute, second)
            except ValueError:
                pass
    return '未知', '未知'


def _wechat_amount(text, items=None):
    source = text or ''
    if items:
        top = [item for item in items if item.get('cy', 999999) <= max(240, max(x.get('cy', 0) for x in items) * 0.35)]
        source = _wechat_items_text(top) + ' ' + source
    match = re.search(r'[-\uFF0D\u2013\u2014]\s*(\d[\d,]*(?:\.\d{1,2})?)', source)
    if not match:
        return '未知'
    return '-' + match.group(1).replace(',', '')


def _wechat_extract(items):
    text = _wechat_items_text(items)
    date_value, time_value = _wechat_date_time(text)
    transfer_label = '\u8f6c\u8d26\u65f6\u95f4'
    label = next((item for item in items if transfer_label in item.get('text', '')), None)
    if label:
        nearby = [item for item in items if item.get('cx', 0) > label.get('cx', 0) and abs(item.get('cy', 0) - label.get('cy', 0)) < max(30, label.get('h', 12) * 2.5)]
        near_date, near_time = _wechat_date_time(_wechat_items_text(nearby))
        if near_date != '未知':
            date_value, time_value = near_date, near_time
    return date_value, time_value, _wechat_amount(text, items)


def _payment_extract(items, payment_type):
    text = _wechat_items_text(items)
    date_value, time_value = _wechat_date_time(text)
    date_labels = {'alipay': '\u652f\u4ed8\u65f6\u95f4', 'huolala': '\u652f\u4ed8\u65f6\u95f4', 'wechat': '\u8f6c\u8d26\u65f6\u95f4'}
    anchor_text = date_labels.get(payment_type, '\u652f\u4ed8\u65f6\u95f4')
    label = next((item for item in items if anchor_text in item.get('text', '')), None)
    if label:
        nearby = [item for item in items if item.get('cx', 0) > label.get('cx', 0) and abs(item.get('cy', 0) - label.get('cy', 0)) < max(30, label.get('h', 12) * 2.5)]
        near_date, near_time = _wechat_date_time(_wechat_items_text(nearby))
        if near_date != '未知':
            date_value, time_value = near_date, near_time
    amount_anchor = {'alipay': '\u4ea4\u6613\u6210\u529f', 'huolala': '\u8d27\u62c9\u62c9', 'wechat': ''}.get(payment_type, '')
    anchor = next((item for item in items if amount_anchor and amount_anchor in item.get('text', '')), None)
    candidates = []
    for item in items:
        if not re.search(r'[-\uFF0D\u2013\u2014]\s*\d', item.get('text', '')):
            continue
        if anchor and item.get('cy', 0) >= anchor.get('cy', 0):
            continue
        distance = abs(item.get('cy', 0) - anchor.get('cy', 0)) if anchor else item.get('cy', 0)
        candidates.append((distance, item))
    if candidates:
        amount = _wechat_amount(min(candidates, key=lambda pair: pair[0])[1].get('text', ''))
    else:
        amount = _wechat_amount(text, items)
    return date_value, time_value, amount.lstrip('-\uFF0D\u2013\u2014')


def _wechat_pdf_page_items(page):
    words = page.get_text('words') or []
    return [{'text': word[4].strip(), 'cx': (word[0] + word[2]) / 2,
             'cy': (word[1] + word[3]) / 2, 'w': word[2] - word[0],
             'h': word[3] - word[1]} for word in words if word[4].strip()]


def process_payment_receipts(file_paths, out_dir, progress=None, payment_type='wechat'):
    """赵淑华微信转账：每张凭证一行，最后一列嵌入原凭证。"""
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    def report(cur, total, message):
        if progress:
            progress(cur, total, message)

    image_dir = os.path.join(out_dir, 'wechat_images')
    os.makedirs(image_dir, exist_ok=True)
    rows = []
    total = max(len(file_paths), 1)
    image_seq = 0
    for file_index, path in enumerate(file_paths, 1):
        if not os.path.isfile(path):
            continue
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == '.pdf':
                doc = tool.fitz.open(path)
                try:
                    for page_index, page in enumerate(doc):
                        image_path = os.path.join(image_dir, "wechat_%05d.png" % image_seq)
                        image_seq += 1
                        pix = page.get_pixmap(matrix=tool.fitz.Matrix(2, 2), alpha=False)
                        pix.save(image_path)
                        items = _wechat_pdf_page_items(page)
                        if not items:
                            items = tool.ocr_lines(image_path)
                        date_value, time_value, amount = _payment_extract(items, payment_type)
                        rows.append([date_value, time_value, amount, image_path])
                finally:
                    doc.close()
            else:
                image_path = os.path.join(image_dir, "wechat_%05d%s" % (image_seq, ext))
                image_seq += 1
                shutil.copy2(path, image_path)
                rows.append([*_payment_extract(tool.ocr_lines(image_path), payment_type), image_path])
        except Exception as exc:
            print('微信转账识别失败：%s：%s' % (os.path.basename(path), exc))
        report(file_index, total, '正在识别微信凭证 %d/%d：%s' % (file_index, total, os.path.basename(path)))

    if not rows:
        raise RuntimeError('没有识别到微信转账凭证')
    output = os.path.join(out_dir, {'wechat': '赵淑华微信转账.xlsx', 'alipay': '赵淑华支付宝支付.xlsx', 'huolala': '赵淑华货拉拉支付.xlsx'}.get(payment_type, '赵淑华支付凭证.xlsx'))
    wb = Workbook()
    ws = wb.active
    ws.title = '微信转账'
    ws.append(['转账时间年', '转账时间', '金额', '附图'])
    for row_index, row in enumerate(rows, 2):
        ws.append(row[:3] + [''])
        image = XLImage(row[3])
        ratio = min(320 / image.width, 190 / image.height)
        image.width = int(image.width * ratio)
        image.height = int(image.height * ratio)
        ws.add_image(image, 'D%d' % row_index)
        ws.row_dimensions[row_index].height = 150
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for index, width in enumerate((16, 14, 16, 46), 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(output)
    report(total, total, '微信转账 Excel 已生成')
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
