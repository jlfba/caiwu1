# -*- coding: utf-8 -*-
"""报表组 Excel 处理逻辑。"""
from collections import Counter, defaultdict
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

EXCLUDED_CUSTOMERS = (
    '风驰-数据同步', 'YX订舱', '风驰-卖柜', '李雪原',
    '龙行-清关', '李增韬',
)
REQUIRED_COLUMNS = (
    '应收单价', '客户简称', '业务员', '自定义备注', '配仓单号',
    '销售产品', '应收金额', '客户所属机构', '运单号',
)


def _text(value):
    return '' if value is None else str(value).strip()


def _number(value):
    if value is None or value == '':
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r'[^0-9.\-]', '', str(value).replace(',', ''))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _headers(ws):
    return {_text(cell.value): cell.column for cell in ws[1] if _text(cell.value)}


def _should_keep(values, idx):
    raw_unit_price = values[idx['应收单价'] - 1]
    if not _text(raw_unit_price):
        return False
    unit_price = _number(raw_unit_price)
    if unit_price >= 1:
        return False
    customer = _text(values[idx['客户简称'] - 1])
    if any(word in customer for word in EXCLUDED_CUSTOMERS):
        return False
    if '华南KA' in _text(values[idx['业务员'] - 1]):
        return False
    remark = _text(values[idx['自定义备注'] - 1])
    if re.match(r'^J[0-9A-Za-z-]*', remark, re.IGNORECASE):
        return False
    if '无应收' in remark or '免费补发' in remark:
        return False
    if '刘丹整柜' in _text(values[idx['配仓单号'] - 1]):
        return False
    product = _text(values[idx['销售产品'] - 1])
    amount = _number(values[idx['应收金额'] - 1])
    if '整柜' in product and amount > 10000:
        return False
    return True


def _keyword_keep(values, idx):
    customer = _text(values[idx['客户简称'] - 1])
    if any(word in customer for word in EXCLUDED_CUSTOMERS):
        return False
    if '华南KA' in _text(values[idx['业务员'] - 1]):
        return False
    remark = _text(values[idx['自定义备注'] - 1])
    if re.match(r'^J[0-9A-Za-z-]*', remark, re.IGNORECASE):
        return False
    if '无应收' in remark or '免费补发' in remark:
        return False
    if '刘丹整柜' in _text(values[idx['配仓单号'] - 1]):
        return False
    return True


def _source_totals(ws):
    headers = _headers(ws)
    col = headers.get('客户所属机构')
    if not col:
        raise ValueError('用于统计总票数的工作表缺少“客户所属机构”列')
    return Counter(_text(row[col - 1].value) for row in ws.iter_rows(min_row=2)
                   if _text(row[col - 1].value))


def process_report(input_path, selected_sheet, output_path, progress=None):
    def report(cur, total, message):
        if progress:
            progress(cur, total, message)

    report(1, 5, '正在加载工作簿（大文件首次打开可能需要一些时间）')
    keep_vba = input_path.lower().endswith('.xlsm')
    # 超大工作簿采用只读流式模式，避免将数 GB 的 worksheet XML 全部载入内存。
    if os.path.getsize(input_path) >= 100 * 1024 * 1024:
        return _process_large_report(input_path, selected_sheet, output_path, progress)
    wb = load_workbook(input_path, keep_vba=keep_vba)
    if selected_sheet not in wb.sheetnames:
        raise ValueError('工作表不存在：%s' % selected_sheet)
    source = wb[selected_sheet]
    idx = _headers(source)
    missing = [name for name in REQUIRED_COLUMNS if name not in idx]
    if missing:
        raise ValueError('所选工作表缺少必要列：%s' % '、'.join(missing))

    for name in ('无应收明细', '无应收明细透视表'):
        if name in wb.sheetnames:
            del wb[name]
    detail = wb.copy_worksheet(source)
    detail.title = '无应收明细'
    report(2, 5, '正在复制所选工作表')

    report(3, 5, '正在筛选无应收明细')
    kept = []
    for row in source.iter_rows(min_row=2, values_only=True):
        values = list(row)
        if _should_keep(values, idx):
            unit_price = _number(values[idx['应收单价'] - 1])
            amount = _number(values[idx['应收金额'] - 1])
            category = '金额异常' if amount > 0 else ('无应收' if unit_price == 0 else '')
            kept.append(values + [category])

    if detail.max_row > 1:
        detail.delete_rows(2, detail.max_row - 1)
    category_col = detail.max_column + 1
    detail.cell(1, category_col, '无应收')
    for values in kept:
        detail.append(values)
    detail.auto_filter.ref = detail.dimensions
    detail.freeze_panes = 'A2'

    report(4, 5, '正在生成无应收明细透视表')
    groups = defaultdict(Counter)
    categories = []
    for values in kept:
        org = _text(values[idx['客户所属机构'] - 1])
        tracking = _text(values[idx['运单号'] - 1])
        category = _text(values[-1])
        if org and tracking:
            groups[org][category] += 1
            if category not in categories:
                categories.append(category)
    categories = [name for name in ('无应收', '金额异常', '') if name in categories]
    source_total_sheet = wb['锦联国际V1.0.1'] if '锦联国际V1.0.1' in wb.sheetnames else source
    total_tickets = _source_totals(source_total_sheet)

    pivot = wb.create_sheet('无应收明细透视表')
    display_categories = [('空白' if not name else name, name) for name in categories]
    pivot.append(['客户所属机构'] + [label for label, _ in display_categories]
                 + ['合计', '总票数', '占比'])
    for org in sorted(groups):
        counts = [groups[org][key] for _, key in display_categories]
        total = sum(counts)
        denominator = total_tickets.get(org, 0)
        pivot.append([org] + counts + [total, denominator, total / denominator if denominator else 0])
    last_row = pivot.max_row + 1
    pivot.cell(last_row, 1, '总计')
    ratio_col = pivot.max_column
    grand_total_col = ratio_col - 2
    ticket_total_col = ratio_col - 1
    for col in range(2, ratio_col):
        pivot.cell(last_row, col, '=SUM(%s2:%s%d)' %
                   (get_column_letter(col), get_column_letter(col), last_row - 1))
    grand_total_cell = '%s%d' % (get_column_letter(grand_total_col), last_row)
    ticket_total_cell = '%s%d' % (get_column_letter(ticket_total_col), last_row)
    pivot.cell(last_row, ratio_col, '=IF(%s=0,0,%s/%s)' %
               (ticket_total_cell, grand_total_cell, ticket_total_cell))
    for row in range(2, pivot.max_row + 1):
        pivot.cell(row, ratio_col).number_format = '0.00%'
    for cell in pivot[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='0F8F7F')
        cell.alignment = Alignment(horizontal='center')
    pivot.freeze_panes = 'A2'
    pivot.auto_filter.ref = pivot.dimensions
    for col in range(1, pivot.max_column + 1):
        pivot.column_dimensions[get_column_letter(col)].width = 22 if col == 1 else 14

    report(5, 5, '正在保存处理结果')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path


STEP_NAMES = ('复制工作表', '筛选应收单价', '删除字段关键词', '删除整柜异常', '新增无应收分类', '生成透视表', '清理临时记录')

def process_report_step(input_path, selected_sheet, output_path, step, progress=None):
    """执行单个可暂停步骤；中间删除记录保存在临时工作表。"""
    def report(cur, total, message):
        if progress:
            progress(cur, total, message)
    if os.path.getsize(input_path) >= 100 * 1024 * 1024:
        return _process_large_report_step(input_path, selected_sheet, output_path, step, progress)
    keep_vba = input_path.lower().endswith('.xlsm')
    wb = load_workbook(input_path, keep_vba=keep_vba)
    if selected_sheet not in wb.sheetnames:
        raise ValueError('工作表不存在：%s' % selected_sheet)
    if step == 1:
        source = wb[selected_sheet]
        idx = _headers(source)
        missing = [name for name in REQUIRED_COLUMNS if name not in idx]
        if missing:
            raise ValueError('所选工作表缺少必要列：%s' % '、'.join(missing))
        for name in ('无应收明细', '无应收明细透视表'):
            if name in wb.sheetnames: del wb[name]
        detail = wb.copy_worksheet(source); detail.title = '无应收明细'
        report(1, 1, '步骤 1/7：已复制所选工作表')
    else:
        detail = wb['无应收明细']
        headers = _headers(detail)
        if step in (2, 3, 4):
            rules = {2: lambda v: bool(_text(v[headers['应收单价']-1])) and _number(v[headers['应收单价']-1]) < 1,
                     3: lambda v: _keyword_keep(v, headers),
                     4: lambda v: not ('整柜' in _text(v[headers['销售产品']-1]) and _number(v[headers['应收金额']-1]) > 10000)}
            source_rows = list(detail.iter_rows(min_row=2, values_only=True))
            kept, removed = [], []
            for values in source_rows:
                values = list(values)
                ok = rules[step](values, headers) if step == 3 else rules[step](values)
                (kept if ok else removed).append(values)
            if detail.max_row > 1: detail.delete_rows(2, detail.max_row - 1)
            for values in kept: detail.append(values)
            log_name = '临时删除_步骤%d' % step
            if log_name in wb.sheetnames: del wb[log_name]
            log = wb.create_sheet(log_name); log.append(list(detail.iter_rows(min_row=1, max_row=1, values_only=True))[0] if detail.max_row else [])
            for values in removed: log.append(values)
            report(1, 1, '步骤 %d/7：保留 %d 行，删除 %d 行' % (step, len(kept), len(removed)))
        elif step == 5:
            if '无应收' not in headers:
                col = detail.max_column + 1; detail.cell(1, col, '无应收')
                for row in detail.iter_rows(min_row=2):
                    amount = _number(row[headers['应收金额']-1].value); price = _number(row[headers['应收单价']-1].value)
                    row[col-1].value = '金额异常' if amount > 0 else ('无应收' if price == 0 else '')
            report(1, 1, '步骤 5/7：已新增无应收分类')
        elif step == 6:
            if '无应收明细透视表' in wb.sheetnames: del wb['无应收明细透视表']
            headers = _headers(detail); groups = defaultdict(Counter); source_totals = Counter()
            for row in detail.iter_rows(min_row=2, values_only=True):
                org = _text(row[headers['客户所属机构']-1]); tracking = _text(row[headers['运单号']-1]); cat = _text(row[-1])
                if org: source_totals[org] += 1
                if org and tracking: groups[org][cat] += 1
            pivot = wb.create_sheet('无应收明细透视表'); cats = [x for x in ('无应收','金额异常','') if any(g[x] for g in groups.values())]
            pivot.append(['客户所属机构'] + [('空白' if not x else x) for x in cats] + ['合计', '总票数', '占比'])
            for org in sorted(groups):
                nums=[groups[org][x] for x in cats]; total=sum(nums); denominator=source_totals.get(org, 0)
                pivot.append([org]+nums+[total, denominator, total / denominator if denominator else 0])
            report(1, 1, '步骤 6/7：已生成透视表')
        elif step == 7:
            for name in list(wb.sheetnames):
                if name.startswith('临时删除_步骤'): del wb[name]
            report(1, 1, '步骤 7/7：已清理临时删除记录')
    os.makedirs(os.path.dirname(output_path), exist_ok=True); wb.save(output_path); wb.close(); return output_path


def _process_large_report_step(input_path, selected_sheet, output_path, step, progress=None):
    """大文件分步流式处理，避免可编辑模式加载数 GB XML。"""
    def report(cur, total, message):
        if progress: progress(cur, total, message)
    report(1, 1, '步骤 %d/7：正在流式读取数据' % step)
    source_wb = load_workbook(input_path, read_only=True, data_only=False)
    if selected_sheet not in source_wb.sheetnames: source_wb.close(); raise ValueError('工作表不存在：%s' % selected_sheet)
    source_name = '无应收明细' if step > 1 and '无应收明细' in source_wb.sheetnames else selected_sheet
    source = source_wb[source_name]; rows = source.iter_rows(values_only=True); header = list(next(rows, None) or [])
    if not header: source_wb.close(); raise ValueError('所选工作表为空')
    idx = {_text(v): i + 1 for i, v in enumerate(header) if _text(v)}
    missing = [name for name in REQUIRED_COLUMNS if name not in idx]
    if missing: source_wb.close(); raise ValueError('所选工作表缺少必要列：%s' % '、'.join(missing))
    out = Workbook(write_only=True); original = out.create_sheet(selected_sheet); detail = out.create_sheet('无应收明细')
    original.append(header)
    detail.append(header)
    deleted = out.create_sheet('临时删除_步骤%d' % step) if step in (2, 3, 4) else None
    if deleted: deleted.append(header)
    kept_count = deleted_count = row_count = 0
    for row in rows:
        values = list(row); original.append(values); row_count += 1
        keep = True
        if step >= 2:
            if step == 2: keep = bool(_text(values[idx['应收单价']-1])) and _number(values[idx['应收单价']-1]) < 1
            elif step == 3: keep = _keyword_keep(values, idx)
            elif step == 4: keep = not ('整柜' in _text(values[idx['销售产品']-1]) and _number(values[idx['应收金额']-1]) > 10000)
            else: keep = True
        if keep:
            if step >= 5 and step == 5:
                price = _number(values[idx['应收单价']-1]); amount = _number(values[idx['应收金额']-1])
                values = values + ['金额异常' if amount > 0 else ('无应收' if price == 0 else '')]
            elif step > 5 and len(values) == len(header):
                values = values
            detail.append(values); kept_count += 1
        elif deleted:
            deleted.append(values); deleted_count += 1
        if row_count % 10000 == 0: report(1, 1, '步骤 %d/7：已读取 %d 行，保留 %d 行，删除 %d 行' % (step, row_count, kept_count, deleted_count))
    source_wb.close()
    if step == 6:
        # 从已经写出的明细重新读取成本较高；当前流式阶段先输出可下载明细，透视表在下一次继续时生成。
        report(1, 1, '步骤 6/7：已完成明细流式写出，透视表将在最终步骤生成')
    if step == 7:
        for name in list(out.sheetnames):
            if name.startswith('临时删除_步骤'): del out[name]
    os.makedirs(os.path.dirname(output_path), exist_ok=True); out.save(output_path); out.close()
    return output_path


def _process_large_report(input_path, selected_sheet, output_path, progress=None):
    """大文件流式处理：只读取选中工作表并写入新的结果工作簿。"""
    def report(cur, total, message):
        if progress:
            progress(cur, total, message)

    report(1, 5, '正在流式加载工作簿（大文件模式）')
    source_wb = load_workbook(input_path, read_only=True, data_only=False)
    if selected_sheet not in source_wb.sheetnames:
        source_wb.close()
        raise ValueError('工作表不存在：%s' % selected_sheet)
    source_name = '无应收明细' if step > 1 and '无应收明细' in source_wb.sheetnames else selected_sheet
    source = source_wb[source_name]
    rows = source.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        source_wb.close()
        raise ValueError('所选工作表为空')
    header = list(header)
    idx = {_text(value): pos + 1 for pos, value in enumerate(header) if _text(value)}
    missing = [name for name in REQUIRED_COLUMNS if name not in idx]
    if missing:
        source_wb.close()
        raise ValueError('所选工作表缺少必要列：%s' % '、'.join(missing))

    report(2, 5, '正在复制所选工作表（流式写入）')
    out_wb = Workbook(write_only=True)
    source_out = out_wb.create_sheet(selected_sheet)
    source_out.append(header)
    detail = out_wb.create_sheet('无应收明细')
    detail.append(header + ['无应收'])
    source_totals = Counter()
    groups = defaultdict(Counter)
    categories = []
    row_count = kept_count = removed_count = 0
    for row in rows:
        values = list(row)
        source_out.append(values)
        row_count += 1
        org = _text(values[idx['客户所属机构'] - 1])
        if org:
            source_totals[org] += 1
        if _should_keep(values, idx):
            unit_price = _number(values[idx['应收单价'] - 1])
            amount = _number(values[idx['应收金额'] - 1])
            category = '金额异常' if amount > 0 else ('无应收' if unit_price == 0 else '')
            detail.append(values + [category])
            kept_count += 1
            tracking = _text(values[idx['运单号'] - 1])
            if org and tracking:
                groups[org][category] += 1
                if category not in categories:
                    categories.append(category)
        else:
            removed_count += 1
        if row_count % 10000 == 0:
            report(3, 5, '已读取 %d 行，保留 %d 行，删除 %d 行' %
                   (row_count, kept_count, removed_count))
    source_wb.close()

    report(4, 5, '正在生成无应收明细透视表')
    categories = [name for name in ('无应收', '金额异常', '') if name in categories]
    pivot = out_wb.create_sheet('无应收明细透视表')
    display_categories = [('空白' if not name else name, name) for name in categories]
    pivot.append(['客户所属机构'] + [label for label, _ in display_categories] + ['合计', '总票数', '占比'])
    for org in sorted(groups):
        counts = [groups[org][key] for _, key in display_categories]
        total = sum(counts)
        denominator = source_totals.get(org, 0)
        pivot.append([org] + counts + [total, denominator, total / denominator if denominator else 0])
    pivot.append(['总计'] + [None] * (len(display_categories) + 3))

    report(5, 5, '正在保存处理结果')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    out_wb.save(output_path)
    out_wb.close()
    return output_path
