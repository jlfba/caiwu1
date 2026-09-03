# -*- coding: utf-8 -*-
"""报表组 Excel 处理逻辑。"""
from collections import Counter, defaultdict
import os
import re

from openpyxl import load_workbook
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
