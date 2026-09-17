# -*- coding: utf-8 -*-
"""资金组：收款审核表 + 服务商付款审核表 × 中信对公，一次处理两个核对。

处理流程：
1. 收款审核表 × 贷方发生额 → 资金核对-收款 临时表
2. 服务商付款审核表 × 借方发生额 → 资金核对-付款 临时表
3. 中信对公"系统"表在各自筛选行内标黄
4. 三个文件分别保存，打包 ZIP 输出
"""
import os
import re
import zipfile
import shutil
from copy import copy
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import openpyxl
from openpyxl.styles import PatternFill, Font

_FILL_LIGHT_RED   = PatternFill(fill_type='solid', fgColor='FFC7CE')
_FILL_YELLOW      = PatternFill(fill_type='solid', fgColor='FFFF00')
_FILL_LIGHT_GREEN = PatternFill(fill_type='solid', fgColor='C6EFCE')
_FILL_BLUE        = PatternFill(fill_type='solid', fgColor='9DC3E6')


def _find_col(headers: list, name: str) -> int:
    for i, h in enumerate(headers):
        if h is not None and str(h).strip() == name:
            return i + 1
    raise ValueError(f'未找到列"{name}"，请检查表头')


def _to_num(val) -> float | None:
    if val is None:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _money_key(val):
    """Return a two-decimal Decimal key for stable monetary matching."""
    if val is None or str(val).strip() == '':
        return None
    try:
        return Decimal(str(val).replace(',', '').strip()).quantize(
            Decimal('0.01'), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def _cell_date(val):
    """Normalize an Excel date/datetime/text cell to a date object."""
    if val is None:
        return None
    if hasattr(val, 'date'):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return date.fromisoformat(str(val).strip().replace('/', '-')[:10])
    except ValueError:
        return None


def _parse_date(date_suffix: str) -> date:
    return date(2000 + int(date_suffix[:2]), int(date_suffix[2:4]), int(date_suffix[4:6]))


def _extract_date(file_path: str) -> str:
    """从文件名提取 6 位日期后缀，找不到抛出 ValueError。"""
    m = re.search(r'(\d{6})', os.path.basename(file_path))
    if not m:
        raise ValueError(
            f'无法从文件名中识别 6 位日期后缀（如 260910），当前文件名：{os.path.basename(file_path)}'
        )
    return m.group(1)


def _original_upload_name(file_path: str) -> str:
    """Remove the internal fund upload prefix while preserving the user's filename."""
    return re.sub(r'^fund\d+_', '', os.path.basename(file_path), count=1)


def identify_fund_files(file_paths):
    """Identify fund workbooks by filename keywords; only the public-account workbook is required."""
    rules = {
        'receipt': ('收款审核',),
        'payment': ('服务商付款',),
        'system': ('对公',),
        'bank_flow': ('流水',),
    }
    found = {}
    unknown = []
    duplicates = []
    for path in file_paths:
        name = os.path.basename(path)
        # 文件名可能同时含“中信对公”和业务表关键词。
        # 流水优先，其次必须识别收款/付款审核表；只有不含这些业务关键词时，
        # “对公”才代表中信对公系统表。
        if any(keyword in name for keyword in rules['bank_flow']):
            matches = ['bank_flow']
        else:
            specific = [key for key in ('receipt', 'payment')
                        if any(keyword in name for keyword in rules[key])]
            if specific:
                matches = specific
            else:
                matches = ['system'] if any(keyword in name for keyword in rules['system']) else []
        if len(matches) == 1:
            kind = matches[0]
            if kind not in found:
                found[kind] = path
            else:
                duplicates.append(kind)
        else:
            unknown.append(name)
    if 'system' not in found or unknown or duplicates:
        labels = {'receipt': '收款审核表', 'payment': '服务商付款审核表',
                  'system': '中信对公', 'bank_flow': '银行账号管理流水'}
        detail = []
        if 'system' not in found:
            detail.append('缺少：' + labels['system'])
        if unknown:
            detail.append('无法识别：' + '、'.join(unknown))
        if duplicates:
            duplicate_labels = [labels[kind] for kind in dict.fromkeys(duplicates)]
            detail.append('重复上传：' + '、'.join(duplicate_labels))
        raise ValueError('资金组文件识别失败；' + '；'.join(detail))
    return (found.get('receipt'), found.get('payment'),
            found['system'], found.get('bank_flow'))


def _auto_sheet(wb, hints: list[str]) -> str:
    """优先找名称含 hints 任意词的工作表，否则取第一张。"""
    for name in wb.sheetnames:
        if any(kw in name for kw in hints):
            return name
    return wb.sheetnames[0]


def _is_summary_row(ws, row: int) -> bool:
    """Identify a total row without relying on a fixed column layout."""
    values = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
    text_values = [str(value).strip() for value in values
                   if value is not None and str(value).strip()]
    if any(any(word in value for word in ('合计', '总计', '汇总')) for value in text_values):
        return True

    # Some source exports put only numeric totals in the final row, without a “合计” label.
    # A normal transaction always contains at least one descriptive text, date, account or reference value.
    has_number = False
    for value in values:
        if value is None or str(value).strip() == '':
            continue
        if _money_key(value) is not None:
            has_number = True
            continue
        if isinstance(value, (date, datetime)):
            return False
        return False
    return has_number


def _process_one(wb_audit, amount_col_name: str, hint_words: list[str],
                 ws_sys, col_日期: int, col_tx: int, tx_label: str,
                 target_date: date, tmp_sheet_name: str,
                 sys_marked: set[tuple[int, int]], sys_fill,
                 match_previous_day: bool = False,
                 match_negative_as_debit: bool = False,
                 match_prior_month: bool = False):
    """
    对一张审核表做匹配核对，结果写入临时工作表，并在 ws_sys 内标色。
    sys_marked: 已在系统表标色的行号集合（跨两次调用共享，防重复）
    返回 (bool, str)：(是否找到日期数据, 警告信息)
    """
    # 选工作表
    sheet_name = _auto_sheet(wb_audit, hint_words)
    ws = wb_audit[sheet_name]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    col_金额 = _find_col(headers, amount_col_name)
    col_备注 = _find_col(headers, '备注')
    col_result = 19
    ws.cell(1, col_result, '查找结果')

    # 末尾合计行不属于待核对业务数据；删除后不会生成“未找到”。
    while ws.max_row >= 2 and _is_summary_row(ws, ws.max_row):
        ws.delete_rows(ws.max_row)

    # 收集数据行
    data_rows = [(ws.cell(r, col_金额).value, ws.cell(r, col_备注).value)
                 for r in range(2, ws.max_row + 1)]

    # 创建临时工作表
    if tmp_sheet_name in wb_audit.sheetnames:
        del wb_audit[tmp_sheet_name]
    ws_tmp = wb_audit.create_sheet(tmp_sheet_name)
    ws_tmp['A1'] = amount_col_name
    ws_tmp['B1'] = '备注'
    ws_tmp['E1'] = tx_label
    ws_tmp['F1'] = f'匹配{amount_col_name}'
    for i, (金额, 备注) in enumerate(data_rows, start=2):
        ws_tmp.cell(i, 1, 金额)
        ws_tmp.cell(i, 2, 备注)

    # 从系统表筛选当天数据
    filtered_rows: list[tuple[int, object]] = []
    for r in range(2, ws_sys.max_row + 1):
        cell_date = ws_sys.cell(r, col_日期).value
        if cell_date is None:
            continue
        if hasattr(cell_date, 'date'):
            row_date = cell_date.date()
        elif isinstance(cell_date, date):
            row_date = cell_date
        else:
            s = str(cell_date).strip().replace('/', '-')
            try:
                row_date = date.fromisoformat(s[:10])
            except ValueError:
                continue
        if row_date == target_date:
            filtered_rows.append((r, ws_sys.cell(r, col_tx).value))

    if not filtered_rows and not (match_previous_day or match_prior_month):
        return False, f'系统表中未找到 {target_date} 的"{tx_label}"数据'

    # 写 E 列
    for i, (_, val) in enumerate(filtered_rows, start=2):
        ws_tmp.cell(i, 5, val)

    # 备注有“开票税点=水单总额”或“运费 + 税点”时，优先以水单总额核对；否则使用审核金额。
    # 负数收款金额不参与贷方匹配，须以绝对值在借方发生额中单独核对。
    a_map: dict[Decimal, list[int]] = {}
    negative_amounts: dict[Decimal, list[int]] = {}
    for r, (amount, note) in enumerate(data_rows, start=2):
        raw_amount = _money_key(amount)
        if match_negative_as_debit and raw_amount is not None and raw_amount < 0:
            negative_amounts.setdefault(abs(raw_amount), []).append(r)
            continue
        v = _invoice_tax_total(note) or _freight_tax_total(note) or raw_amount
        if v is not None:
            a_map.setdefault(v, []).append(r)

    a_marked: set[int] = set()
    matched_a_rows: set[int] = set()

    for i, _ in enumerate(filtered_rows):
        e_row = i + 2
        e_val = _money_key(ws_tmp.cell(e_row, 5).value)
        if e_val is not None and e_val in a_map:
            ws_tmp.cell(e_row, 6, e_val)
            for a_row in a_map[e_val]:
                if a_row not in a_marked:
                    ws_tmp.cell(a_row, 1).fill = _FILL_LIGHT_RED
                    a_marked.add(a_row)
                    matched_a_rows.add(a_row)
                    break
        else:
            ws_tmp.cell(e_row, 6).fill = _FILL_YELLOW

    # 收款审核表当天未找到的金额，补查对公表的前一天。
    # 该补查只在当天核对完成后执行，且不影响服务商付款的核对规则。
    if match_previous_day:
        previous_date = target_date - timedelta(days=1)
        for sys_row in range(2, ws_sys.max_row + 1):
            if _cell_date(ws_sys.cell(sys_row, col_日期).value) != previous_date:
                continue
            sys_value = _money_key(ws_sys.cell(sys_row, col_tx).value)
            if sys_value is None or sys_value not in a_map:
                continue
            for audit_row in a_map[sys_value]:
                if audit_row not in a_marked:
                    a_marked.add(audit_row)
                    matched_a_rows.add(audit_row)
                    # 纳入统一标色集合，使前一天的对公金额也会标黄。
                    filtered_rows.append((sys_row, ws_sys.cell(sys_row, col_tx).value))
                    break

    # 收款金额为负数代表支出：去掉负号后，仅在当天借方发生额中匹配。
    if match_negative_as_debit and negative_amounts:
        sys_headers = [ws_sys.cell(1, c).value for c in range(1, ws_sys.max_column + 1)]
        col_debit = _find_col(sys_headers, '借方发生额')
        for sys_row in range(2, ws_sys.max_row + 1):
            if _cell_date(ws_sys.cell(sys_row, col_日期).value) != target_date:
                continue
            debit_value = _money_key(ws_sys.cell(sys_row, col_debit).value)
            if debit_value is None or debit_value not in negative_amounts:
                continue
            for audit_row in negative_amounts[debit_value]:
                if audit_row not in a_marked:
                    a_marked.add(audit_row)
                    matched_a_rows.add(audit_row)
                    mark_key = (sys_row, col_debit)
                    if mark_key not in sys_marked:
                        ws_sys.cell(sys_row, col_debit).fill = _FILL_LIGHT_GREEN
                        sys_marked.add(mark_key)
                    break

    # 当天（及收款的前一天）仍未找到时，在前一个月内补查。
    # 只使用对公发生额中“无填充”的单元格，避免占用先前已经核对过的金额。
    if match_prior_month:
        earliest_date = target_date - timedelta(days=31)
        for sys_row in range(2, ws_sys.max_row + 1):
            row_date = _cell_date(ws_sys.cell(sys_row, col_日期).value)
            if row_date is None or not earliest_date <= row_date <= target_date:
                continue
            system_cell = ws_sys.cell(sys_row, col_tx)
            if system_cell.fill.fill_type is not None:
                continue
            system_value = _money_key(system_cell.value)
            if system_value is None or system_value not in a_map:
                continue
            for audit_row in a_map[system_value]:
                if audit_row not in a_marked:
                    a_marked.add(audit_row)
                    matched_a_rows.add(audit_row)
                    # 纳入统一标色集合，按收款黄、付款绿回写系统表。
                    filtered_rows.append((sys_row, system_cell.value))
                    break

    # 系统表筛选行内标色（共享 sys_marked 防跨两次误标）
    for a_row in range(2, len(data_rows) + 2):
        amount = _to_num(ws.cell(a_row, col_金额).value)
        if amount is None:
            ws.cell(a_row, col_result, '')
        elif a_row in matched_a_rows:
            ws.cell(a_row, col_result, '已找到')
        else:
            ws.cell(a_row, col_result, '未找到')
            for cell in ws[a_row]:
                if cell.column != col_result:
                    cell.fill = _FILL_BLUE

    matched_set = {
        _invoice_tax_total(data_rows[a_row - 2][1])
        or _freight_tax_total(data_rows[a_row - 2][1])
        or _money_key(data_rows[a_row - 2][0])
        for a_row in matched_a_rows
    }
    for sys_row, sys_val in filtered_rows:
        v = _money_key(sys_val)
        mark_key = (sys_row, col_tx)
        if v is not None and v in matched_set and mark_key not in sys_marked:
            ws_sys.cell(sys_row, col_tx).fill = sys_fill
            sys_marked.add(mark_key)

    return True, ''


def _tax_amount(note):
    """Extract the amount written after 税点, with or without an equals sign."""
    if note is None:
        return None
    match = re.search(r'税点\s*[=:：]?\s*([+-]?\d[\d,]*(?:\.\d+)?)', str(note))
    return _money_key(match.group(1)) if match else None


def _invoice_tax_total(note):
    """Extract the stated total after “开票税点=” / “开票税点：”."""
    if note is None:
        return None
    match = re.search(r'开票税点\s*[=:：]\s*([+-]?\d[\d,]*(?:\.\d+)?)', str(note))
    return _money_key(match.group(1)) if match else None


def _freight_tax_total(note):
    """Return 运费 + 税点 amount written in a remark, when both are present."""
    if note is None:
        return None
    text = str(note)
    freight = re.search(r'运费\s*[=:：]?\s*([+-]?\d[\d,]*(?:\.\d+)?)', text)
    tax = re.search(r'税点\s*[=:：]?\s*([+-]?\d[\d,]*(?:\.\d+)?)', text)
    if not freight or not tax:
        return None
    freight_amount = _money_key(freight.group(1))
    tax_amount = _money_key(tax.group(1))
    if freight_amount is None or tax_amount is None:
        return None
    return (freight_amount + tax_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


def _process_bank_account_flow(wb_flow, ws_sys, target_date: date,
                               sys_marked: set[tuple[int, int]]):
    """Match non-sales bank-account flow income/expense rows to the system sheet."""
    ws = wb_flow[_auto_sheet(wb_flow, ['银行账号管理流水', '流水'])]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_type = _find_col(headers, '费用类型')
    col_income = _find_col(headers, '收入')
    col_expense = _find_col(headers, '支出')
    col_note = _find_col(headers, '备注')
    col_result = next((i + 1 for i, header in enumerate(headers)
                       if str(header).strip() == '查找结果'), None)
    if col_result is None:
        col_result = ws.max_column + 1
        ws.cell(1, col_result, '查找结果')

    sys_headers = [ws_sys.cell(1, c).value for c in range(1, ws_sys.max_column + 1)]
    col_date = _find_col(sys_headers, '交易日期')
    col_credit = _find_col(sys_headers, '贷方发生额')
    col_debit = _find_col(sys_headers, '借方发生额')
    system_rows = [r for r in range(2, ws_sys.max_row + 1)
                   if _cell_date(ws_sys.cell(r, col_date).value) == target_date]
    if not system_rows:
        raise ValueError(f'系统表中未找到 {target_date} 的交易日期数据')

    used_credit, used_debit = set(), set()

    def match(amount, column, used_rows, fill):
        if amount is None:
            return False
        for row in system_rows:
            if row in used_rows:
                continue
            if _money_key(ws_sys.cell(row, column).value) == amount:
                ws_sys.cell(row, column).fill = fill
                sys_marked.add((row, column))
                used_rows.add(row)
                return True
        return False

    processed = 0
    for row in range(2, ws.max_row + 1):
        if _is_summary_row(ws, row):
            continue
        note = str(ws.cell(row, col_note).value or '')
        # 先于费用类型筛选处理：无论所属费用类型，均记作作废且不参与核对。
        if '修改付款' in note or '作废' in note:
            ws.cell(row, col_result, '作废')
            continue
        fee_type = str(ws.cell(row, col_type).value or '').strip()
        if fee_type in ('销售收入', '销售成本'):
            continue
        income = _money_key(ws.cell(row, col_income).value)
        expense = _money_key(ws.cell(row, col_expense).value)

        if income not in (None, Decimal('0.00')):
            processed += 1
            if '税点' in note:
                # 水单备注如“运费4998.80+税点12”应按合计 5010.80 匹配。
                amount = _freight_tax_total(note) or _tax_amount(note)
                if match(amount, col_credit, used_credit, _FILL_YELLOW):
                    ws.cell(row, col_result, '已找到')
                else:
                    ws.cell(row, col_result, '未找到')
                    for cell in ws[row]:
                        cell.fill = _FILL_BLUE
            elif fee_type == '其他收入':
                if match(income, col_credit, used_credit, _FILL_YELLOW):
                    ws.cell(row, col_result, '已找到')
                else:
                    ws.cell(row, col_result, '未找到')
                    for cell in ws[row]:
                        cell.fill = _FILL_BLUE
            else:
                ws.cell(row, col_result, '人工处理')

        if expense not in (None, Decimal('0.00')):
            processed += 1
            if '美金转账手续费' in note:
                for cell in ws[row]:
                    cell.fill = _FILL_LIGHT_GREEN
                ws.cell(row, col_result, '人工处理')
            elif match(expense, col_debit, used_debit, _FILL_LIGHT_GREEN):
                ws.cell(row, col_result, '已找到')
            else:
                ws.cell(row, col_result, '未找到')
                for cell in ws[row]:
                    cell.fill = _FILL_BLUE
    return processed


def _append_bank_flow_to_rmb(wb_flow, wb_system):
    """Clear flow balances and append rows after the last real RMB column-A row."""
    flow_ws = wb_flow[_auto_sheet(wb_flow, ['银行账号管理流水', '流水'])]
    if '人民币' not in wb_system.sheetnames:
        raise ValueError(f'中信对公文件中未找到"人民币"工作表，当前工作表：{wb_system.sheetnames}')
    rmb_ws = wb_system['人民币']

    # 银行账号管理流水的 E 列为余额，保留表头、清空数据。
    for row in range(2, flow_ws.max_row + 1):
        flow_ws.cell(row, 5).value = None

    last_flow_row = 1
    for row in range(2, flow_ws.max_row + 1):
        if any(flow_ws.cell(row, col).value not in (None, '')
               for col in range(1, flow_ws.max_column + 1)):
            last_flow_row = row
    headers = [flow_ws.cell(1, col).value for col in range(1, flow_ws.max_column + 1)]
    col_income = _find_col(headers, '收入')
    col_expense = _find_col(headers, '支出')
    col_result = _find_col(headers, '查找结果')
    source_rows = []
    for row in range(2, last_flow_row + 1):
        # 合计行仅作源表统计，不参与人民币追加。
        if _is_summary_row(flow_ws, row):
            continue
        values = [flow_ws.cell(row, col).value for col in range(1, flow_ws.max_column + 1)]
        unmatched = str(flow_ws.cell(row, col_result).value or '').strip() == '未找到'
        income = _money_key(flow_ws.cell(row, col_income).value)
        expense = _money_key(flow_ws.cell(row, col_expense).value)
        p_value = None
        if unmatched and income not in (None, Decimal('0.00')):
            p_value = -income
        elif unmatched and expense not in (None, Decimal('0.00')):
            p_value = expense
        source_rows.append({'values': values, 'unmatched': unmatched, 'p_value': p_value})
    if not source_rows:
        return 0

    # 只按 A 列真实数据定位续接位置，避免 max_row 被格式残留推到很远。
    last_rmb_row = 1
    for row in range(2, rmb_ws.max_row + 1):
        if rmb_ws.cell(row, 1).value not in (None, ''):
            last_rmb_row = row
    insert_at = last_rmb_row + 1
    red_styles = {}
    signed_red_style = None
    for col in range(1, max(rmb_ws.max_column, flow_ws.max_column, 16) + 1):
        sample = rmb_ws.cell(insert_at, col)
        sample._style = copy(rmb_ws.cell(last_rmb_row, col)._style)
        sample.font = copy(sample.font)
        sample.font = sample.font.copy(color='FFFF0000')
        red_styles[openpyxl.utils.get_column_letter(col)] = sample.style_id
        if col == 16:
            sample.number_format = '+0.00;-0.00;0.00'
            signed_red_style = sample.style_id
    return {
        'insert_at': insert_at,
        'last_row': last_rmb_row,
        'formula': rmb_ws.cell(last_rmb_row, 5).value,
        'result_column': col_result,
        'styles': [rmb_ws.cell(last_rmb_row, col).style_id
                   for col in range(1, max(rmb_ws.max_column, flow_ws.max_column, 16) + 1)],
        'red_styles': red_styles,
        'signed_red_style': signed_red_style,
        'rows': source_rows,
    }


def _save_system_workbook_preserving_template(original_path: str, workbook, output_path: str,
                                             rmb_payload=None, sys_marked=None):
    """Use Excel's native save to preserve this template's Excel-only structures."""
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            '当前中信对公模板包含 Excel 专有的数组公式和扩展结构，'
            'Linux 纯 Python 保存会导致 Excel 修复并丢失处理结果。'
            '请使用安装了 Microsoft Excel 的 Windows 版本处理资金组。'
        ) from exc

    def excel_color(rgb: str) -> int:
        rgb = rgb[-6:]
        red, green, blue = int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16)
        return red | (green << 8) | (blue << 16)

    def com_value(value):
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, date) and not isinstance(value, datetime):
            return datetime(value.year, value.month, value.day)
        return value

    known_fills = {'FFFF00', 'C6EFCE', 'FFC7CE', '9DC3E6'}
    shutil.copy2(original_path, output_path)
    excel = None
    native_book = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        native_book = excel.Workbooks.Open(os.path.abspath(output_path), UpdateLinks=0, ReadOnly=False)

        # 只写入核对流程实际标记过的系统表单元格。
        # 不能遍历整张系统表：该模板虽然约一万行，却带有 16,384 列的格式范围，
        # 逐个 COM 单元格读取会让保存阶段耗时数分钟。
        system_source = workbook['系统']
        system_target = native_book.Worksheets('系统')
        for row, column in (sys_marked or set()):
            source_cell = system_source.cell(row, column)
            color = (source_cell.fill.fgColor.rgb or '').upper()
            if color[-6:] in known_fills:
                system_target.Cells(row, column).Interior.Color = excel_color(color)

        if rmb_payload:
            rmb_sheet = native_book.Worksheets('人民币')
            insert_at = rmb_payload['insert_at']
            row_count = len(rmb_payload['rows'])
            last_row = rmb_payload['last_row']
            # Native insertion updates formulas, references, drawing anchors, merged cells and dimensions.
            rmb_sheet.Rows(f'{insert_at}:{insert_at + row_count - 1}').Insert()
            for offset, item in enumerate(rmb_payload['rows']):
                target_row = insert_at + offset
                values = [
                    None if column in (rmb_payload['result_column'], 5) else com_value(value)
                    for column, value in enumerate(item['values'], start=1)
                ]
                # 一次写入整行，避免每个单元格跨进程调用 Excel。
                rmb_sheet.Range(
                    rmb_sheet.Cells(target_row, 1),
                    rmb_sheet.Cells(target_row, len(values)),
                ).Value = [values]
                if item['unmatched']:
                    rmb_sheet.Rows(target_row).Font.Color = excel_color('FF0000')
                if item['p_value'] is not None:
                    p_cell = rmb_sheet.Cells(target_row, 16)
                    p_cell.Value = com_value(item['p_value'])
                    p_cell.NumberFormat = '+0.00;-0.00;0.00'
            # The balance formula is extended by Excel itself, preserving formulas that reference other sheets.
            source = rmb_sheet.Range(f'E{last_row}')
            destination = rmb_sheet.Range(f'E{last_row}:E{insert_at + row_count - 1}')
            source.AutoFill(destination)

        native_book.Save()
    finally:
        if native_book is not None:
            native_book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def process_fund(file1_path: str | None, file2_path: str | None, file3_path: str, file4_path: str | None,
                 out_dir: str, progress) -> str:
    """
    主处理入口。
    file1_path: 可选收款审核表
    file2_path: 可选服务商付款审核表
    file3_path: 中信对公
    file4_path: 可选银行账号管理流水
    返回输出 ZIP 文件路径。
    """
    date_source = next(path for path in (file1_path, file2_path, file4_path, file3_path) if path)
    date_suffix = _extract_date(date_source)
    wb1 = None
    wb2 = None
    wb4 = None

    progress(2, 7, '正在读取中信对公系统表…')
    wb3 = openpyxl.load_workbook(file3_path)
    if '系统' not in wb3.sheetnames:
        raise ValueError(f'中信对公文件中未找到"系统"工作表，当前工作表：{wb3.sheetnames}')
    ws_sys = wb3['系统']
    headers_sys = [ws_sys.cell(1, c).value for c in range(1, ws_sys.max_column + 1)]
    col_日期  = _find_col(headers_sys, '交易日期')
    col_贷方  = _find_col(headers_sys, '贷方发生额')
    col_借方  = _find_col(headers_sys, '借方发生额')

    sys_marked: set[tuple[int, int]] = set()  # 系统表已标色单元格（两次核对共享）

    if file1_path:
        progress(3, 7, '正在核对收款数据…')
        wb1 = openpyxl.load_workbook(file1_path)
        target_date1 = _parse_date(_extract_date(file1_path))
        ok1, warn1 = _process_one(
            wb1, '收款金额', ['收款', '审核'],
            ws_sys, col_日期, col_贷方, '贷方发生额',
            target_date1, '资金核对-收款', sys_marked, _FILL_YELLOW,
            match_previous_day=True, match_negative_as_debit=True, match_prior_month=True
        )
        if not ok1:
            raise ValueError(f'收款核对失败：{warn1}')
    else:
        progress(3, 7, '未上传收款审核表，已跳过…')

    if file2_path:
        progress(4, 7, '正在核对付款数据…')
        wb2 = openpyxl.load_workbook(file2_path)
        target_date2 = _parse_date(_extract_date(file2_path))
        ok2, warn2 = _process_one(
            wb2, '付款金额', ['付款', '审核'],
            ws_sys, col_日期, col_借方, '借方发生额',
            target_date2, '资金核对-付款', sys_marked, _FILL_LIGHT_GREEN,
            match_prior_month=True
        )
        if not ok2:
            raise ValueError(f'付款核对失败：{warn2}')
    else:
        progress(4, 7, '未上传服务商付款表，已跳过…')

    rmb_payload = None
    if file4_path:
        progress(5, 7, '正在核对银行账号管理流水…')
        wb4 = openpyxl.load_workbook(file4_path)
        date_suffix4 = _extract_date(file4_path)
        target_date4 = _parse_date(date_suffix4)
        _process_bank_account_flow(wb4, ws_sys, target_date4, sys_marked)
        progress(6, 7, '正在追加银行流水到人民币工作表…')
        rmb_payload = _append_bank_flow_to_rmb(wb4, wb3)
    else:
        progress(5, 7, '未上传银行账号管理表，已跳过…')
        progress(6, 7, '未上传银行账号管理表，已跳过追加…')

    progress(7, 7, '正在保存并打包结果…')

    output_files = []
    if wb1:
        name1 = _original_upload_name(file1_path)
        path1 = os.path.join(out_dir, name1)
        wb1.save(path1)
        output_files.append((path1, name1))
    if wb2:
        name2 = _original_upload_name(file2_path)
        path2 = os.path.join(out_dir, name2)
        wb2.save(path2)
        output_files.append((path2, name2))

    name3 = _original_upload_name(file3_path)
    path3 = os.path.join(out_dir, name3)
    _save_system_workbook_preserving_template(file3_path, wb3, path3, rmb_payload, sys_marked)
    output_files.append((path3, name3))

    if wb4:
        name4 = _original_upload_name(file4_path)
        path4 = os.path.join(out_dir, name4)
        wb4.save(path4)
        output_files.append((path4, name4))

    zip_name = f'资金核对-{date_suffix}.zip'
    zip_path = os.path.join(out_dir, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for path, name in output_files:
            zf.write(path, name)

    return zip_path
