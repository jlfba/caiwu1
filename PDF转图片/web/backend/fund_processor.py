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
from datetime import date

import openpyxl
from openpyxl.styles import PatternFill

_FILL_LIGHT_RED = PatternFill(fill_type='solid', fgColor='FFC7CE')
_FILL_YELLOW    = PatternFill(fill_type='solid', fgColor='FFFF00')


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


def _auto_sheet(wb, hints: list[str]) -> str:
    """优先找名称含 hints 任意词的工作表，否则取第一张。"""
    for name in wb.sheetnames:
        if any(kw in name for kw in hints):
            return name
    return wb.sheetnames[0]


def _process_one(wb_audit, amount_col_name: str, hint_words: list[str],
                 ws_sys, col_日期: int, col_tx: int, tx_label: str,
                 target_date: date, tmp_sheet_name: str,
                 sys_marked: set[int]):
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

    # 删合计行（仅金额列有值，其余空）
    last_r = ws.max_row
    if last_r >= 2:
        last_vals = {c: ws.cell(last_r, c).value for c in range(1, ws.max_column + 1)}
        non_empty = {c for c, v in last_vals.items() if v is not None and str(v).strip() != ''}
        if non_empty and non_empty.issubset({col_金额}):
            ws.delete_rows(last_r)

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

    if not filtered_rows:
        return False, f'系统表中未找到 {target_date} 的"{tx_label}"数据'

    # 写 E 列
    for i, (_, val) in enumerate(filtered_rows, start=2):
        ws_tmp.cell(i, 5, val)

    # A 列金额 → 行号映射
    a_map: dict[float, list[int]] = {}
    for r in range(2, len(data_rows) + 2):
        v = _to_num(ws_tmp.cell(r, 1).value)
        if v is not None:
            a_map.setdefault(v, []).append(r)

    a_marked: set[int] = set()
    matched_amounts: list[float] = []

    for i, _ in enumerate(filtered_rows):
        e_row = i + 2
        e_val = _to_num(ws_tmp.cell(e_row, 5).value)
        if e_val is not None and e_val in a_map:
            ws_tmp.cell(e_row, 6, e_val)
            matched_amounts.append(e_val)
            for a_row in a_map[e_val]:
                if a_row not in a_marked:
                    ws_tmp.cell(a_row, 1).fill = _FILL_LIGHT_RED
                    a_marked.add(a_row)
                    break
        else:
            ws_tmp.cell(e_row, 6).fill = _FILL_YELLOW

    # 系统表筛选行内标黄（共享 sys_marked 防跨两次误标）
    matched_set = set(matched_amounts)
    for sys_row, sys_val in filtered_rows:
        v = _to_num(sys_val)
        if v is not None and v in matched_set and sys_row not in sys_marked:
            ws_sys.cell(sys_row, col_tx).fill = _FILL_YELLOW
            sys_marked.add(sys_row)

    return True, ''


def process_fund(file1_path: str, file2_path: str, file3_path: str,
                 out_dir: str, progress) -> str:
    """
    主处理入口。
    file1_path: 收款审核表
    file2_path: 服务商付款审核表
    file3_path: 中信对公
    返回输出 ZIP 文件路径。
    """
    progress(1, 5, '正在读取收款审核表…')
    wb1 = openpyxl.load_workbook(file1_path)
    date_suffix1 = _extract_date(file1_path)
    target_date1 = _parse_date(date_suffix1)

    progress(2, 5, '正在读取中信对公系统表…')
    wb3 = openpyxl.load_workbook(file3_path)
    if '系统' not in wb3.sheetnames:
        raise ValueError(f'中信对公文件中未找到"系统"工作表，当前工作表：{wb3.sheetnames}')
    ws_sys = wb3['系统']
    headers_sys = [ws_sys.cell(1, c).value for c in range(1, ws_sys.max_column + 1)]
    col_日期  = _find_col(headers_sys, '交易日期')
    col_贷方  = _find_col(headers_sys, '贷方发生额')
    col_借方  = _find_col(headers_sys, '借方发生额')

    sys_marked: set[int] = set()  # 系统表已标色行（两次核对共享）

    progress(3, 5, '正在核对收款数据…')
    ok1, warn1 = _process_one(
        wb1, '收款金额', ['收款', '审核'],
        ws_sys, col_日期, col_贷方, '贷方发生额',
        target_date1, '资金核对-收款', sys_marked
    )
    if not ok1:
        raise ValueError(f'收款核对失败：{warn1}')

    progress(4, 5, '正在核对付款数据…')
    wb2 = openpyxl.load_workbook(file2_path)
    date_suffix2 = _extract_date(file2_path)
    target_date2 = _parse_date(date_suffix2)
    ok2, warn2 = _process_one(
        wb2, '付款金额', ['付款', '审核'],
        ws_sys, col_日期, col_借方, '借方发生额',
        target_date2, '资金核对-付款', sys_marked
    )
    if not ok2:
        raise ValueError(f'付款核对失败：{warn2}')

    progress(5, 5, '正在保存并打包结果…')

    name1 = f'收款审核表-资金核对-{date_suffix1}.xlsx'
    path1 = os.path.join(out_dir, name1)
    wb1.save(path1)

    name2 = f'付款审核表-资金核对-{date_suffix2}.xlsx'
    path2 = os.path.join(out_dir, name2)
    wb2.save(path2)

    name3 = f'中信对公-标色-{date_suffix1}.xlsx'
    path3 = os.path.join(out_dir, name3)
    wb3.save(path3)

    zip_name = f'资金核对-{date_suffix1}.zip'
    zip_path = os.path.join(out_dir, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(path1, name1)
        zf.write(path2, name2)
        zf.write(path3, name3)

    return zip_path
