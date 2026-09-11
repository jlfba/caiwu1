# -*- coding: utf-8 -*-
"""资金组：审核表 × 中信对公对照处理。

支持两种核对模式：
  receipt  收款核对：收款审核表 × 贷方发生额
  payment  付款核对：服务商付款审核表 × 借方发生额

处理流程：
1. 从审核表文件名识别日期后缀，删合计行，提取金额/备注 → 临时表 A/B 列
2. 从中信对公"系统"表按日期筛选对应发生额 → 临时表 E 列
3. E 列 vs A 列匹配：找到→ F 列填值 + A 列标浅红；未找到→ F 列标黄
4. 在系统表筛选行内将匹配金额标黄
5. 两文件分别保存，打包 ZIP 输出
"""
import os
import re
import zipfile
from datetime import date

import openpyxl
from openpyxl.styles import PatternFill

# 颜色定义
_FILL_LIGHT_RED = PatternFill(fill_type='solid', fgColor='FFC7CE')   # A列：已匹配
_FILL_YELLOW    = PatternFill(fill_type='solid', fgColor='FFFF00')    # 未匹配 / 系统表已匹配

# 模式配置
_MODE_CFG = {
    'receipt': {
        'msg1':        '正在读取收款审核表…',
        'amount_col':  '收款金额',
        'key_cols_extra': ['本位币'],   # 合计行判断时额外的关键列
        'tx_col':      '贷方发生额',
        'match_label': '匹配收款金额',
        'sheet_hint':  ['收款', '审核'],
        'file1_label': '收款审核表',
        'file2_label': '中信对公（贷方）',
    },
    'payment': {
        'msg1':        '正在读取服务商付款审核表…',
        'amount_col':  '付款金额',
        'key_cols_extra': [],
        'tx_col':      '借方发生额',
        'match_label': '匹配付款金额',
        'sheet_hint':  ['付款', '审核'],
        'file1_label': '付款审核表',
        'file2_label': '中信对公（借方）',
    },
}


def _find_col(headers: list, name: str) -> int:
    """在 headers 中找列名的 1-based 索引；找不到抛 ValueError。"""
    for i, h in enumerate(headers):
        if h is not None and str(h).strip() == name:
            return i + 1
    raise ValueError(f'未找到列"{name}"，请检查表头')


def _to_num(val) -> float | None:
    """单元格值转 float，无效返回 None。"""
    if val is None:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _parse_date(date_suffix: str) -> date:
    """YYMMDD → date(20YY, MM, DD)"""
    yy = int(date_suffix[:2])
    mm = int(date_suffix[2:4])
    dd = int(date_suffix[4:6])
    return date(2000 + yy, mm, dd)


def process_fund(file1_path: str, file2_path: str, out_dir: str,
                 progress, fund_mode: str = 'receipt') -> str:
    """
    主处理入口。
    fund_mode: 'receipt'（收款核对）或 'payment'（付款核对）
    返回输出 ZIP 文件的绝对路径。
    """
    cfg = _MODE_CFG.get(fund_mode, _MODE_CFG['receipt'])

    # ── Step 1：读取审核表 ────────────────────────────────────────────────────
    progress(1, 5, cfg['msg1'])
    wb1 = openpyxl.load_workbook(file1_path)

    # 从文件名提取 6 位日期后缀
    basename = os.path.basename(file1_path)
    m = re.search(r'(\d{6})', basename)
    if not m:
        raise ValueError(
            f'无法从文件名中识别 6 位日期后缀（如 260910），当前文件名：{basename}'
        )
    date_suffix = m.group(1)
    target_date = _parse_date(date_suffix)

    # 自动选工作表：优先找名称含模式关键词的，否则取第一张
    target_sheet = None
    for name in wb1.sheetnames:
        if any(kw in name for kw in cfg['sheet_hint']):
            target_sheet = name
            break
    if target_sheet is None:
        target_sheet = wb1.sheetnames[0]

    ws1 = wb1[target_sheet]
    headers1 = [ws1.cell(1, c).value for c in range(1, ws1.max_column + 1)]

    col_金额 = _find_col(headers1, cfg['amount_col'])
    col_备注 = _find_col(headers1, '备注')

    # 合计行判断：最后一行只有金额列（+ 额外关键列）有数据，其余全空
    extra_key_cols = set()
    for col_name in cfg['key_cols_extra']:
        try:
            extra_key_cols.add(_find_col(headers1, col_name))
        except ValueError:
            pass
    key_cols = {col_金额} | extra_key_cols

    last_r = ws1.max_row
    if last_r >= 2:
        last_vals = {c: ws1.cell(last_r, c).value
                     for c in range(1, ws1.max_column + 1)}
        non_empty = {c for c, v in last_vals.items()
                     if v is not None and str(v).strip() != ''}
        if non_empty and non_empty.issubset(key_cols):
            ws1.delete_rows(last_r)

    # 收集数据行
    data_rows = []
    for r in range(2, ws1.max_row + 1):
        data_rows.append((
            ws1.cell(r, col_金额).value,
            ws1.cell(r, col_备注).value,
        ))

    # 创建临时工作表
    tmp_name = '资金核对'
    if tmp_name in wb1.sheetnames:
        del wb1[tmp_name]
    ws_tmp = wb1.create_sheet(tmp_name)
    ws_tmp['A1'] = cfg['amount_col']
    ws_tmp['B1'] = '备注'
    ws_tmp['E1'] = cfg['tx_col']
    ws_tmp['F1'] = cfg['match_label']
    for i, (金额, 备注) in enumerate(data_rows, start=2):
        ws_tmp.cell(i, 1, 金额)
        ws_tmp.cell(i, 2, 备注)

    # ── Step 2：读取中信对公"系统"表，按日期筛选 ──────────────────────────────
    progress(2, 5, '正在读取中信对公系统表…')
    wb2 = openpyxl.load_workbook(file2_path)
    if '系统' not in wb2.sheetnames:
        raise ValueError(
            f'中信对公文件中未找到"系统"工作表，当前工作表：{wb2.sheetnames}'
        )
    ws_sys = wb2['系统']
    headers2 = [ws_sys.cell(1, c).value for c in range(1, ws_sys.max_column + 1)]
    col_日期 = _find_col(headers2, '交易日期')
    col_tx   = _find_col(headers2, cfg['tx_col'])

    # 仅取符合目标日期的行
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
        raise ValueError(
            f'在中信对公"系统"表的"交易日期"列中未找到 {target_date} 的数据，'
            '请确认日期格式与表后缀一致'
        )

    # 写入临时表 E 列
    for i, (_, val) in enumerate(filtered_rows, start=2):
        ws_tmp.cell(i, 5, val)

    # ── Step 3：匹配与标色 ─────────────────────────────────────────────────────
    progress(3, 5, '正在匹配金额并标注颜色…')

    # A 列金额 → 行号列表
    a_map: dict[float, list[int]] = {}
    for r in range(2, len(data_rows) + 2):
        v = _to_num(ws_tmp.cell(r, 1).value)
        if v is not None:
            a_map.setdefault(v, []).append(r)

    a_marked: set[int] = set()
    matched_amounts: list[float] = []

    for i, (_, _val) in enumerate(filtered_rows):
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

    # 系统表筛选行内标黄（防止跨日期误标）
    matched_set = set(matched_amounts)
    sys_marked: set[int] = set()
    for sys_row, sys_val in filtered_rows:
        v = _to_num(sys_val)
        if v is not None and v in matched_set and sys_row not in sys_marked:
            ws_sys.cell(sys_row, col_tx).fill = _FILL_YELLOW
            sys_marked.add(sys_row)

    # ── Step 4：保存 + 打包 ZIP ───────────────────────────────────────────────
    progress(4, 5, '正在保存结果文件…')

    label = '收款审核表' if fund_mode == 'receipt' else '付款审核表'
    name1 = f'{label}-资金核对-{date_suffix}.xlsx'
    path1 = os.path.join(out_dir, name1)
    wb1.save(path1)

    name2 = f'中信对公-标色-{date_suffix}.xlsx'
    path2 = os.path.join(out_dir, name2)
    wb2.save(path2)

    zip_name = f'资金核对-{date_suffix}.zip'
    zip_path = os.path.join(out_dir, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(path1, name1)
        zf.write(path2, name2)

    progress(5, 5, f'处理完成，输出：{zip_name}')
    return zip_path
