# -*- coding: utf-8 -*-
"""资金组：收款审核表 × 中信对公对照处理。

处理流程：
1. 从"收款审核表"识别日期后缀，删除合计行，提取 收款金额/备注 → 临时表 A/B 列
2. 从"中信对公"的"系统"工作表按日期筛选贷方发生额 → 临时表 E 列
3. 将 E 列值在 A 列查找匹配：找到→ F 列填值、A 列标浅红；未找到→ F 列标黄
4. 在"系统"表筛选行范围内，将 F 列匹配到的金额对应单元格标黄
5. 保存输出到结果文件
"""
import os
import re
from datetime import date

import openpyxl
from openpyxl.styles import PatternFill

# 颜色定义
_FILL_LIGHT_RED = PatternFill(fill_type='solid', fgColor='FFC7CE')   # A列：已匹配
_FILL_YELLOW    = PatternFill(fill_type='solid', fgColor='FFFF00')    # 未匹配标黄 / 系统表标黄


def _find_col(headers: list, name: str) -> int:
    """在 headers 列表中找到列名的 1-based 索引；找不到抛出 ValueError。"""
    for i, h in enumerate(headers):
        if h is not None and str(h).strip() == name:
            return i + 1
    raise ValueError(f'未找到列"{name}"，请检查表头')


def _to_num(val) -> float | None:
    """将单元格值转为 float，无效时返回 None。"""
    if val is None:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def process_fund(file1_path: str, file2_path: str, out_dir: str, progress) -> str:
    """
    主处理入口。
    file1_path: 收款审核表路径
    file2_path: 中信对公路径
    out_dir:    输出目录
    progress:   progress(cur, total, msg) 回调
    返回输出文件绝对路径。
    """
    progress(1, 5, '正在读取收款审核表…')

    # ── Step 1：读取收款审核表 ──────────────────────────────────────────────────
    wb1 = openpyxl.load_workbook(file1_path)

    # 找工作表并提取日期后缀
    target_sheet = None
    date_suffix = None
    for name in wb1.sheetnames:
        m = re.search(r'中信收款审核表-?(\d{6})', name)
        if m:
            target_sheet = name
            date_suffix = m.group(1)
            break
    if target_sheet is None:
        raise ValueError(
            '未找到名称含"中信收款审核表-YYMMDD"的工作表，'
            f'当前工作表：{wb1.sheetnames}'
        )

    # 解析日期 YYMMDD → date
    yy, mm, dd = int(date_suffix[:2]), int(date_suffix[2:4]), int(date_suffix[4:6])
    target_date = date(2000 + yy, mm, dd)

    ws1 = wb1[target_sheet]
    headers1 = [ws1.cell(1, c).value for c in range(1, ws1.max_column + 1)]

    col_金额  = _find_col(headers1, '收款金额')
    col_本位币 = _find_col(headers1, '本位币')
    col_备注  = _find_col(headers1, '备注')

    # 删合计行：最后一行仅 收款金额/本位币 有值，其余全空
    last_r = ws1.max_row
    if last_r >= 2:
        last_vals = {c: ws1.cell(last_r, c).value for c in range(1, ws1.max_column + 1)}
        non_empty = {c for c, v in last_vals.items() if v is not None and str(v).strip() != ''}
        key_cols   = {col_金额, col_本位币}
        # 如果非空列都在 key_cols 内，且至少有一个 key_col 有数据 → 合计行
        if non_empty and non_empty.issubset(key_cols):
            ws1.delete_rows(last_r)

    # 收集数据行（跳过表头第 1 行）
    data_rows = []  # [(收款金额, 备注), ...]
    for r in range(2, ws1.max_row + 1):
        v_金额 = ws1.cell(r, col_金额).value
        v_备注 = ws1.cell(r, col_备注).value
        data_rows.append((v_金额, v_备注))

    # ── 创建临时工作表"资金核对" ──────────────────────────────────────────────
    if '资金核对' in wb1.sheetnames:
        del wb1['资金核对']
    ws_tmp = wb1.create_sheet('资金核对')
    ws_tmp['A1'] = '收款金额'
    ws_tmp['B1'] = '备注'
    ws_tmp['E1'] = '贷方发生额'
    ws_tmp['F1'] = '匹配收款金额'

    for i, (金额, 备注) in enumerate(data_rows, start=2):
        ws_tmp.cell(i, 1, 金额)
        ws_tmp.cell(i, 2, 备注)

    progress(2, 5, '正在读取中信对公系统表…')

    # ── Step 2：读取中信对公"系统"工作表并按日期筛选 ──────────────────────────
    wb2 = openpyxl.load_workbook(file2_path)
    if '系统' not in wb2.sheetnames:
        raise ValueError(f'中信对公文件中未找到"系统"工作表，当前工作表：{wb2.sheetnames}')
    ws_sys = wb2['系统']

    headers2 = [ws_sys.cell(1, c).value for c in range(1, ws_sys.max_column + 1)]
    col_日期  = _find_col(headers2, '交易日期')
    col_贷方  = _find_col(headers2, '贷方发生额')

    # 筛选日期行，记录 (行号, 贷方发生额)
    filtered_rows: list[tuple[int, object]] = []
    for r in range(2, ws_sys.max_row + 1):
        cell_date = ws_sys.cell(r, col_日期).value
        # 支持 datetime / date / 字符串 三种格式
        row_date = None
        if cell_date is None:
            continue
        if hasattr(cell_date, 'date'):           # datetime
            row_date = cell_date.date()
        elif isinstance(cell_date, date):         # date
            row_date = cell_date
        else:                                     # 字符串 "2026-09-10" 或 "2026/09/10"
            s = str(cell_date).strip().replace('/', '-')
            try:
                row_date = date.fromisoformat(s[:10])
            except ValueError:
                continue
        if row_date == target_date:
            val = ws_sys.cell(r, col_贷方).value
            filtered_rows.append((r, val))

    if not filtered_rows:
        raise ValueError(
            f'在中信对公"系统"表的"交易日期"列中未找到 {target_date} 的数据，'
            '请确认日期格式与表后缀一致'
        )

    # 写入 E 列
    for i, (_, val) in enumerate(filtered_rows, start=2):
        ws_tmp.cell(i, 5, val)

    progress(3, 5, '正在匹配金额并标注颜色…')

    # ── Step 3：匹配与标色 ─────────────────────────────────────────────────────
    # 建立 A 列金额 → [行号] 映射（数值比较）
    a_map: dict[float, list[int]] = {}
    for r in range(2, len(data_rows) + 2):
        v = _to_num(ws_tmp.cell(r, 1).value)
        if v is not None:
            a_map.setdefault(v, []).append(r)

    # 已匹配的 A 列行号（防止重复标色）
    a_marked: set[int] = set()
    # 已匹配到的金额值（用于在系统表标色）
    matched_amounts: list[float] = []

    e_count = len(filtered_rows)
    for i in range(e_count):
        e_row = i + 2           # 临时表 E 列所在行
        e_val = _to_num(ws_tmp.cell(e_row, 5).value)

        if e_val is not None and e_val in a_map:
            # 找到匹配 → F 列填值
            ws_tmp.cell(e_row, 6, e_val)
            matched_amounts.append(e_val)
            # A 列对应第一个未标色行标浅红
            for a_row in a_map[e_val]:
                if a_row not in a_marked:
                    ws_tmp.cell(a_row, 1).fill = _FILL_LIGHT_RED
                    a_marked.add(a_row)
                    break
        else:
            # 未找到 → F 列标黄
            ws_tmp.cell(e_row, 6).fill = _FILL_YELLOW

    # 在系统表中，仅在筛选行内，将匹配金额标黄
    matched_set = set(matched_amounts)
    # 用集合记录已标色的 (row, col) 防重
    sys_marked: set[int] = set()
    for sys_row, sys_贷方 in filtered_rows:
        v = _to_num(sys_贷方)
        if v is not None and v in matched_set and sys_row not in sys_marked:
            ws_sys.cell(sys_row, col_贷方).fill = _FILL_YELLOW
            sys_marked.add(sys_row)

    progress(4, 5, '正在保存结果文件…')

    # ── Step 4：保存 ──────────────────────────────────────────────────────────
    out_name = f'资金核对结果-{date_suffix}.xlsx'
    out_path = os.path.join(out_dir, out_name)
    # 将中信对公的标色结果嵌入同一文件：以附加工作表方式复制
    # 方案：将 ws_sys 的标色信息写回 wb2 然后保存为另一个文件也可以；
    # 但业务上需要一个文件同时包含两张表，这里选择：
    #   收款审核表工作簿（含临时表）为主文件，
    #   并将"系统"标色后的数据追加为"中信对公-系统"工作表
    if '中信对公-系统' in wb1.sheetnames:
        del wb1['中信对公-系统']
    ws_copy = wb1.create_sheet('中信对公-系统')

    # 复制系统表所有数据及标色
    for r_idx, row in enumerate(ws_sys.iter_rows(), start=1):
        for c_idx, cell in enumerate(row, start=1):
            new_cell = ws_copy.cell(r_idx, c_idx, cell.value)
            if cell.fill and cell.fill.fill_type == 'solid':
                new_cell.fill = PatternFill(
                    fill_type='solid',
                    fgColor=cell.fill.fgColor.rgb if cell.fill.fgColor else 'FFFFFF'
                )

    wb1.save(out_path)

    progress(5, 5, f'处理完成，输出：{out_name}')
    return out_path
