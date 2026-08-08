# -*- coding: utf-8 -*-
"""
多表合并透视工具（python-Excel）

功能：
1. 支持一次粘贴/拖入多个 Excel 表格（*.xlsx / *.xlsm），回车确认；
   - 资源管理器多选后 Ctrl+C，再到终端 Ctrl+V 多行粘贴（每行一个路径）；
   - 也可输入 c 从剪贴板直接读取全部路径（参考 pdf-v-photo/pdf转图片.py）。
2. 列出全部表名（合并去重）与全部列名（合并去重），选择要合并的列；
   子表列名所在行与工作表序号可配置（默认第 1 行、第 1 个工作表）。
3. 按粘贴顺序追加合并所选列，每行末尾新增“数据来源表”列。
4. 透视汇总：
   - 固定模式：按 运单号、币种 分组，对 运费、附加费1/2/3、报关费、合计金额 求和；
   - 自定义模式：自由选择行标签列与求和数值列。
5. 透视表去向：
   - 直接保存为新的 xlsx 文件；
   - 或插入粘贴的主表中作为新工作表，并按“运单号”回填透视表“合计金额”列，
     再新增一列“透视表合计金额 - 主表汇总金额”的差异列；
     主表列名从第 3 行开始（可输入调整），“汇总金额”一般位于 AH 列。

使用：
    python 多表合并透视.py
"""

import os
import re
import sys
from collections import OrderedDict
from datetime import datetime

try:
    from openpyxl import Workbook, load_workbook
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


# ---------------------------------------------------------------------------
# 输入工具（参考 pdf-v-photo/pdf转图片.py）
# ---------------------------------------------------------------------------
# 匹配拖入路径：优先引号包裹（路径可含空格），否则连续非空白
PATH_RE = re.compile(r'"([^"]*)"|(\S+)')

# 预读缓冲：多行粘贴时，非路径行先暂存，让下一个问题优先读取
_PUTBACK = []


def _ask(prompt):
    """读一行；若预读缓冲有内容则优先取缓冲，否则阻塞等待输入。"""
    if _PUTBACK:
        return _PUTBACK.pop(0)
    return input(prompt)


def is_path_like(s):
    """粗略判断一行是否为文件路径（存在该文件，或含路径分隔符/盘符）。"""
    if not s:
        return False
    if os.path.isfile(s):
        return True
    if '/' in s or '\\' in s or ':' in s:
        return True
    return False


def split_line_paths(line):
    """把一行拆成多个路径（引号包裹优先，否则连续非空白）。"""
    parts = []
    for m in PATH_RE.finditer(line):
        p = m.group(1) if m.group(1) is not None else m.group(2)
        p = p.strip()
        if p:
            parts.append(p)
    return parts


def read_paths(prompt):
    """读取一行输入，按 Windows 拖拽/粘贴形态拆分成多个路径。
    兼容多种形态：
    - 单行单个/多个：d:/a.xlsx、"d:/a.xlsx" "d:/b c.xlsx"、d:/a.xlsx d:/b.xlsx
    - 输入 c 时从剪贴板读取全部路径（资源管理器多选后 Ctrl+C）
    - Ctrl+V 多行粘贴：每行一个路径，自动连续读取；非路径行暂存给下一个问题
    """
    first = _ask(prompt).strip()
    if first.lower() in ('c', 'cb', 'clip', '粘贴'):
        return read_clipboard_paths()
    paths = []
    if not first:
        return []
    if not is_path_like(first):
        return []  # 非路径，交由上层重试
    if os.path.isfile(first):
        paths.append(first)
    else:
        paths.extend(split_line_paths(first))
    # 继续读取后续路径行（支持 Ctrl+V 多行粘贴）
    while True:
        nxt = _ask('').strip()
        if not nxt:
            break
        if is_path_like(nxt):
            if os.path.isfile(nxt):
                paths.append(nxt)
            else:
                paths.extend(split_line_paths(nxt))
        else:
            _PUTBACK.append(nxt)
            break
    return paths


def read_clipboard_paths():
    """从剪贴板读取文件路径列表（资源管理器多选后 Ctrl+C 时，每行一个引号路径）。"""
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.update()
        try:
            data = root.clipboard_get()
        except tk.TclError:
            return []
        finally:
            root.destroy()
        paths = []
        for line in data.splitlines():
            line = line.strip()
            if not line:
                continue
            paths.extend(split_line_paths(line))
        return paths
    except Exception:
        return []


def read_save_path(prompt, default_name):
    """读取保存路径：支持直接输入、拖入文件夹/文件（自动去掉 Windows 拖入时带的引号）。
    空输入用默认名；拖入文件夹则在其下生成默认文件名；无扩展名补 .xlsx。"""
    s = _ask(prompt).strip()
    if len(s) >= 2 and s[0] == s[-1] == '"':
        s = s[1:-1].strip()
    if not s:
        return default_name
    if os.path.isdir(s):
        return os.path.join(s, default_name)
    if not s.lower().endswith(('.xlsx', '.xlsm')):
        s += '.xlsx'
    return s


# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------
def parse_multi_choice(text, n):
    """把 '1,3,5-7' 解析为 0 起索引列表（去重、校验范围）；支持 all/全部/*。"""
    t = text.strip()
    if not t:
        return []
    if t.lower() in ('all', '全部', '*'):
        return list(range(n))
    idxs = []
    for part in re.split(r'[,\s，、]+', t):
        if not part:
            continue
        if '-' in part or '~' in part:
            a, b = re.split(r'[-~]', part, maxsplit=1)
            try:
                a, b = int(a), int(b)
            except ValueError:
                raise ValueError('无效范围：%s' % part)
            if a > b:
                a, b = b, a
            if a < 1 or b > n:
                raise ValueError('序号超出范围：%s' % part)
            idxs.extend(range(a - 1, b))
        else:
            try:
                i = int(part)
            except ValueError:
                raise ValueError('无效序号：%s' % part)
            if i < 1 or i > n:
                raise ValueError('序号超出范围：%s' % part)
            idxs.append(i - 1)
    seen = set()
    out = []
    for i in idxs:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def sanitize_sheet_name(name):
    """去掉工作表名中的非法字符。"""
    name = re.sub(r'[\\/:*?"<>|\r\n]', '_', str(name)).strip()
    return name or '透视汇总'


def to_number(v):
    """把单元格值转成可求和的数值，无法解析的按 0 处理。"""
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '').replace('，', '')
    s = re.sub(r'[¥￥$€\s]', '', s)
    s = s.rstrip('元')
    if s in ('', '-', '--', '/', 'N/A', 'NA'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def norm_key(v):
    """把值规范成用于匹配的文本键（运单号匹配用，兼容数字/字符串）。"""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return repr(v)
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


def find_column(headers, candidates):
    """在表头列表中找列，返回 0 起列索引或 None；优先精确匹配，再匹配包含。"""
    cands = [c for c in candidates if c]
    for c in cands:
        if c in headers:
            return headers.index(c)
    for i, h in enumerate(headers):
        hs = str(h).strip()
        for c in cands:
            if c in hs:
                return i
    return None


def find_col_in_table(headers, name, aliases):
    """在子表表头中找列：精确匹配所选名 → 精确匹配替代名 → 模糊包含匹配。
    返回 0 起列索引或 None。"""
    h = {str(x).strip(): i for i, x in enumerate(headers)}
    for cand in [name] + list(aliases):
        if cand in h:
            return h[cand]
    for cand in [name] + list(aliases):
        for i, x in enumerate(headers):
            xs = str(x).strip()
            if cand in xs:
                return i
    return None


# ---------------------------------------------------------------------------
# 读取表格
# ---------------------------------------------------------------------------
def read_table(path, sheet_name=None, sheet_idx=0, header_row=1):
    """读取 Excel 的指定工作表，返回 (工作表名, 表头列表, 数据行列表)。
    sheet_name 优先：按工作表名称取（所有子表统一用选中的工作表合并）；
    未指定时按 sheet_idx（0 起，默认第 1 个）取第几个工作表。
    header_row 为列名所在行（1 起，默认第 1 行），数据从 header_row+1 行开始。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                raise ValueError('该文件没有工作表"%s"' % sheet_name)
            ws = wb[sheet_name]
        else:
            ws = wb.worksheets[sheet_idx]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if len(rows) < header_row:
        return ws.title, [], []
    header_cells = rows[header_row - 1]
    headers = []
    for i, v in enumerate(header_cells):
        name = v if isinstance(v, str) else ('' if v is None else str(v))
        name = name.strip()
        if name == '':
            name = '未命名列%d' % (i + 1)
        headers.append(name)
    data = [r for r in rows[header_row:] if any(v is not None and str(v).strip() != '' for v in r)]
    return ws.title, headers, data


# ---------------------------------------------------------------------------
# 透视汇总
# ---------------------------------------------------------------------------
FIXED_ROW_LABELS = ['运单号', '币种']
FIXED_VALUE_BASES = ['运费', '附加费2（如超围长重附加费）',
                     '附加费1（如偏远费）', '附加费3（如超长重附加费）',
                     '报关费', '合计金额']
# 合并时用于自动生成“汇总”列的单项费用（子串匹配）
FEE_BASES = ('运费', '附加费1', '附加费2', '附加费3', '报关费')
# 固定数值列在子表/汇总表中的常见别名（优先精确，再按别名/包含匹配）
VALUE_ALIASES = {
    '合计金额': ['费用合计', '合计', '费用总计', '汇总'],
}


def resolve_columns(headers, names):
    """把固定列名解析到汇总表实际列名：
    优先精确匹配；否则匹配 '求和项:'+名称；再按别名/包含匹配；找不到返回 None。"""
    result = []
    for name in names:
        if name in headers:
            result.append(name)
            continue
        if ('求和项:%s' % name) in headers:
            result.append('求和项:%s' % name)
            continue
        found = None
        for cand in [name] + VALUE_ALIASES.get(name, []):
            for h in headers:
                if cand in str(h):
                    found = h
                    break
            if found:
                break
        result.append(found)
    return result


def make_pivot(headers, rows, row_labels, value_cols):
    """按行标签分组，对数值列求和，返回 (透视表头, 透视表行)。"""
    idx = {name: i for i, name in enumerate(headers)}
    missing = [c for c in row_labels + value_cols if c not in idx]
    if missing:
        raise KeyError('汇总表中不存在列：%s' % '、'.join(missing))
    rl_idx = [idx[x] for x in row_labels]
    v_idx = [idx[x] for x in value_cols]
    groups = OrderedDict()
    for r in rows:
        key = tuple(r[i] if i < len(r) else None for i in rl_idx)
        if all(norm_key(k) == '' for k in key):
            continue  # 行标签全空的行跳过
        if key not in groups:
            groups[key] = [to_number(r[i] if i < len(r) else None) for i in v_idx]
        else:
            sums = groups[key]
            for j, i in enumerate(v_idx):
                sums[j] += to_number(r[i] if i < len(r) else None)
    pivot_headers = list(row_labels) + ['求和项:%s' % c for c in value_cols]
    pivot_rows = [list(key) + sums for key, sums in groups.items()]
    return pivot_headers, pivot_rows


# ---------------------------------------------------------------------------
# 写表
# ---------------------------------------------------------------------------
def write_table(path, sheet_title, headers, rows):
    """把表头+数据行写入新的 xlsx 文件。"""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(path)


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 回填透视合计金额 + 差异列
# ---------------------------------------------------------------------------
def backfill_pivot_total(wb, sheet_name, pivot_headers, pivot_rows, header_row=3):
    """在指定工作表中按运单号回填透视表合计金额，并新增差异列。
    主表列名位于 header_row 行（默认第 3 行），数据从 header_row+1 行开始；
    “汇总金额”列优先取 AH 列（第 34 列），否则按列名查找。
    返回 (差异行数, 透视合计列名, 差异列名)。"""
    ws = wb[sheet_name]
    headers = [ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]
    headers_str = ['' if v is None else str(v).strip() for v in headers]

    yd_idx = find_column(headers_str, ['运单号'])
    if yd_idx is None:
        raise ValueError('所选工作表未找到“运单号”列，无法回填。')
    hz_idx = None
    if ws.max_column >= 34:
        ah_val = ws.cell(row=header_row, column=34).value
        if ah_val is not None and '汇总金额' in str(ah_val):
            hz_idx = 33
    if hz_idx is None:
        hz_idx = find_column(headers_str, ['汇总金额'])
    if hz_idx is None:
        raise ValueError('所选工作表未找到“汇总金额”列，无法计算差异。')

    p_headers_str = ['' if v is None else str(v).strip() for v in pivot_headers]
    p_yd_idx = find_column(p_headers_str, ['运单号'])
    if p_yd_idx is None:
        raise ValueError('透视表中未找到“运单号”列，无法回填。')
    p_hz_idx = find_column(p_headers_str, ['求和项:合计金额', '合计金额'])
    if p_hz_idx is None:
        raise ValueError('透视表中未找到“合计金额”列，无法回填。')

    # 透视表按运单号建索引（重复取第一个）
    pivot_map = {}
    dup = 0
    for row in pivot_rows:
        k = norm_key(row[p_yd_idx] if p_yd_idx < len(row) else None)
        if not k:
            continue
        if k in pivot_map:
            dup += 1
            continue
        pivot_map[k] = row[p_hz_idx] if p_hz_idx < len(row) else None
    if dup:
        print('提示：透视表中有 %d 个重复运单号（如不同币种），回填时使用第一个匹配值。' % dup)

    # 已存在回填列则复用，否则在末尾新增
    pf_exist = find_column(headers_str, ['透视表合计金额'])
    diff_exist = find_column(headers_str, ['合计金额差异（透视-原表）'])
    if pf_exist is not None and diff_exist is not None:
        col_pf, col_diff = pf_exist + 1, diff_exist + 1
    else:
        col_pf = ws.max_column + 1
        col_diff = col_pf + 1
        ws.cell(row=header_row, column=col_pf, value='透视表合计金额')
        ws.cell(row=header_row, column=col_diff, value='合计金额差异（透视-原表）')

    filled = 0
    for r in range(header_row + 1, ws.max_row + 1):
        key = norm_key(ws.cell(row=r, column=yd_idx + 1).value)
        pf = pivot_map.get(key)
        ws.cell(row=r, column=col_pf, value=pf)
        if pf is None or norm_key(pf) == '':
            continue  # 透视表无此运单号，差异留空
        orig = ws.cell(row=r, column=hz_idx + 1).value
        if orig is None or norm_key(orig) == '':
            continue
        ws.cell(row=r, column=col_diff, value=to_number(pf) - to_number(orig))
        filled += 1
    return filled, '透视表合计金额', '合计金额差异（透视-原表）'


# 主流程
# ---------------------------------------------------------------------------
def main():
    if not OPENPYXL_OK:
        print('提示：未安装 openpyxl，无法读写 Excel。请运行：pip install openpyxl')
        return

    print('多表合并透视工具')
    print('功能：粘贴多表 -> 选择列 -> 合并 -> 透视汇总 -> 保存或插入主表')
    print('-' * 64)

    # ---- 1. 粘贴/读取表格 ----
    while True:
        paths = read_paths('请粘贴要合并的表格文件（可多选复制后 Ctrl+V 粘贴，或输入 c 从剪贴板读取），粘贴后回车：')
        if not paths:
            print('未检测到文件路径，请重新粘贴。')
            continue
        valid, seen = [], set()
        for p in paths:
            ap = os.path.normcase(os.path.abspath(p))
            if ap in seen:
                continue
            seen.add(ap)
            if not os.path.isfile(p):
                print('跳过（找不到文件）：%s' % p)
                continue
            if not (p.lower().endswith('.xlsx') or p.lower().endswith('.xlsm')):
                print('跳过（仅支持 .xlsx/.xlsm）：%s' % p)
                continue
            valid.append(p)
        if not valid:
            print('没有可用的表格文件，请重新粘贴。')
            continue
        break

    # 立即反馈：确认已识别的表格（避免“读了没反应”的困惑）
    if len(valid) == 1:
        print('已识别 1 个表格：%s' % os.path.basename(valid[0]))
    else:
        print('已识别 %d 个表格（顺序即合并顺序）：' % len(valid))
        for i, p in enumerate(valid, start=1):
            print('  %d. %s' % (i, os.path.basename(p)))

    # ---- 1.5 选择要合并的工作表：识别全部文件的工作表名称，合并去重后选择 ----
    all_sheets, seen_sheets = [], set()
    for p in valid:
        try:
            wb = load_workbook(p, read_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception as e:
            print('跳过（无法读取工作表列表）：%s（%s）' % (os.path.basename(p), e))
            continue
        for s in sheets:
            s = s.strip()
            if s and s not in seen_sheets:
                seen_sheets.add(s)
                all_sheets.append(s)
    if not all_sheets:
        print('无法读取任何工作表名称，请检查表格文件。')
        return
    print('全部工作表名称（合并去重）：')
    for i, s in enumerate(all_sheets, start=1):
        print('  %d. %s' % (i, s))
    while True:
        sel = _ask('请选择要合并的工作表（输入序号或工作表名称，回车默认 1）：').strip()
        if not sel:
            sheet_name = all_sheets[0]
            break
        if sel.isdigit():
            idx = int(sel)
            if 1 <= idx <= len(all_sheets):
                sheet_name = all_sheets[idx - 1]
                break
            print('序号超出范围（1-%d）。' % len(all_sheets))
            continue
        if sel in seen_sheets:
            sheet_name = sel
            break
        print('不存在工作表名称：%s' % sel)
    print('已选择工作表：%s（全部表格都使用该工作表的数据合并）' % sheet_name)

    # 列名默认在第 1 行读取
    header_row = 1

    tables = []
    for p in valid:
        try:
            sheet_name, headers, rows = read_table(p, sheet_name=sheet_name, header_row=header_row)
        except Exception as e:
            print('读取失败，跳过：%s（%s）' % (os.path.basename(p), e))
            continue
        tables.append({'path': p, 'base': os.path.basename(p),
                       'sheet': sheet_name, 'headers': headers, 'rows': rows})
    if not tables:
        print('没有成功读取任何表格。')
        return

    # 显示名：同名不同目录时附加目录
    name_counts = {}
    for t in tables:
        name_counts[t['base']] = name_counts.get(t['base'], 0) + 1
    for t in tables:
        if name_counts[t['base']] > 1:
            t['display'] = '%s（%s）' % (t['base'], os.path.dirname(os.path.abspath(t['path'])))
        else:
            t['display'] = t['base']

    print('共读取 %d 个表格（顺序即合并顺序）：' % len(tables))
    for i, t in enumerate(tables, start=1):
        print('  %d. %s（工作表：%s，%d 行数据）' % (i, t['display'], t['sheet'], len(t['rows'])))

    # ---- 2. 列选择 ----
    all_cols = []
    seen_cols = set()
    for t in tables:
        for h in t['headers']:
            if h not in seen_cols:
                seen_cols.add(h)
                all_cols.append(h)
    if not all_cols:
        print('表格中没有可用的列名，无法合并。')
        return
    print('全部列名（合并去重）：')
    for i, c in enumerate(all_cols, start=1):
        print('  %d. %s' % (i, c))
    while True:
        print('全部列名（合并去重）：')
        for i, c in enumerate(all_cols, start=1):
            print('  %d. %s' % (i, c))
        ans = _ask('请选择要合并的列（多个用逗号/空格分隔，如 1,3,5-7；输入 all 全选）：')
        try:
            sel_idx = parse_multi_choice(ans, len(all_cols))
        except ValueError as e:
            print('选择无效：%s，请重新输入。' % e)
            continue
        if not sel_idx:
            print('至少选择一列。')
            continue
        sel_cols = [all_cols[i] for i in sel_idx]

        # 检查哪些表缺少所选列（精确匹配）
        missing_tables = {}
        for t in tables:
            hidx = {h: i for i, h in enumerate(t['headers'])}
            miss = [c for c in sel_cols if c not in hidx]
            if miss:
                missing_tables[t['display']] = miss
        if missing_tables:
            for disp, miss in missing_tables.items():
                print('提示：表 %s 缺少列 %s，这些列在该表留空。' % (disp, '、'.join(miss)))
            if _ask('检测到有表格缺少所选列，是否重新选择列？(y/n，回车默认 y)：').strip().lower() in ('', 'y', 'yes'):
                continue  # 重新给一次选择
        break
    print('已选择 %d 列：%s' % (len(sel_cols), '、'.join(sel_cols)))

    # 列名替代映射：所选列在部分子表里列名不同，可指定替代列名匹配
    col_aliases = {}
    rename_ans = _ask('是否要修改某个列的名称（为所选列指定替代列名，匹配子表中列名不同的数据）？\n'
                      '  y/n，回车默认 n；也可直接输入要改名的列名或序号，直接进入改名：').strip()
    first_pick = None
    if rename_ans.lower() in ('', 'n', 'no', '否'):
        pass  # 不改名
    elif rename_ans.isdigit():
        idx = int(rename_ans)
        if 1 <= idx <= len(sel_cols):
            first_pick = sel_cols[idx - 1]
        else:
            print('序号超出范围（1-%d），进入改名流程。' % len(sel_cols))
    elif rename_ans.lower() in ('y', 'yes', '是'):
        pass  # 纯 yes，正常进入
    elif rename_ans in sel_cols:
        first_pick = rename_ans
    else:
        print('未识别输入（%s），进入改名流程。' % rename_ans)

    if first_pick is not None or rename_ans.lower() not in ('', 'n', 'no', '否'):
        while True:
            if first_pick is not None:
                sel_name = first_pick
                first_pick = None
            else:
                print('当前所选列：')
                for i, c in enumerate(sel_cols, start=1):
                    print('  %d. %s' % (i, c))
                sel_name = _ask('请选择要修改名称的列（输入序号或列名；输入 done 结束）：').strip()
            if not sel_name or sel_name.lower() == 'done':
                break
            if sel_name.isdigit():
                idx = int(sel_name)
                if 1 <= idx <= len(sel_cols):
                    col = sel_cols[idx - 1]
                else:
                    print('序号超出范围（1-%d）。' % len(sel_cols))
                    continue
            elif sel_name in sel_cols:
                col = sel_name
            else:
                print('所选列中不存在：%s' % sel_name)
                continue
            aliases = _ask('请输入 %s 的替代列名（子表实际使用的列名，多个用逗号/空格分隔）：' % col).strip()
            alias_list = [a.strip() for a in re.split(r'[,\s，、]+', aliases) if a.strip()]
            if not alias_list:
                print('未输入有效列名，跳过。')
                continue
            col_aliases[col] = alias_list
            print('已设置：%s -> %s' % (col, '、'.join(alias_list)))
        if col_aliases:
            print('列名替代映射：' + '；'.join('%s -> %s' % (k, '、'.join(v)) for k, v in col_aliases.items()))

    # ---- 3. 追加合并 ----
    # 自动生成“汇总”列：运费+附加费1/2/3+报关费，放在最后一个费用列后面
    fee_cols = [c for c in sel_cols if any(base in c for base in FEE_BASES)]
    merged_headers = list(sel_cols)
    sum_pos = None
    if fee_cols:
        sum_pos = max(sel_cols.index(c) for c in fee_cols) + 1
        merged_headers.insert(sum_pos, '汇总')
    merged_headers.append('数据来源表')
    merged_rows = []
    for t in tables:
        hidx_map = {}
        missing = []
        for c in sel_cols:
            i = find_col_in_table(t['headers'], c, col_aliases.get(c, []))
            hidx_map[c] = i
            if i is None:
                missing.append(c)
        if missing:
            print('提示：表 %s 缺少列 %s，这些列在该表留空。' % (t['display'], '、'.join(missing)))
        for r in t['rows']:
            row = []
            for c in sel_cols:
                i = hidx_map.get(c)
                row.append(r[i] if i is not None and i < len(r) else None)
            if sum_pos is not None:
                total = 0.0
                has_val = False
                for c in fee_cols:
                    i = hidx_map.get(c)
                    v = r[i] if i is not None and i < len(r) else None
                    if v is not None and str(v).strip() != '':
                        total += to_number(v)
                        has_val = True
                row.insert(sum_pos, total if has_val else None)
            row.append(t['display'])
            merged_rows.append(row)
    print('合并完成：共 %d 行，%d 列（含“汇总”“数据来源表”）。' % (len(merged_rows), len(merged_headers)))

    # 可选：保存合并汇总表
    if _ask('是否保存合并后的汇总表？(y/n，回车默认 n)：').strip().lower() in ('y', 'yes'):
        out = read_save_path('请输入保存路径（可拖入文件夹；回车默认：合并汇总.xlsx）：', '合并汇总.xlsx')
        try:
            write_table(out, '合并汇总', merged_headers, merged_rows)
            print('已保存合并汇总表：%s' % out)
        except Exception as e:
            print('保存合并汇总表失败：%s' % e)

    # ---- 4. 透视模式 ----
    print('-' * 64)
    print('合并后汇总表列（含“数据来源表”）：')
    for i, c in enumerate(merged_headers, start=1):
        print('  %d. %s' % (i, c))
    while True:
        mode = _ask('请选择透视模式：1 固定模式（运单号、币种分组，对运费/附加费1/2/3、报关费、合计金额求和）| 2 自定义模式。输入 1 或 2：').strip()
        if mode == '1':
            rl = resolve_columns(merged_headers, FIXED_ROW_LABELS)
            if any(x is None for x in rl):
                miss = [FIXED_ROW_LABELS[i] for i, x in enumerate(rl) if x is None]
                print('汇总表中缺少固定模式的分组列：%s（上面列表里没有这些列），请改用自定义模式。' % '、'.join(miss))
                continue
            vals = resolve_columns(merged_headers, FIXED_VALUE_BASES)
            missing_vals = [FIXED_VALUE_BASES[i] for i, x in enumerate(vals) if x is None]
            if missing_vals:
                print('提示：汇总表中缺少数值列 %s，透视时跳过这些列。' % '、'.join(missing_vals))
            value_cols = [x for x in vals if x is not None]
            if not value_cols:
                print('汇总表中没有可求和的数值列，无法透视。')
                continue
            row_labels, value_cols = rl, value_cols
            break
        elif mode == '2':
            print('汇总表全部列：')
            for i, c in enumerate(merged_headers, start=1):
                print('  %d. %s' % (i, c))
            while True:
                try:
                    rl_idx = parse_multi_choice(
                        _ask('请选择行标签列（分组列，多个用逗号/空格分隔，如 1,2）：'),
                        len(merged_headers))
                    v_idx = parse_multi_choice(
                        _ask('请选择求和数值列（多个用逗号/空格分隔，如 3,4,5）：'),
                        len(merged_headers))
                except ValueError as e:
                    print('选择无效：%s，请重新输入。' % e)
                    continue
                if not rl_idx or not v_idx:
                    print('行标签列和数值列都至少选一个。')
                    continue
                break
            row_labels = [merged_headers[i] for i in rl_idx]
            value_cols = [merged_headers[i] for i in v_idx]
            break
        else:
            print('请输入 1 或 2。')

    try:
        pivot_headers, pivot_rows = make_pivot(merged_headers, merged_rows, row_labels, value_cols)
    except KeyError as e:
        print('透视失败：%s' % e)
        return
    print('透视完成：共 %d 行 × %d 列。' % (len(pivot_rows), len(pivot_headers)))
    print('透视表列：%s' % '、'.join(pivot_headers))
    for i, r in enumerate(pivot_rows[:5], start=1):
        print('  %d. %s' % (i, ' | '.join('' if v is None else str(v) for v in r)))
    if len(pivot_rows) > 5:
        print('  ...（共 %d 行）' % len(pivot_rows))

    # ---- 5. 透视表去向 ----
    print('-' * 64)
    while True:
        dest = _ask('请选择透视表去向：1 直接保存为文件 | 2 插入到已有的主表（Excel）中。输入 1 或 2：').strip()
        if dest == '1':
            out = read_save_path('请输入保存路径（可拖入文件夹；回车默认当前目录 透视汇总.xlsx）：', '透视汇总.xlsx')
            if os.path.exists(out):
                base, ext = os.path.splitext(out)
                out = '%s_%s%s' % (base, datetime.now().strftime('%Y%m%d_%H%M%S'), ext)
            try:
                write_table(out, '透视汇总', pivot_headers, pivot_rows)
                print('已保存透视表：%s' % out)
            except Exception as e:
                print('保存透视表失败：%s' % e)
            break
        elif dest == '2':
            while True:
                mains = read_paths('请粘贴要插入的主表（*.xlsx），粘贴后回车：')
                if not mains:
                    print('未检测到文件路径，请重新粘贴。')
                    continue
                main_path = mains[0]
                if not os.path.isfile(main_path):
                    print('找不到文件：%s' % main_path)
                    continue
                if not (main_path.lower().endswith('.xlsx') or main_path.lower().endswith('.xlsm')):
                    print('仅支持 .xlsx/.xlsm 主表，请重新粘贴。')
                    continue
                break
            try:
                wb = load_workbook(main_path)
            except Exception as e:
                print('打开主表失败：%s' % e)
                return

            sheet_name = sanitize_sheet_name(
                _ask('请输入透视表工作表名称（回车默认：透视汇总）：').strip() or '透视汇总')
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            pws = wb.create_sheet(title=sheet_name)
            pws.append(pivot_headers)
            for row in pivot_rows:
                pws.append(row)
            print('已插入透视表工作表：%s' % sheet_name)

            print('主表 %s 的全部工作表：' % os.path.basename(main_path))
            for i, name in enumerate(wb.sheetnames, start=1):
                print('  %d. %s' % (i, name))
            while True:
                sel = _ask('请选择要回填合计金额的工作表（输入序号，回车默认 1）：').strip() or '1'
                try:
                    target = wb.sheetnames[int(sel) - 1]
                    break
                except (ValueError, IndexError):
                    print('请输入有效序号（1-%d）。' % len(wb.sheetnames))
            print('已选择工作表：%s' % target)

            while True:
                hdr_in = _ask('请输入主表列名所在行（该主表列名从第 3 行开始，回车默认 3）：').strip()
                if not hdr_in:
                    header_row = 3
                    break
                if hdr_in.isdigit() and int(hdr_in) >= 1:
                    header_row = int(hdr_in)
                    break
                print('请输入有效的行号（正整数）。')
            try:
                filled, f_col, d_col = backfill_pivot_total(wb, target, pivot_headers, pivot_rows, header_row)
            except Exception as e:
                print('回填失败：%s' % e)
                return
            wb.save(main_path)
            print('已新增列：%s、%s，共计算 %d 行合计金额差异。' % (f_col, d_col, filled))
            print('已保存主表：%s' % main_path)
            break
        else:
            print('请输入 1 或 2。')

    print('已完成，谢谢使用。')
    print('玛卡巴卡""')


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('\n用户取消操作。')
        print('玛卡巴卡""')
    sys.exit(0)