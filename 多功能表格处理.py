from __future__ import annotations

import os
import sys
import time
import traceback
from collections import OrderedDict
from pathlib import Path
import re
import subprocess
from typing import Iterable, Sequence

try:
    import pythoncom
    import pywintypes
    import win32com.client as win32
    import win32process
except ModuleNotFoundError as exc:
    missing_name = exc.name or "pywin32"
    print(
        "缺少运行依赖："
        f"{missing_name}\n"
        "请先执行：python -m pip install pywin32"
    )
    raise SystemExit(1)


XL_CALCULATION_AUTOMATIC = -4105
XL_PASTE_VALUES = -4163
XL_PASTE_VALUES_AND_NUMBER_FORMATS = 12
XL_UP = -4162
EXCEL_BUSY_HRESULTS = {-2147418111, -2147417846}


PREFIX_MAP = OrderedDict(
    [
        ("ZSZK-", ("中山锦联", "深圳锦联", "华南")),
        ("ZSTH-", ("中山锦联", "深圳锦联", "华南")),
        ("ZS-", ("中山锦联", "深圳锦联", "华南")),
        ("XJ-", ("深圳操作", "深圳锦联", "华南")),
        ("SZ4-", ("深圳直客四部", "深圳锦联", "华南")),
        ("SZ3-", ("深圳直客三部", "深圳锦联", "华南")),
        ("SZ2-", ("深圳直客二部", "深圳锦联", "华南")),
        ("SZ1-", ("深圳直客一部", "深圳锦联", "华南")),
        ("ST2-", ("深圳同行二部", "深圳锦联", "华南")),
        ("ST1-", ("深圳同行一部", "深圳锦联", "华南")),
        ("GZ1-", ("广州直客一部", "广州锦联", "华南")),
        ("GT2-", ("广州同行二部", "广州锦联", "华南")),
        ("GT1-", ("广州同行一部", "广州锦联", "华南")),
        ("14-", ("销售十四部", "事业三部", "事业三部")),
        ("13-", ("销售十三部", "事业三部", "事业三部")),
        ("12-", ("销售十二部", "事业三部", "事业三部")),
        ("11-", ("销售十一部", "事业三部", "事业三部")),
        ("10-", ("销售十部", "事业三部", "事业三部")),
        ("9-", ("销售九部", "事业三部", "事业三部")),
        ("8-", ("销售八部", "事业三部", "事业三部")),
        ("7-", ("销售七部", "事业三部", "事业三部")),
        ("6-", ("销售六部", "事业一部", "事业一部")),
        ("5-", ("销售五部", "事业一部", "事业一部")),
        ("4-", ("销售四部", "事业一部", "事业一部")),
        ("3-", ("销售三部", "事业一部", "事业一部")),
        ("2-", ("销售二部", "事业一部", "事业一部")),
        ("1-", ("销售一部", "事业一部", "事业一部")),
    ]
)

PURE_NAME_DEPARTMENT_OVERRIDE_MAP = OrderedDict(
    [
        ("郑州", "郑州锦联"),
        ("宁波", "宁波锦联"),
        ("合肥", "合肥锦联"),
        ("厦门", "厦门锦联"),
        ("杭州", "杭州锦联"),
    ]
)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return ("%f" % value).rstrip("0").rstrip(".")
    return str(value).strip()


def preserve_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return ("%f" % value).rstrip("0").rstrip(".")
    return str(value)


def normalize_code_key(value: object) -> str:
    text = normalize_text(value).replace(",", "").replace("'", "")
    if not text:
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    if text.lstrip("-").isdigit():
        sign = "-" if text.startswith("-") else ""
        digits = text.lstrip("-").lstrip("0")
        return f"{sign}{digits or '0'}"
    return text.strip()


def to_excel_number(value: object) -> object:
    text = normalize_code_key(value)
    if not text:
        return None
    if text.lstrip("-").isdigit():
        return int(text)
    try:
        numeric = float(text)
    except ValueError:
        return value
    if numeric.is_integer():
        return int(numeric)
    return numeric


def normalize_department_name(value: object) -> str:
    text = normalize_text(value)
    if text.startswith("市场") and text.endswith("部"):
        return f"销售{text[2:]}"
    return text


def clean_dragged_path(raw: str) -> str:
    text = raw.strip()
    if not text:
        return ""

    if text.startswith("&"):
        text = text[1:].strip()

    quoted_match = re.fullmatch(r"""['"](.+?)['"]""", text)
    if quoted_match:
        text = quoted_match.group(1).strip()
    elif text.startswith('"') and text.endswith('"'):
        text = text[1:-1]

    return text.strip()


# 批量粘贴输入支持（参考 多表合并透视.py）：
# - 单行多个路径：引号包裹优先（路径可含空格），否则连续非空白分隔
# - Ctrl+V 多行粘贴：每行一个或多个路径，自动连续读取
# - 输入 c / cb / clip / 粘贴：直接从剪贴板读取全部路径（资源管理器多选后 Ctrl+C）
PATH_RE = re.compile(r'"([^"]*)"|(\S+)')

# 预读缓冲：多行粘贴时，非路径行先暂存，让下一个问题优先读取
_PUTBACK: list[str] = []


def _ask(prompt: str = "") -> str:
    """读一行；若预读缓冲有内容则优先取缓冲，否则阻塞等待输入。"""
    if _PUTBACK:
        return _PUTBACK.pop(0)
    return input(prompt)


def is_path_like(s: str) -> bool:
    """粗略判断一行是否为文件路径（存在该文件，或含路径分隔符/盘符）。"""
    if not s:
        return False
    if os.path.isfile(s):
        return True
    if "/" in s or "\\" in s or ":" in s:
        return True
    return False


def split_line_paths(line: str) -> list[str]:
    """把一行拆成多个路径（引号包裹优先，否则连续非空白）。"""
    parts: list[str] = []
    for m in PATH_RE.finditer(line):
        p = m.group(1) if m.group(1) is not None else m.group(2)
        p = p.strip()
        if p:
            parts.append(p)
    return parts


def _clipboard_text_with_pywin32() -> list[str] | None:
    """用 pywin32 从剪贴板读取文本格式的文件路径。失败返回 None，未读到返回 []。"""
    try:
        import win32clipboard
    except Exception:
        return None
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                return []
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return None
    if not data:
        return []
    paths: list[str] = []
    for line in str(data).splitlines():
        line = line.strip()
        if not line:
            continue
        paths.extend(split_line_paths(line))
    return paths


def read_clipboard_paths() -> list[str]:
    """从剪贴板读取文件路径列表（资源管理器多选后 Ctrl+C 时，每行一个引号路径）。

    脚本已依赖 pywin32，优先用 win32clipboard 读取；读不到时退回 tkinter。
    """
    paths = _clipboard_text_with_pywin32()
    if paths:
        return paths
    # 后备：tkinter 读取剪贴板文本
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


def read_paths_batch(prompt: str = "") -> list[str]:
    """批量读取文件路径，返回原始路径字符串列表（可能存在无效项）。

    兼容多种粘贴形态：
    - 单行单个/多个：d:/a.xlsx、"d:/a b.xlsx"、d:/a.xlsx d:/b.xlsx
    - 输入 c / cb / clip / 粘贴：从剪贴板读取全部路径
    - Ctrl+V 多行粘贴：每行一个或多个路径，自动连续读取；非路径行暂存给下一个问题
    遇到空行结束本轮读取。
    若预读缓冲顶部已有路径（如上一个问题暂存的多余文件），直接消费并返回，无需再等待输入。
    """
    if _PUTBACK and is_path_like(_PUTBACK[0]):
        buffered: list[str] = []
        while _PUTBACK and is_path_like(_PUTBACK[0]):
            buffered.append(_PUTBACK.pop(0))
        paths: list[str] = []
        for line in buffered:
            if os.path.isfile(line):
                paths.append(line)
            else:
                paths.extend(split_line_paths(line))
        return paths

    first = _ask(prompt).strip()
    while True:
        if first.lower() in ("c", "cb", "clip", "粘贴"):
            clipboard_paths = read_clipboard_paths()
            if clipboard_paths:
                return clipboard_paths
            print("剪贴板中未读取到文件路径：请先在资源管理器选中文件后 Ctrl+C 复制，")
            print("再输入 c；或直接粘贴文件路径。")
            first = _ask(prompt).strip()
            continue
        break
    paths: list[str] = []
    if not first:
        return []
    if not is_path_like(first):
        return []
    if os.path.isfile(first):
        paths.append(first)
    else:
        paths.extend(split_line_paths(first))
    while True:
        nxt = _ask("").strip()
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


def _collect_valid_paths(raw_list: list[str]) -> list[Path]:
    """清洗并校验路径，返回有效 Path 列表；无效的打印提示并忽略。"""
    valid: list[Path] = []
    for raw in raw_list:
        cleaned = clean_dragged_path(raw)
        if not cleaned:
            continue
        p = Path(cleaned)
        if p.exists():
            valid.append(p)
        else:
            print(f"文件不存在，已忽略：{p}")
    return valid


def _print_identified_paths(paths: Sequence[Path]) -> None:
    """打印当前已识别的文件列表，供用户确认。"""
    print(f"已识别 {len(paths)} 个文件（顺序即处理顺序）：")
    for i, p in enumerate(paths, 1):
        print(f"  {i}. {p.name}")


def ask_paths(expected_count: int) -> list[Path]:
    print("请按顺序粘贴基础文件路径。")
    print("支持一次粘贴多个：Ctrl+V 多行粘贴、单行多个路径，或输入 c 从剪贴板读取全部路径。")
    print("粘贴完成后按一次回车结束本轮。")
    paths: list[Path] = []
    while True:
        raw_list = read_paths_batch("> ")
        if not raw_list:
            if not paths:
                print("未检测到任何文件，请重新输入。")
                continue
            print("直接回车结束输入。")
            break
        valid = _collect_valid_paths(raw_list)
        if not valid:
            print("未读取到有效文件，请重新粘贴/输入。")
            continue
        paths.extend(valid)
        _print_identified_paths(paths)
        if len(paths) >= expected_count:
            break
        print(f"当前已收到 {len(paths)} 个有效文件，还需要 {expected_count - len(paths)} 个，请继续粘贴。")
    if len(paths) > expected_count:
        print(f"收到 {len(paths)} 个文件，按顺序取前 {expected_count} 个。")
    return paths[:expected_count]


def ask_single_path(prompt: str, allow_empty: bool = False) -> Path | None:
    print(prompt)
    while True:
        raw_list = read_paths_batch("> ")
        valid = _collect_valid_paths(raw_list)
        if not valid:
            if allow_empty and not raw_list:
                return None
            print("未检测到有效文件路径，请重新粘贴/输入。")
            continue
        if len(valid) > 1:
            print(f"检测到 {len(valid)} 个文件，本步骤只需 1 个，其余文件暂存给后续步骤。")
            _PUTBACK[:0] = [str(p) for p in valid[1:]]
        print(f"已识别文件：{valid[0].name}")
        return valid[0]


def ask_optional_paths(prompt: str, max_count: int) -> list[Path]:
    print(prompt)
    print("支持一次粘贴多个：Ctrl+V 多行粘贴、单行多个路径，或输入 c 从剪贴板读取全部路径。")
    print("粘贴完成后按一次回车结束本轮；全部文件输入完成后按回车结束。")
    paths: list[Path] = []
    while len(paths) < max_count:
        raw_list = read_paths_batch("> ")
        if not raw_list:
            break
        valid = _collect_valid_paths(raw_list)
        if not valid:
            print("未读取到有效文件，请重新粘贴/输入。")
            continue
        paths.extend(valid)
        _print_identified_paths(paths)
        if len(paths) < max_count:
            print("还可继续粘贴下一批文件，直接回车结束。")
    return paths


def is_excel_busy_error(exc: Exception) -> bool:
    if isinstance(exc, pywintypes.com_error):
        hresult = getattr(exc, "hresult", None)
        if hresult is None and exc.args:
            hresult = exc.args[0]
        return hresult in EXCEL_BUSY_HRESULTS
    if isinstance(exc, AttributeError):
        text = str(exc)
        return any(
            token in text
            for token in (
                ".Save",
                ".Close",
                ".Quit",
                ".RefreshAll",
                ".Open",
                ".CutCopyMode",
                ".CalculateUntilAsyncQueriesDone",
            )
        )
    return False


def excel_retry(action: str, func, attempts: int = 120, delay: float = 1.0):
    last_exc: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            return func()
        except Exception as exc:
            if not is_excel_busy_error(exc) or attempt >= attempts:
                raise
            last_exc = exc
            print(f"{action} 时 Excel 忙，{delay:.0f} 秒后自动重试（{attempt}/{attempts}）...", flush=True)
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(delay)
    if last_exc is not None:
        raise last_exc
    raise RuntimeError(f"{action} 失败。")


def get_excel_process_id(excel) -> int | None:
    try:
        hwnd = excel.Hwnd
        if not hwnd:
            return None
        _thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        return int(process_id) if process_id else None
    except Exception:
        return None


def is_process_alive(process_id: int | None) -> bool:
    if not process_id:
        return False
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {process_id}"],
            check=False,
            capture_output=True,
        )
        return str(process_id).encode("ascii") in result.stdout
    except Exception:
        return False


def wait_for_process_exit(process_id: int | None, timeout_seconds: float = 10.0) -> bool:
    if not process_id:
        return True
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if not is_process_alive(process_id):
            return True
        time.sleep(0.5)
    return not is_process_alive(process_id)


def force_kill_excel_process(process_id: int | None) -> None:
    if not process_id or not is_process_alive(process_id):
        return
    print(f"检测到脚本启动的 Excel 进程仍未退出，正在强制关闭 PID={process_id} ...", flush=True)
    subprocess.run(
        ["taskkill", "/PID", str(process_id), "/T", "/F"],
        check=False,
        capture_output=True,
    )


def as_column_matrix(values: Sequence[object]) -> tuple[tuple[object], ...]:
    return tuple((excel_safe_value(value),) for value in values)


def as_row_matrix(values: Sequence[Sequence[object]]) -> tuple[tuple[object, ...], ...]:
    # excel_safe_value：以“=”开头的文本会被 Excel 当作公式解析，
    # 统一在此处加单引号转为纯文本（详见 excel_safe_value 注释）。
    return tuple(tuple(excel_safe_value(value) for value in row) for row in values)


# 超大表（数十万行）通过 COM 一次性读写容易触发 Excel 内存溢出（0x8007000E），
# 因此新增分块读写。经实测（源表 39.8万行×40列）：块大小对每块耗时是线性的，
# 但超大块（5万行）因 Excel 对超大 Range 的处理开销反而更慢——
# 2万块全量约 53.6s，5万块约 64s。故取 2 万行一块，速度更快也更稳。
CHUNK_ROWS = 20000


def excel_safe_value(value: object) -> object:
    """把写入 Excel 的单元格值做安全化处理。

    数据源里存在以“=”开头的文本（如 "=844元 @Luo"，见 源表 第8792行），
    这类文本通过 COM 写入时会被 Excel 当作公式解析，导致：
    - 非法公式直接抛错（0x8007000E / 0x800A03EC）
    - 部分“公式”被解析成错误值，静默写坏数据（如 =FBA19HX22PYG）
    修复方式：在字符串前加一个单引号，Excel 会把整串当作文本原样写入。
    """
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def iter_range_rows_in_chunks(
    sheet,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    chunk_rows: int | None = None,
):
    """按块（默认 CHUNK_ROWS，可传 chunk_rows 覆盖）逐块读取单元格区域并 yield 为 list[list]。"""
    chunk = chunk_rows or CHUNK_ROWS
    for block_start in range(start_row, end_row + 1, chunk):
        block_end = min(block_start + chunk - 1, end_row)
        raw = sheet.Range(sheet.Cells(block_start, start_col), sheet.Cells(block_end, end_col)).Value
        rows = range_to_rows(raw)
        yield block_start, rows


def write_rows_in_chunks(sheet, start_row: int, start_col: int, rows: Sequence[Sequence[object]]) -> None:
    """按 CHUNK_ROWS（默认 5 万行）一块，把 rows 写入目标区域（每块内部列宽对齐）。"""
    if not rows:
        return
    row_cursor = start_row
    for block_start in range(0, len(rows), CHUNK_ROWS):
        block = rows[block_start:block_start + CHUNK_ROWS]
        col_count = max(len(row) for row in block)
        matrix = as_row_matrix([row + [None] * (col_count - len(row)) for row in block])
        sheet.Range(sheet.Cells(row_cursor, start_col), sheet.Cells(row_cursor + len(block) - 1, start_col + col_count - 1)).Value = matrix
        row_cursor += len(block)


def clear_range_in_chunks(sheet, start_row: int, start_col: int, end_row: int, end_col: int) -> None:
    """按 CHUNK_ROWS（默认 5 万行）一块，清空超大区域（ClearContents）。"""
    for block_start in range(start_row, end_row + 1, CHUNK_ROWS):
        block_end = min(block_start + CHUNK_ROWS - 1, end_row)
        sheet.Range(sheet.Cells(block_start, start_col), sheet.Cells(block_end, end_col)).ClearContents()


def range_to_rows(range_value: object) -> list[list[object]]:
    if range_value is None:
        return []
    if not isinstance(range_value, tuple):
        return [[range_value]]
    rows: list[list[object]] = []
    for row in range_value:
        if isinstance(row, tuple):
            rows.append(list(row))
        else:
            rows.append([row])
    return rows


def range_to_list(range_value: object) -> list[object]:
    rows = range_to_rows(range_value)
    result: list[object] = []
    for row in rows:
        result.extend(row)
    return result


def unique_preserve_order(values: Iterable[object], key_fn) -> list[object]:
    seen: set[str] = set()
    result: list[object] = []
    for value in values:
        key = key_fn(value)
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def last_used_row(sheet, col: int | None = None) -> int:
    if col is not None:
        row = sheet.Cells(sheet.Rows.Count, col).End(XL_UP).Row
        if row < 1:
            return 1
        return row
    used = sheet.UsedRange
    return used.Row + used.Rows.Count - 1


def last_used_col(sheet) -> int:
    used = sheet.UsedRange
    return used.Column + used.Columns.Count - 1


def clear_rows_from(sheet, start_row: int) -> None:
    end_row = last_used_row(sheet)
    if end_row >= start_row:
        sheet.Rows(f"{start_row}:{end_row}").Delete()


def clear_column_contents(sheet, col_letter: str) -> None:
    sheet.Range(f"{col_letter}:{col_letter}").ClearContents()


def get_sheet_by_name_or_header(workbook, preferred_name: str, header_value: str):
    for sheet in workbook.Worksheets:
        if sheet.Name == preferred_name:
            return sheet
    for sheet in workbook.Worksheets:
        if normalize_text(sheet.Cells(1, 1).Value) == header_value:
            return sheet
    raise ValueError(f"未找到工作表：{preferred_name} / {header_value}")


def extract_prefix_and_name(value: object) -> tuple[str | None, str]:
    text = preserve_text(value)
    if not text:
        return None, ""
    if "-" not in text:
        return None, text
    before, after = text.split("-", 1)
    prefix = before
    clean_name = after.replace("-", "")
    if not prefix:
        return None, clean_name
    return f"{prefix}-", clean_name


def print_section(title: str, values: Sequence[object]) -> None:
    print()
    print(title)
    if not values:
        print("无")
        return
    for value in values:
        print(normalize_text(value))


def parse_csv_input(raw: str) -> list[str]:
    text = raw.replace("，", ",").strip()
    return [part.strip() for part in text.split(",") if part.strip()]


def ask_cde(a_value: object, b_value: object) -> tuple[str, str, str]:
    print()
    print("以下业务员需要人工补充 C/D/E：")
    print(f"A={preserve_text(a_value)}")
    print(f"B={preserve_text(b_value)}")
    print("请输入 C,D,E，使用英文逗号或中文逗号分隔。")
    while True:
        parts = parse_csv_input(input("> "))
        if len(parts) == 3:
            return parts[0], parts[1], parts[2]
        print("输入格式不正确，请重新输入，例如：销售六部,事业一部,事业一部")


def confirm_or_edit_row(row_data: dict[str, object]) -> dict[str, object]:
    while True:
        print()
        print("请确认以下 A/B/C/D/E：")
        print(
            f"A={preserve_text(row_data['A'])} | "
            f"B={preserve_text(row_data['B'])} | "
            f"C={normalize_text(row_data['C'])} | "
            f"D={normalize_text(row_data['D'])} | "
            f"E={normalize_text(row_data['E'])}"
        )
        print("输入 1 确认，输入 2 修改。")
        choice = input("> ").strip()
        if choice == "1":
            return row_data
        if choice != "2":
            print("只接受 1 或 2，请重新输入。")
            continue
        print("请输入修改后的数据。支持两种格式：")
        print("1. 只改 C/D/E：C,D,E")
        print("2. 整行覆盖：A,B,C,D,E")
        parts = parse_csv_input(input("> "))
        if len(parts) == 3:
            row_data["C"], row_data["D"], row_data["E"] = parts
            continue
        if len(parts) == 5:
            row_data["A"], row_data["B"], row_data["C"], row_data["D"], row_data["E"] = parts
            continue
        print("输入格式不正确，请重新输入。")


def find_summary_row(source_sheet) -> int | None:
    end_row = last_used_row(source_sheet)
    start_row = max(1, end_row - 10)
    for row in range(end_row, start_row - 1, -1):
        left_values = range_to_list(source_sheet.Range(source_sheet.Cells(row, 1), source_sheet.Cells(row, 17)).Value)
        right_values = range_to_list(source_sheet.Range(source_sheet.Cells(row, 18), source_sheet.Cells(row, 33)).Value)
        has_left = any(normalize_text(value) for value in left_values)
        has_right = any(normalize_text(value) for value in right_values)
        if has_right and not has_left:
            return row
    last_right = range_to_list(source_sheet.Range(source_sheet.Cells(end_row, 18), source_sheet.Cells(end_row, 33)).Value)
    if any(normalize_text(value) for value in last_right):
        return end_row
    return None


def convert_column_to_numbers(sheet, col: int, start_row: int, end_row: int) -> None:
    if end_row < start_row:
        return
    raw_values = range_to_list(sheet.Range(sheet.Cells(start_row, col), sheet.Cells(end_row, col)).Value)
    converted = [to_excel_number(value) for value in raw_values]
    sheet.Range(sheet.Cells(start_row, col), sheet.Cells(end_row, col)).Value = as_column_matrix(converted)


def detect_header_row(first_row_values: Sequence[object], expected_headers: Sequence[str]) -> bool:
    matched = 0
    for index, expected in enumerate(expected_headers):
        if index >= len(first_row_values):
            break
        if normalize_text(first_row_values[index]) == expected:
            matched += 1
    return matched >= 2


def find_formula_seed_row(sheet, col: int, start_row: int) -> int | None:
    for row in range(start_row - 1, 1, -1):
        formula = sheet.Cells(row, col).Formula
        if normalize_text(formula):
            return row
    return None


def fill_down_formula(sheet, col: int, start_row: int, end_row: int) -> None:
    if end_row < start_row:
        return
    seed_row = find_formula_seed_row(sheet, col, start_row)
    if seed_row is None:
        raise ValueError(f"未找到第 {col} 列可下拉的公式模板。")
    sheet.Range(sheet.Cells(seed_row, col), sheet.Cells(end_row, col)).FillDown()


def build_summary_name_departments_map(
    sales_people: Sequence[object],
    sales_departments: Sequence[object],
) -> dict[str, list[str]]:
    name_departments: dict[str, list[str]] = {}
    for person, department in zip(sales_people, sales_departments):
        name = preserve_text(person)
        dept = normalize_department_name(department)
        if not name or not dept:
            continue
        departments = name_departments.setdefault(name, [])
        if dept not in departments:
            departments.append(dept)
    return name_departments


def derive_existing_sales_maps(
    sheet,
) -> tuple[
    dict[str, tuple[str, str, str]],
    dict[str, tuple[str, str, str]],
    dict[str, tuple[str, str]],
]:
    prefix_map = dict(PREFIX_MAP)
    name_map: dict[str, tuple[str, str, str]] = {}
    department_map: dict[str, tuple[str, str]] = {}
    end_row = last_used_row(sheet, 1)
    if end_row < 2:
        return prefix_map, name_map, department_map
    values = range_to_rows(sheet.Range(sheet.Cells(2, 1), sheet.Cells(end_row, 5)).Value)
    for row in values:
        row += [None] * (5 - len(row))
        a_value, b_value, c_value, d_value, e_value = row[:5]
        department = normalize_department_name(c_value)
        org = normalize_text(d_value)
        performance = normalize_text(e_value)
        cde = (department, org, performance)
        if any(cde):
            prefix, name = extract_prefix_and_name(a_value)
            if prefix and prefix not in prefix_map:
                prefix_map[prefix] = cde
            if name and name not in name_map:
                name_map[name] = cde
            pure_name = preserve_text(b_value)
            if pure_name and pure_name not in name_map:
                name_map[pure_name] = cde
        if department and department not in department_map and (org or performance):
            department_map[department] = (org, performance)
    return prefix_map, name_map, department_map


def resolve_pure_name_department(
    pure_name: str,
    summary_name_departments: dict[str, list[str]],
    department_map: dict[str, tuple[str, str]],
) -> tuple[str, str, str] | None:
    departments = summary_name_departments.get(pure_name, [])
    if not departments:
        return None
    if len(departments) > 1:
        print()
        print(f"纯姓名 {pure_name} 在综合数据明细中匹配到多个部门：{'、'.join(departments)}")
        print("该条将转为人工确认。")
        return None

    department = departments[0]
    org, performance = department_map.get(department, ("", ""))
    return department, org, performance


def override_pure_name_department(department: object) -> tuple[str, str] | None:
    original_department = normalize_text(department)
    if not original_department:
        return None

    for city_name, mapped_org in PURE_NAME_DEPARTMENT_OVERRIDE_MAP.items():
        if original_department.startswith(city_name) and original_department.endswith("部"):
            return original_department, mapped_org
    return None


def open_workbook(excel, path: Path):
    return excel_retry(
        f"打开工作簿 {path.name}",
        lambda: excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False),
        attempts=30,
        delay=1.0,
    )


def process_org_delivery_workbook(excel) -> None:
    target_path = ask_single_path("请粘贴“组织机构新增发货表.xlsx”路径（支持批量粘贴或输入 c 从剪贴板读取），然后回车。")
    source_path = ask_single_path("请粘贴“组织机构新增发货表”的新数据源表路径（支持批量粘贴或输入 c 从剪贴板读取），然后回车。")

    target_wb = open_workbook(excel, target_path)
    source_wb = open_workbook(excel, source_path)
    try:
        target_ws = target_wb.Worksheets(1)
        source_ws = source_wb.Worksheets(1)

        date_seed_value = target_ws.Cells(2, 1).Value
        date_seed_formula = target_ws.Cells(2, 1).Formula
        date_seed_format = target_ws.Cells(2, 1).NumberFormat

        # 数据源有表头，从第2行开始读数据，源文件不做任何修改
        source_last_row = last_used_row(source_ws)
        source_last_col = last_used_col(source_ws)
        if source_last_row < 2:
            raise ValueError("组织机构新增发货表的新数据源表没有可用数据。")

        target_last_row = last_used_row(target_ws)
        target_clear_end_col = max(last_used_col(target_ws), 14)
        if target_last_row >= 2:
            target_ws.Range(
                target_ws.Cells(2, 2),
                target_ws.Cells(target_last_row, target_clear_end_col),
            ).ClearContents()
            target_ws.Range(
                target_ws.Cells(2, 1),
                target_ws.Cells(target_last_row, 1),
            ).ClearContents()

        source_values = range_to_rows(
            source_ws.Range(
                source_ws.Cells(2, 1),
                source_ws.Cells(source_last_row, source_last_col),
            ).Value
        )
        target_end_row = 1 + len(source_values)
        target_ws.Range(
            target_ws.Cells(2, 2),
            target_ws.Cells(target_end_row, 1 + source_last_col),
        ).Value = as_row_matrix(source_values)

        fill_count = len(source_values)
        if fill_count > 0:
            target_ws.Cells(2, 1).Formula = date_seed_formula
            target_ws.Cells(2, 1).NumberFormat = date_seed_format
            if not preserve_text(target_ws.Cells(2, 1).Value):
                target_ws.Cells(2, 1).Value = date_seed_value
            if target_end_row > 2:
                target_ws.Range(
                    target_ws.Cells(2, 1),
                    target_ws.Cells(target_end_row, 1),
                ).FillDown()

        target_wb.Save()
    finally:
        source_wb.Close(SaveChanges=False)
        target_wb.Close(SaveChanges=False)


def process_source_and_summary(
    excel,
    source_path: Path,
    summary_path: Path,
) -> tuple[list[object], list[object], list[object]]:
    source_wb = open_workbook(excel, source_path)
    summary_wb = open_workbook(excel, summary_path)
    try:
        source_ws = source_wb.Worksheets(1)
        summary_ws = summary_wb.Worksheets(1)

        source_ws.Rows(1).Delete()
        summary_row = find_summary_row(source_ws)
        if summary_row is not None:
            source_ws.Rows(summary_row).Delete()

        source_last_row = last_used_row(source_ws)
        source_last_col = last_used_col(source_ws)
        if source_last_row < 1:
            raise ValueError("数据源表在删除首行和汇总行后已无可用数据。")

        clear_rows_from(summary_ws, 2)

        # 原实现是整体 Copy/PasteSpecial，超大表容易因剪贴板承载过多
        # 数据导致 Excel 内存溢出，改用分块读取后直接写值。
        summary_row_cursor = 2
        print(f"正在写入综合数据明细（共约 {source_last_row} 行）...", flush=True)
        for _, block_rows in iter_range_rows_in_chunks(source_ws, 1, 1, source_last_row, source_last_col):
            write_rows_in_chunks(summary_ws, summary_row_cursor, 1, block_rows)
            summary_row_cursor += len(block_rows)
        summary_last_row = summary_row_cursor - 1
        convert_column_to_numbers(summary_ws, 10, 2, summary_last_row)
        customer_codes = range_to_list(
            summary_ws.Range(summary_ws.Cells(2, 10), summary_ws.Cells(summary_last_row, 10)).Value
        )
        sales_people = range_to_list(
            summary_ws.Range(summary_ws.Cells(2, 13), summary_ws.Cells(summary_last_row, 13)).Value
        )
        sales_departments = range_to_list(
            summary_ws.Range(summary_ws.Cells(2, 14), summary_ws.Cells(summary_last_row, 14)).Value
        )

        summary_wb.Save()
    finally:
        source_wb.Close(SaveChanges=False)
        summary_wb.Close(SaveChanges=False)

    return customer_codes, sales_people, sales_departments


def process_customer_info(
    excel,
    customer_info_path: Path,
    customer_codes: list[object],
) -> None:
    customer_wb = open_workbook(excel, customer_info_path)
    new_customer_wb = None
    try:
        customer_ws = get_sheet_by_name_or_header(customer_wb, "Sheet1", "客户代码")

        clear_column_contents(customer_ws, "T")
        customer_ws.Cells(1, 20).Value = "待检测客户代码"
        if customer_codes:
            customer_ws.Range(customer_ws.Cells(2, 20), customer_ws.Cells(1 + len(customer_codes), 20)).Value = as_column_matrix(
                customer_codes
            )

        unique_codes = unique_preserve_order(customer_codes, normalize_code_key)
        unique_codes = [to_excel_number(code) for code in unique_codes]
        clear_column_contents(customer_ws, "T")
        customer_ws.Cells(1, 20).Value = "待检测客户代码"
        if unique_codes:
            customer_ws.Range(customer_ws.Cells(2, 20), customer_ws.Cells(1 + len(unique_codes), 20)).Value = as_column_matrix(
                unique_codes
            )

        existing_last_row = last_used_row(customer_ws, 1)
        existing_codes = range_to_list(customer_ws.Range(customer_ws.Cells(2, 1), customer_ws.Cells(existing_last_row, 1)).Value)
        existing_code_keys = {normalize_code_key(value) for value in existing_codes if normalize_code_key(value)}
        new_code_list = [code for code in unique_codes if normalize_code_key(code) not in existing_code_keys]
        print_section("T列检测出的未重复客户代码如下：", new_code_list)

        if not new_code_list:
            new_customer_path = ask_single_path(
                "未检测到新增客户代码。如无需导入新客户表，请直接回车跳过；如仍需导入，请粘贴新客户表路径后回车。",
                allow_empty=True,
            )
            if new_customer_path is None:
                customer_ws.Columns("T").Delete()
                customer_wb.Save()
                return
        else:
            new_customer_path = ask_single_path("请粘贴新客户表路径（支持批量粘贴或输入 c 从剪贴板读取），然后回车。")

        new_customer_wb = open_workbook(excel, new_customer_path)
        new_customer_ws = new_customer_wb.Worksheets(1)

        new_customer_ws.Rows(1).Delete()
        source_last_row = last_used_row(new_customer_ws)
        source_last_col = last_used_col(new_customer_ws)
        source_start_row = 1
        if source_last_row < source_start_row:
            raise ValueError("新客户表没有可追加的数据。")

        copy_col_count = min(source_last_col, 16)
        source_values = range_to_rows(
            new_customer_ws.Range(
                new_customer_ws.Cells(source_start_row, 1),
                new_customer_ws.Cells(source_last_row, copy_col_count),
            ).Value
        )

        append_start_row = last_used_row(customer_ws, 1) + 1
        append_end_row = append_start_row + len(source_values) - 1
        customer_ws.Range(
            customer_ws.Cells(append_start_row, 1),
            customer_ws.Cells(append_end_row, copy_col_count),
        ).Value = as_row_matrix(source_values)

        convert_column_to_numbers(customer_ws, 2, append_start_row, append_end_row)
        fill_down_formula(customer_ws, 17, append_start_row, append_end_row)

        customer_ws.Columns("T").Delete()
        customer_wb.Save()
    finally:
        if new_customer_wb is not None:
            new_customer_wb.Close(SaveChanges=False)
        customer_wb.Close(SaveChanges=False)


def process_sales_account(
    excel,
    sales_account_path: Path,
    sales_people: list[object],
    sales_departments: list[object],
) -> None:
    sales_wb = open_workbook(excel, sales_account_path)
    try:
        sales_ws = get_sheet_by_name_or_header(sales_wb, "账号明细表", "业务员")

        clear_column_contents(sales_ws, "H")
        sales_ws.Cells(1, 8).Value = "待检测业务员"
        if sales_people:
            sales_ws.Range(sales_ws.Cells(2, 8), sales_ws.Cells(1 + len(sales_people), 8)).Value = as_column_matrix(
                sales_people
            )

        unique_sales_people = unique_preserve_order(sales_people, preserve_text)
        clear_column_contents(sales_ws, "H")
        sales_ws.Cells(1, 8).Value = "待检测业务员"
        if unique_sales_people:
            sales_ws.Range(
                sales_ws.Cells(2, 8), sales_ws.Cells(1 + len(unique_sales_people), 8)
            ).Value = as_column_matrix(unique_sales_people)

        existing_last_row = last_used_row(sales_ws, 1)
        existing_sales = range_to_list(sales_ws.Range(sales_ws.Cells(2, 1), sales_ws.Cells(existing_last_row, 1)).Value)
        existing_sales_keys = {preserve_text(value) for value in existing_sales if preserve_text(value)}
        new_sales_people = [name for name in unique_sales_people if preserve_text(name) not in existing_sales_keys]
        print()
        print("H列检测出的未重复业务员如下：")
        if not new_sales_people:
            print("无")
        else:
            for value in new_sales_people:
                print(preserve_text(value))

        if new_sales_people:
            append_start_row = existing_last_row + 1
            append_end_row = append_start_row + len(new_sales_people) - 1
            sales_ws.Range(sales_ws.Cells(append_start_row, 1), sales_ws.Cells(append_end_row, 1)).Value = as_column_matrix(
                new_sales_people
            )

            prefix_map, name_map, department_map = derive_existing_sales_maps(sales_ws)
            summary_name_departments = build_summary_name_departments_map(sales_people, sales_departments)

            for row in range(append_start_row, append_end_row + 1):
                a_value = sales_ws.Cells(row, 1).Value
                prefix, name = extract_prefix_and_name(a_value)
                pure_name_department_override: tuple[str, str] | None = None
                row_data = {
                    "A": preserve_text(a_value),
                    "B": name or preserve_text(a_value),
                    "C": "",
                    "D": "",
                    "E": "",
                }

                needs_confirmation = False
                if prefix and prefix in prefix_map:
                    row_data["C"], row_data["D"], row_data["E"] = prefix_map[prefix]
                elif not prefix and (
                    resolved := resolve_pure_name_department(
                        row_data["B"],
                        summary_name_departments,
                        department_map,
                    )
                ):
                    row_data["C"], row_data["D"], row_data["E"] = resolved
                    needs_confirmation = True
                elif not prefix and row_data["B"] in name_map:
                    row_data["C"], row_data["D"], row_data["E"] = name_map[row_data["B"]]
                    needs_confirmation = True
                else:
                    row_data["C"], row_data["D"], row_data["E"] = ask_cde(row_data["A"], row_data["B"])
                    needs_confirmation = True

                if not prefix:
                    pure_name_department_override = override_pure_name_department(row_data["C"])
                    if pure_name_department_override is not None:
                        original_department, mapped_department = pure_name_department_override
                        row_data["C"] = mapped_department
                        mapped_de = department_map.get(mapped_department)
                        if mapped_de is not None:
                            row_data["D"], row_data["E"] = mapped_de
                        else:
                            row_data["D"] = mapped_department
                            row_data["E"] = mapped_department

                    print()
                    print("纯姓名回填结果如下：")
                    if pure_name_department_override is not None:
                        original_department, mapped_department = pure_name_department_override
                        print(
                            f"A={preserve_text(row_data['A'])}，"
                            f"B={preserve_text(row_data['B'])}，"
                            f"映射前：{original_department}，"
                            f"映射后：{mapped_department}。"
                        )
                    print(
                        f"A={preserve_text(row_data['A'])} | "
                        f"B={preserve_text(row_data['B'])} | "
                        f"C={normalize_text(row_data['C'])} | "
                        f"D={normalize_text(row_data['D'])} | "
                        f"E={normalize_text(row_data['E'])}"
                    )
                    needs_confirmation = True

                if needs_confirmation:
                    row_data = confirm_or_edit_row(row_data)

                sales_ws.Cells(row, 1).Value = row_data["A"]
                sales_ws.Cells(row, 2).Value = row_data["B"]
                sales_ws.Cells(row, 3).Value = row_data["C"]
                sales_ws.Cells(row, 4).Value = row_data["D"]
                sales_ws.Cells(row, 5).Value = row_data["E"]

        sales_ws.Columns("H").Delete()
        sales_wb.Save()
    finally:
        sales_wb.Close(SaveChanges=False)


def touch_base_date_workbook(excel, base_date_path: Path) -> None:
    base_wb = open_workbook(excel, base_date_path)
    try:
        base_wb.Save()
    finally:
        base_wb.Close(SaveChanges=False)


def touch_workbook(excel, workbook_path: Path) -> None:
    workbook = open_workbook(excel, workbook_path)
    try:
        workbook.Save()
    finally:
        workbook.Close(SaveChanges=False)


def create_excel_app():
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    try:
        excel.Calculation = XL_CALCULATION_AUTOMATIC
    except Exception:
        pass
    return excel, get_excel_process_id(excel)


def quit_excel_app(excel, excel_process_id: int | None) -> None:
    try:
        try:
            workbook_count = excel.Workbooks.Count
        except Exception:
            workbook_count = 0

        while workbook_count > 0:
            excel_retry(
                "关闭残留工作簿",
                lambda: excel.Workbooks.Item(excel.Workbooks.Count).Close(SaveChanges=False),
                attempts=60,
                delay=1.0,
            )
            try:
                workbook_count = excel.Workbooks.Count
            except Exception:
                workbook_count = 0

        excel_retry("清理 Excel 剪贴板状态", lambda: setattr(excel, "CutCopyMode", False), attempts=30, delay=1.0)
        excel_retry("退出 Excel", lambda: excel.Quit(), attempts=60, delay=1.0)
        if not wait_for_process_exit(excel_process_id, timeout_seconds=8.0):
            force_kill_excel_process(excel_process_id)
    finally:
        pythoncom.CoUninitialize()


def ask_mode() -> int:
    print("请选择模式：")
    print("  模式 1：完整业务处理")
    print("         综合数据明细→客户信息→销售账号明细→组织机构发货→保存日期表")
    print("  模式 2：数据覆盖")
    print("         粘贴数据源 + 目标数据表，将数据源全部内容覆盖到目标数据表")
    print("  模式 3：合并表格")
    print("         粘贴多个表格，可逐个文件选择要合并的工作表，并支持按列关键词筛选数据行，再合并成一个新文件")
    print("  模式 4：工作表处理")
    print("         粘贴一个表格，列出全部工作表供选择，再二选一：")
    print("         ① 复制全部数据并粘贴（清除公式）")
    print("         ② 列出列名，选择要保留的列，删除其余列")
    print()
    while True:
        choice = input("请输入模式编号（1、2、3 或 4）：> ").strip()
        if choice in ("1", "2", "3", "4"):
            return int(choice)
        print("输入无效，请重新输入 1、2、3 或 4。")


def ask_sheet_selection(sheet_names: list[str]) -> list[int]:
    """列出工作表并让用户选择要处理的编号，返回 0-based 索引列表。

    直接回车表示选择全部工作表。
    """
    print()
    print("该文件包含以下工作表：")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    while True:
        line = input(
            "请选择工作表编号（可多选，英文逗号或中文逗号分隔，如 1,3；直接回车表示全部）：> "
        ).strip()
        if not line:
            return list(range(len(sheet_names)))
        parts = parse_csv_input(line)
        indices: list[int] = []
        valid = True
        for part in parts:
            if not part.isdigit():
                print(f"无效编号：{part}")
                valid = False
                break
            idx = int(part)
            if 1 <= idx <= len(sheet_names):
                indices.append(idx - 1)
            else:
                print(f"编号 {idx} 超出范围（1-{len(sheet_names)}）。")
                valid = False
                break
        if valid:
            return indices


def ask_row_filter() -> tuple[int, list[str]] | None:
    """询问是否按列关键词筛选数据行。返回 (列编号, 关键词列表) 或 None。"""
    print()
    answer = input(
        "是否需要对合并的数据做行筛选（只保留某列包含指定关键词的行，第 1 行表头始终保留）？（y/n，直接回车默认 n）：> "
    ).strip().lower()
    if answer not in ("y", "yes", "是"):
        return None
    while True:
        col_raw = input("请输入用于筛选的列编号（第 1 列 = 1）：> ").strip()
        if col_raw.isdigit() and int(col_raw) >= 1:
            col = int(col_raw)
            break
        print("请输入正整数列编号。")
    while True:
        keywords = parse_csv_input(
            input("请输入筛选关键词（多个用英文逗号或中文逗号分隔，匹配任一即保留）：> ")
        )
        if keywords:
            return col, keywords
        print("至少输入一个关键词。")


# 模式 3 合并结果写入时使用的小块大小，避免单次赋值超大矩阵触发 Excel 内存不足（0x8007000E）
MODE3_WRITE_CHUNK = 500


def write_merged_rows_in_small_chunks(sheet, start_row: int, rows: Sequence[Sequence[object]]) -> int:
    """按小块把合并数据写入目标表，返回写入后的下一行光标。"""
    if not rows:
        return start_row
    for block_start in range(0, len(rows), MODE3_WRITE_CHUNK):
        block = rows[block_start:block_start + MODE3_WRITE_CHUNK]
        col_count = max(len(r) for r in block)
        matrix = as_row_matrix([r + [None] * (col_count - len(r)) for r in block])
        sheet.Range(
            sheet.Cells(start_row, 1),
            sheet.Cells(start_row + len(block) - 1, col_count),
        ).Value = matrix
        start_row += len(block)
    return start_row


def clear_formulas_by_paste_values(excel, sheet) -> int:
    """复制整个已用区域并原位粘贴为数值，清除公式但保留格式/合并单元格。返回单元格总数。"""
    last_row = last_used_row(sheet)
    last_col = last_used_col(sheet)
    if last_row < 1 or last_col < 1:
        return 0
    target_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_col))
    excel_retry("复制全部数据", lambda: target_range.Copy(), attempts=60, delay=1.0)
    try:
        excel_retry(
            "原位粘贴为数值",
            lambda: sheet.Range(
                sheet.Cells(1, 1), sheet.Cells(last_row, last_col)
            ).PasteSpecial(XL_PASTE_VALUES),
            attempts=60,
            delay=1.0,
        )
    finally:
        excel_retry(
            "清理 Excel 剪贴板状态",
            lambda: setattr(excel, "CutCopyMode", False),
            attempts=30,
            delay=1.0,
        )
    return last_row * last_col


def mode3(excel, excel_process_id: int | None) -> None:
    print()
    print("运行前请先关闭要处理的 Excel 文件，避免占用。")
    print()

    source_paths = ask_optional_paths(
        "请粘贴需要合并的表格文件（支持 Ctrl+V 多行粘贴、单行多个路径，或输入 c 从剪贴板读取全部路径；粘贴完按回车确认，全部输完再按回车结束）：",
        max_count=50,
    )
    if len(source_paths) < 2:
        raise ValueError("合并表格至少需要 2 个文件。")

    row_filter = ask_row_filter()

    first_file = source_paths[0]
    target_path = first_file.parent / f"{first_file.stem}_合并结果.xlsx"

    total = len(source_paths)

    # 边读边写：分块读取源表数据、小块写入目标表，避免一次性大矩阵触发
    # Excel 内存不足（0x8007000E），也不把全部数据堆在 Python 内存里。
    new_wb = excel.Workbooks.Add()
    target_ws = new_wb.Worksheets(1)
    saved = False
    try:
        write_row = 1
        total_rows = 0
        merged_sheet_count = 0

        for i, source_path in enumerate(source_paths, 1):
            print(f"正在读取 ({i}/{total}): {source_path.name}", flush=True)
            source_wb = open_workbook(excel, source_path)
            try:
                sheet_names = [ws.Name for ws in source_wb.Worksheets]
                selected_indices = ask_sheet_selection(sheet_names)
                for sheet_idx in selected_indices:
                    source_ws = source_wb.Worksheets(sheet_idx + 1)
                    source_last_row = last_used_row(source_ws)
                    source_last_col = last_used_col(source_ws)
                    if source_last_row < 1 or source_last_col < 1:
                        continue
                    merged_sheet_count += 1
                    for block_index, (block_start, block_rows) in enumerate(
                        iter_range_rows_in_chunks(
                            source_ws, 1, 1, source_last_row, source_last_col, chunk_rows=2000
                        ),
                        1,
                    ):
                        if row_filter is not None:
                            filter_col, keywords = row_filter
                            if filter_col <= source_last_col:
                                keep_rows = []
                                for row in block_rows:
                                    # 每个工作表第 1 行（表头）始终保留
                                    if block_start == 1 and block_index == 1 and row is block_rows[0]:
                                        keep_rows.append(row)
                                        continue
                                    cell_text = (
                                        normalize_text(row[filter_col - 1])
                                        if filter_col - 1 < len(row)
                                        else ""
                                    )
                                    if any(kw in cell_text for kw in keywords):
                                        keep_rows.append(row)
                                block_rows = keep_rows
                        if not block_rows:
                            continue
                        write_row = write_merged_rows_in_small_chunks(target_ws, write_row, block_rows)
                        total_rows += len(block_rows)
                        excel_retry(
                            "空闲清理",
                            lambda: pythoncom.PumpWaitingMessages(),
                            attempts=1,
                            delay=0.0,
                        )
            finally:
                source_wb.Close(SaveChanges=False)

        if total_rows < 1:
            raise ValueError("没有读取到任何数据。")

        print(f"正在保存结果（共 {total_rows} 行）...", flush=True)
        new_wb.SaveAs(str(target_path))
        saved = True
    finally:
        new_wb.Close(SaveChanges=saved)

    print()
    print(
        f"合并完成！共合并 {merged_sheet_count} 个工作表（{total} 个文件），"
        f"{total_rows} 行，输出到：{target_path}"
    )


def collect_all_sheet_names(excel, paths: Sequence[Path]) -> list[str]:
    """合并所有文件的工作表名称（去重，保持出现顺序）。"""
    names: list[str] = []
    seen: set[str] = set()
    for path in paths:
        wb = open_workbook(excel, path)
        try:
            for ws in wb.Worksheets:
                n = ws.Name
                if n not in seen:
                    seen.add(n)
                    names.append(n)
        finally:
            wb.Close(SaveChanges=False)
    return names


def ask_sheet_name_uniform(sheet_names: list[str]) -> str:
    """列出合并后的工作表，用户选一个（应用到所有文件）；直接回车默认第 1 个。"""
    print()
    print("全部表格的工作表合并如下（该选择将应用到所有文件，按工作表名称匹配）：")
    for i, name in enumerate(sheet_names, 1):
        print(f"  {i}. {name}")
    while True:
        line = input(
            f"请选择工作表编号（1-{len(sheet_names)}，直接回车默认第 1 个）：> "
        ).strip()
        if not line:
            return sheet_names[0]
        if line.isdigit() and 1 <= int(line) <= len(sheet_names):
            return sheet_names[int(line) - 1]
        print(f"请输入有效编号（1-{len(sheet_names)}）。")


def ask_keep_columns_from_union(
    col_names_all: list[str],
    owners_count: dict[str, int],
    total_files: int,
) -> set[str]:
    """从合并后的列名清单中选择保留列，返回保留列名集合。

    清单会标注每列是「全部文件都有」还是「仅部分文件有」，帮助区分不同表结构。
    """
    print()
    print(f"该工作表在 {total_files} 个文件中的列名合并如下（标注每列出现情况）：")
    for i, name in enumerate(col_names_all, 1):
        count = owners_count.get(name, 0)
        mark = "全部文件都有" if count >= total_files else f"仅 {count}/{total_files} 个文件有"
        label = name if name else f"（第 {i} 列无表头）"
        print(f"  {i}. {label}  [{mark}]")
    while True:
        line = input(
            "请选择要保留的列编号（可多选，英文逗号或中文逗号分隔，如 1,3,5；直接回车全选）：> "
        ).strip()
        if not line:
            return {name for name in col_names_all if name}
        parts = parse_csv_input(line)
        keep_indices: list[int] = []
        valid = True
        for part in parts:
            if not part.isdigit():
                print(f"无效编号：{part}")
                valid = False
                break
            idx = int(part)
            if 1 <= idx <= len(col_names_all):
                keep_indices.append(idx)
            else:
                print(f"编号 {idx} 超出范围（1-{len(col_names_all)}）。")
                valid = False
                break
        if valid and keep_indices:
            return {col_names_all[i - 1] for i in keep_indices}
        if valid:
            print("至少选择一个列。")


def keep_selected_columns_by_names(
    sheet,
    keep_names: set[str],
    col_names_template: list[str],
) -> int:
    """按统一选择的列名批量保留本表列，删除其余列。返回删除列数。

    本表列数与模板一致时按位置精确匹配；不一致时按列名匹配（无表头列删除）。
    """
    last_col = last_used_col(sheet)
    if last_col < 1:
        return 0
    headers = range_to_list(
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, last_col)).Value
    )
    col_names = [normalize_text(h) for h in headers]
    deleted = 0
    if len(col_names) == len(col_names_template):
        keep_by_position = [name in keep_names for name in col_names_template]
        for col in range(last_col, 0, -1):
            if keep_by_position[col - 1]:
                continue
            sheet.Columns(col).Delete()
            deleted += 1
    else:
        for col in range(last_col, 0, -1):
            name = col_names[col - 1] if col - 1 < len(col_names) else ""
            if name and name in keep_names:
                continue
            sheet.Columns(col).Delete()
            deleted += 1
    return deleted


def mode4(excel, excel_process_id: int | None) -> None:
    print()
    print("运行前请先关闭要处理的 Excel 文件，避免占用。")
    print()

    target_paths = ask_optional_paths(
        "请粘贴要处理的表格文件（支持一次粘贴多个：Ctrl+V 多行粘贴、单行多个路径，或输入 c 从剪贴板读取全部路径）：",
        max_count=500,
    )
    if not target_paths:
        raise ValueError("未选择任何文件。")

    print()
    print(f"共识别 {len(target_paths)} 个文件，将对每个文件执行所选功能。")

    print()
    print("请选择要执行的功能：")
    print("  1. 复制全部数据并粘贴（清除公式）")
    print("  2. 保留指定列，删除其余列")
    while True:
        choice = input("请输入功能编号（1 或 2）：> ").strip()
        if choice in ("1", "2"):
            function_no = int(choice)
            break
        print("输入无效，请重新输入 1 或 2。")

    # 合并所有文件的工作表名称，统一选择一个工作表
    all_sheet_names = collect_all_sheet_names(excel, target_paths)
    if not all_sheet_names:
        raise ValueError("未能读取任何工作表名称。")
    selected_sheet_name = ask_sheet_name_uniform(all_sheet_names)
    print(f"已选择工作表：{selected_sheet_name}（应用到所有文件，按名称匹配）")

    # 功能 2：先合并所有文件该工作表的列名，统一选择保留列，再统一处理
    keep_names: set[str] = set()
    col_names_template: list[str] = []
    if function_no == 2:
        col_names_all: list[str] = []
        owners_count: dict[str, int] = {}
        total_with_sheet = 0
        for path in target_paths:
            workbook = open_workbook(excel, path)
            try:
                sheets_by_name = {ws.Name: ws for ws in workbook.Worksheets}
                sheet = sheets_by_name.get(selected_sheet_name)
                if sheet is None:
                    continue
                total_with_sheet += 1
                last_col = last_used_col(sheet)
                headers = (
                    range_to_list(
                        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, last_col)).Value
                    )
                    if last_col >= 1
                    else []
                )
                colnames = [normalize_text(h) for h in headers]
                if not col_names_template:
                    col_names_template = colnames
                seen_in_file: set[str] = set()
                for name in colnames:
                    if name and name not in seen_in_file:
                        owners_count[name] = owners_count.get(name, 0) + 1
                        seen_in_file.add(name)
                    if name not in col_names_all:
                        col_names_all.append(name)
            finally:
                workbook.Close(SaveChanges=False)
        if not col_names_all:
            raise ValueError(f"所选工作表“{selected_sheet_name}”在所有文件中都没有可用列。")
        keep_names = ask_keep_columns_from_union(col_names_all, owners_count, total_with_sheet)

    # 逐个处理文件，逐个列出结果
    processed = 0
    for i, path in enumerate(target_paths, 1):
        workbook = open_workbook(excel, path)
        try:
            sheets_by_name = {ws.Name: ws for ws in workbook.Worksheets}
            sheet = sheets_by_name.get(selected_sheet_name)
            if sheet is None:
                print(f"[{i}/{len(target_paths)}] {path.name}：没有工作表“{selected_sheet_name}”，跳过。")
                continue
            if function_no == 1:
                cell_count = clear_formulas_by_paste_values(excel, sheet)
                workbook.Save()
                print(
                    f"[{i}/{len(target_paths)}] {path.name}：清除公式完成"
                    f"（工作表 {sheet.Name} 共 {cell_count} 个单元格）"
                )
            else:
                deleted = keep_selected_columns_by_names(sheet, keep_names, col_names_template)
                workbook.Save()
                print(
                    f"[{i}/{len(target_paths)}] {path.name}：保留所选列完成"
                    f"（工作表 {sheet.Name} 删除 {deleted} 列）"
                )
            processed += 1
        finally:
            workbook.Close(SaveChanges=False)

    print()
    print(f"处理完成：共处理 {processed}/{len(target_paths)} 个文件，均保存在原文件。")


def mode1(excel, excel_process_id: int | None) -> None:
    print("运行前请先关闭要处理的 Excel 文件，避免占用。")
    base_paths = ask_paths(expected_count=5)
    source_path, summary_path, customer_info_path, sales_account_path, base_date_path = base_paths

    customer_codes, sales_people, sales_departments = process_source_and_summary(excel, source_path, summary_path)

    # 综合明细表中的"阿里内勤"替换为"项目部营销"
    print()
    print("正在替换“阿里内勤”→“项目部营销”...", flush=True)
    summary_wb = open_workbook(excel, summary_path)
    try:
        summary_ws = summary_wb.Worksheets(1)
        summary_ws.Cells.Replace("阿里内勤", "项目部营销")
        summary_wb.Save()
    finally:
        summary_wb.Close(SaveChanges=False)

    process_customer_info(excel, customer_info_path, customer_codes)
    process_sales_account(excel, sales_account_path, sales_people, sales_departments)
    process_org_delivery_workbook(excel)
    touch_base_date_workbook(excel, base_date_path)
    touch_workbook(excel, customer_info_path)


def mode2(excel, excel_process_id: int | None) -> None:
    print()
    print("运行前请先关闭要处理的 Excel 文件，避免占用。")
    print()

    source_path = ask_single_path("请粘贴数据源文件路径（支持批量粘贴或输入 c 从剪贴板读取），然后回车。")
    target_path = ask_single_path("请粘贴目标数据表文件路径（支持批量粘贴或输入 c 从剪贴板读取），然后回车。")

    source_wb = open_workbook(excel, source_path)
    target_wb = open_workbook(excel, target_path)
    try:
        source_ws = source_wb.Worksheets(1)

        # 列出目标文件中的所有工作表供用户选择
        sheet_names = [ws.Name for ws in target_wb.Worksheets]
        if len(sheet_names) == 1:
            selected_sheet_name = sheet_names[0]
            print(f"目标文件只有一个工作表：{selected_sheet_name}，自动选择。")
        else:
            print("目标文件中包含以下工作表：")
            for i, name in enumerate(sheet_names, 1):
                print(f"  {i}. {name}")
            while True:
                choice = input(f"请选择要覆盖的工作表编号（1-{len(sheet_names)}）：> ").strip()
                try:
                    idx = int(choice)
                    if 1 <= idx <= len(sheet_names):
                        selected_sheet_name = sheet_names[idx - 1]
                        break
                    print(f"编号超出范围，请输入 1 到 {len(sheet_names)} 之间的数字。")
                except ValueError:
                    print("输入无效，请输入数字编号。")
        print(f"已选择工作表：{selected_sheet_name}")
        target_ws = target_wb.Worksheets(selected_sheet_name)

        # 同模式1：删除数据源第1行标题
        source_ws.Rows(1).Delete()

        # 同模式1：检测并删除合计行
        source_last_row = last_used_row(source_ws)
        if source_last_row >= 1:
            last_row_value = normalize_text(source_ws.Cells(source_last_row, 1).Value)
            if "合计" in last_row_value:
                print(f"检测到数据源末行为合计行，已自动排除（第 {source_last_row} 行）。")
                source_ws.Rows(source_last_row).Delete()

        source_last_row = last_used_row(source_ws)
        source_last_col = last_used_col(source_ws)
        if source_last_row < 1:
            raise ValueError("数据源表没有可用数据。")

        # 只清除目标表第2行起、前 source_last_col 列的内容；
        # 第 source_last_col + 1 列之后的内容（如查找/公式列）保留，不被覆盖。
        target_last_row = last_used_row(target_ws)
        if target_last_row >= 2:
            print(f"数据源共 {source_last_col} 列，目标表第 {source_last_col + 1} 列之后的内容将保留，不被覆盖。")
            clear_range_in_chunks(target_ws, 2, 1, target_last_row, source_last_col)

        # 同模式1：复制数据源全部内容到目标表第2行起。
        # 数据量可能达到数十万行，必须分块读取/写入，避免一次性
        # 搬运上千万个单元格导致 Excel 内存溢出（0x8007000E）。
        total_blocks = (source_last_row + CHUNK_ROWS - 1) // CHUNK_ROWS
        print(f"正在读取数据源（共约 {source_last_row} 行，分 {total_blocks} 块，每块 {CHUNK_ROWS} 行）...", flush=True)
        for block_index, (block_start, block_rows) in enumerate(
            iter_range_rows_in_chunks(source_ws, 1, 1, source_last_row, source_last_col),
            1,
        ):
            print(f"  [{block_index}/{total_blocks}] 第 {block_start}~{block_start + len(block_rows) - 1} 行", flush=True)
            write_rows_in_chunks(
                target_ws,
                1 + block_start,
                1,
                block_rows,
            )
            excel.CutCopyMode = False
            excel_retry("空闲清理", lambda: pythoncom.PumpWaitingMessages(), attempts=1, delay=0.0)

        target_wb.Save()
        print()
        print(f"数据覆盖完成：共 {source_last_row} 行数据 × {source_last_col} 列")
    finally:
        source_wb.Close(SaveChanges=False)
        target_wb.Close(SaveChanges=False)


def main() -> int:
    for stream_name in ("stdin", "stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    excel, excel_process_id = create_excel_app()
    try:
        while True:
            mode = ask_mode()
            if mode == 1:
                mode1(excel, excel_process_id)
            elif mode == 2:
                mode2(excel, excel_process_id)
            elif mode == 3:
                mode3(excel, excel_process_id)
            elif mode == 4:
                mode4(excel, excel_process_id)

            print()
            print("全部处理完成。")
            again = input("是否继续处理其他模式？（输入 y 继续，其他任意键结束）：> ").strip().lower()
            if again not in ("y", "yes", "是"):
                break
    finally:
        quit_excel_app(excel, excel_process_id)

    print()
    print("已结束，谢谢使用。")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover
        print()
        print(f"处理失败：{exc}")
        print(traceback.format_exc())
        raise SystemExit(1)
