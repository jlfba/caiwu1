#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将预估报表“汇总”按表头追加到利润报表“明细表”。

双击运行可自行多选文件；--all 可按默认目录在命令行无界面运行。
需要本机 Microsoft Excel、openpyxl、pywin32。
"""

from __future__ import annotations

import argparse
import ctypes
import logging
import os
import shutil
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import openpyxl
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils.datetime import to_excel


ROOT = Path("D:/caiwu/报表部/微盘专线汇总(1)/微盘专线汇总")
REPORT_DIR = ROOT / "微盘专线报表"
SOURCE_DIR = ROOT / "专线数据 取数"
EUROPE_SHEETS = (
    "美国空派", "欧洲空派", "欧洲卡航", "欧洲海派", "英国铁路",
    "欧洲铁路", "英国空派", "英国卡航", "英国海派",
)
COPY_COLUMNS = (
    "创建时间", "主单号", "配仓机构", "配仓分类", "柜号", "出货计费重", "应收",
    "业务员名义总成本", "服务商总应付", "主单毛利率",
    "应付-基础运费", "应付-海运费", "应付-进仓费", "应付-报关费", "应付-关税",
    "应付-清关费", "应付-拆柜费", "应付-打单费", "应付-调整费用", "应付-提货费",
    "应付-卡派", "应付-运费赔偿", "应付-拖车费",
)
FORMULA_COLUMNS = ("月份", "周数", "应付-其他费用")
XL_UP = -4162
XL_PASTE_FORMATS = -4122
XL_NONE = -4142
LOG = logging.getLogger("专线利润报表追加")


@dataclass(frozen=True)
class ReportFile:
    path: Path
    line: str
    role: str


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return "".join(str(value).split()).replace("－", "-").replace("—", "-")


def identify(path: Path) -> ReportFile | None:
    name = path.stem
    if path.name.startswith("~$") or path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return None
    role = "预估报表" if "预估报表" in name else "利润报表" if "利润报表" in name else None
    lines = [key for key in ("加", "美", "欧") if key in name]
    if role is None or len(lines) != 1:
        return None
    return ReportFile(path.resolve(), lines[0], role)


def discover(source_dir: Path, report_dir: Path) -> list[ReportFile]:
    found = []
    seen: set[Path] = set()
    folders = (source_dir, report_dir) if source_dir.is_dir() and report_dir.is_dir() else (ROOT,)
    for folder in folders:
        if not folder.is_dir():
            raise FileNotFoundError(f"目录不存在：{folder}")
        for path in sorted(folder.iterdir(), key=lambda item: item.name):
            if path.is_file():
                item = identify(path)
                if item and item.path not in seen:
                    found.append(item)
                    seen.add(item.path)
    return found


def get_pairs(files: list[ReportFile]) -> tuple[list[tuple[ReportFile, list[ReportFile]]], list[str]]:
    grouped: dict[str, dict[str, list[ReportFile]]] = defaultdict(lambda: defaultdict(list))
    for item in files:
        grouped[item.line][item.role].append(item)
    pairs = []
    skipped = []
    for line in sorted(grouped):
        sources = grouped[line]["预估报表"]
        targets = grouped[line]["利润报表"]
        if len(targets) != 1:
            skipped.append(f"{line}线：需要恰好一份利润报表，实际 {len(targets)} 份")
        elif not sources:
            skipped.append(f"{line}线：没有选中预估报表")
        else:
            pairs.append((targets[0], sources))
    return pairs, skipped


def setup_log(folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"专线利润报表追加_{datetime.now():%Y%m%d_%H%M%S}.log"
    close_log()
    LOG.setLevel(logging.INFO)
    LOG.propagate = False
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    for handler in (logging.FileHandler(path, encoding="utf-8"), logging.StreamHandler(sys.stdout)):
        handler.setFormatter(formatter)
        LOG.addHandler(handler)
    return path


def close_log() -> None:
    for handler in LOG.handlers[:]:
        LOG.removeHandler(handler)
        handler.close()


def source_headers(ws: Any) -> dict[str, int]:
    headers: dict[str, int] = {}
    for cell in ws[1]:
        name = normalize_header(cell.value)
        if not name:
            continue
        if name in headers:
            if name in COPY_COLUMNS:
                LOG.warning("源表重复表头 %s：取第 %s 列，忽略第 %s 列", name, headers[name], cell.column)
        else:
            headers[name] = cell.column
    return headers


def destination_headers(ws: Any) -> tuple[int, dict[str, int]]:
    last_col = max(1, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)
    values = ws.Range(ws.Cells(1, 1), ws.Cells(10, last_col)).Value2
    best_row, best_headers, best_score = 0, {}, -1
    expected = set(COPY_COLUMNS) | set(FORMULA_COLUMNS)
    for row_number, row in enumerate(values, start=1):
        headers: dict[str, int] = {}
        for col_number, value in enumerate(row, start=1):
            name = normalize_header(value)
            if name and name not in headers:
                headers[name] = col_number
        score = len(expected.intersection(headers))
        if "主单号" in headers and score > best_score:
            best_row, best_headers, best_score = row_number, headers, score
    if best_score < 2:
        raise ValueError(f"{ws.Name} 前十行没有找到有效表头")
    return best_row, best_headers


def european_copy_fields(headers: dict[str, int]) -> tuple[str, ...]:
    """欧线复制双方同名的业务列与费用列，跳过目标公式和利润列。"""
    omitted = {"序号", "月份", "周数", "主单毛利", "毛利率", "主单毛利率",
               "业务员毛利", "业务毛利率", "应付-其他费用"}
    return tuple(field for field in headers if field not in omitted)


def theme_colors(workbook: Any) -> list[str]:
    if not workbook.loaded_theme:
        return []
    root = ElementTree.fromstring(workbook.loaded_theme)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    if scheme is None:
        return []
    colors = []
    for element in scheme:
        child = next(iter(element), None)
        colors.append((child.attrib.get("lastClr") or child.attrib.get("val") or "000000")[-6:]
                      if child is not None else "000000")
    if len(colors) >= 2:
        colors[0], colors[1] = colors[1], colors[0]
    return colors


def fill_rgb(cell: Any, theme: list[str]) -> str | None:
    fill = cell.fill
    if fill is None or fill.patternType is None:
        return None
    color = fill.fgColor
    rgb = None
    if color.type == "rgb" and isinstance(color.rgb, str):
        rgb = color.rgb[-6:]
    elif color.type == "theme" and color.theme is not None and color.theme < len(theme):
        rgb = theme[color.theme]
    elif color.type == "indexed" and color.indexed is not None and color.indexed < len(COLOR_INDEX):
        rgb = COLOR_INDEX[color.indexed][-6:]
    if rgb is None:
        LOG.warning("%s 的填充色无法解析：%s", cell.coordinate, color.type)
        return None
    tint = color.tint or 0
    components = [int(rgb[i:i + 2], 16) for i in (0, 2, 4)]
    adjusted = [round(value * (1 + tint) if tint < 0 else value * (1 - tint) + 255 * tint)
                for value in components]
    return "".join(f"{max(0, min(255, value)):02X}" for value in adjusted)


def excel_color(rgb: str) -> int:
    return int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)


def excel_value(value: Any, epoch: Any, field: str) -> Any:
    if value is None:
        return None
    if field == "创建时间" and isinstance(value, str):
        try:
            value = datetime.fromisoformat(value.strip())
        except ValueError:
            return value
    if isinstance(value, (datetime, date)):
        return to_excel(value, epoch=epoch)
    return value


def read_sheet_records(workbook: Any, sheet_name: str, target_headers: dict[str, int],
                       copy_fields: tuple[str, ...]) -> list[dict[str, tuple[Any, str | None]]]:
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"缺少“{sheet_name}”工作表")
    ws = workbook[sheet_name]
    headers = source_headers(ws)
    if "主单号" not in headers:
        raise ValueError(f"“{sheet_name}”缺少“主单号”列")
    missing = [field for field in copy_fields if field in target_headers and field not in headers]
    if missing:
        LOG.warning("%s 缺少列，目标新增行留空：%s", sheet_name, "、".join(missing))
    theme = theme_colors(workbook)
    needed = {field: headers[field] for field in copy_fields if field in headers and field in target_headers}
    last_col = max(needed.values(), default=headers["主单号"])
    records = []
    source_errors: Counter[str] = Counter()
    for cells in ws.iter_rows(min_row=2, max_col=last_col):
        if cells[headers["主单号"] - 1].value in (None, ""):
            continue
        record = {}
        for field, col in needed.items():
            cell = cells[col - 1]
            if getattr(cell, "data_type", None) == "e":
                source_errors[f"{field}={cell.value}"] += 1
            record[field] = (excel_value(cell.value, workbook.epoch, field), fill_rgb(cell, theme))
        records.append(record)
    if source_errors:
        LOG.warning("%s 源数据含 Excel 错误：%s", sheet_name,
                    "、".join(f"{field} {count} 个" for field, count in source_errors.items()))
    LOG.info("%s：读取 %s 条有效数据", sheet_name, len(records))
    return records


def read_source(path: Path, target_headers: dict[str, int]) -> list[dict[str, tuple[Any, str | None]]]:
    LOG.info("读取源文件：%s", path)
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        return read_sheet_records(workbook, "汇总", target_headers, COPY_COLUMNS)
    finally:
        workbook.close()


def unique_output(path: Path, folder: Path) -> Path:
    candidate = folder / path.name
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime("%H%M%S")
    candidate = folder / f"{path.stem}_{stamp}{path.suffix}"
    index = 2
    while candidate.exists():
        candidate = folder / f"{path.stem}_{stamp}_{index}{path.suffix}"
        index += 1
    return candidate


def publish_output(staged: Path, final: Path) -> None:
    """发布本次结果；同日旧结果移入历史版本，发布失败则恢复旧结果。"""
    if not final.exists():
        staged.replace(final)
        return
    archive_dir = final.parent / "历史版本"
    archive_dir.mkdir(exist_ok=True)
    stamped = final.with_name(f"{final.stem}_{datetime.now():%Y%m%d_%H%M%S}{final.suffix}")
    archive = unique_output(stamped, archive_dir)
    final.replace(archive)
    try:
        staged.replace(final)
    except Exception:
        archive.replace(final)
        raise
    LOG.info("同日旧结果已移入：%s", archive)


def fill_column(ws: Any, column: int, start: int,
                records: list[dict[str, tuple[Any, str | None]]], field: str) -> None:
    run_start = start
    previous = None
    for offset in range(len(records) + 1):
        color = records[offset].get(field, (None, None))[1] if offset < len(records) else "__END__"
        if offset == 0:
            previous = color
            continue
        if color == previous:
            continue
        area = ws.Range(ws.Cells(run_start, column), ws.Cells(start + offset - 1, column))
        if previous is None:
            area.Interior.Pattern = XL_NONE
        else:
            area.Interior.Pattern = 1
            area.Interior.Color = excel_color(previous)
        run_start = start + offset
        previous = color


def append_rows(ws: Any, records: list[dict[str, tuple[Any, str | None]]],
                header_row: int, headers: dict[str, int], *, europe: bool = False) -> tuple[int, int]:
    required = {"创建时间", "主单号", "服务商总应付", "月份", "周数"}
    if not europe:
        required.add("应付-其他费用")
    missing = sorted(required - headers.keys())
    if missing:
        raise ValueError(f"{ws.Name} 缺少必要列：{', '.join(missing)}")
    if not records:
        return 0, 0
    key_col = headers["主单号"]
    # 欧线目标表预留大量空白格式行，按主单号的最后一条实际记录续接。
    # 加美线仍参考 UsedRange，兼容旧表中 End(xlUp) 停在中途的情况。
    key_last = ws.Cells(ws.Rows.Count, key_col).End(XL_UP).Row
    used_last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    last_row = key_last if europe else max(key_last, used_last)
    if last_row < header_row:
        last_row = header_row
    start = last_row + 1
    end = start + len(records) - 1
    if end > ws.Rows.Count:
        raise ValueError(f"{ws.Name} 超出 Excel 最大行数")
    last_col = max(headers.values())
    if last_row > header_row:
        template = ws.Range(ws.Cells(last_row, 1), ws.Cells(last_row, last_col))
        destination = ws.Range(ws.Cells(start, 1), ws.Cells(end, last_col))
        template.Copy()
        destination.PasteSpecial(Paste=XL_PASTE_FORMATS)
        ws.Application.CutCopyMode = False
        height = ws.Rows(last_row).RowHeight
        if height is not None:
            destination.RowHeight = height
    copy_fields = european_copy_fields(headers) if europe else COPY_COLUMNS
    for field in copy_fields:
        col = headers.get(field)
        if col is None:
            LOG.warning("%s 目标表缺少列，跳过：%s", ws.Name, field)
            continue
        values = tuple((record.get(field, (None, None))[0],) for record in records)
        area = ws.Range(ws.Cells(start, col), ws.Cells(end, col))
        if field in {"主单号", "柜号", "配仓单号"}:
            area.NumberFormat = "@"  # 主单号可能全是数字，保留其文本类型和前导零。
        area.Value2 = values
        fill_column(ws, col, start, records, field)
    date_col = headers["创建时间"]
    for field in ("月份", "周数"):
        col = headers[field]
        previous = ws.Cells(start - 1, col).FormulaR1C1 if start - 1 > header_row else None
        offset = date_col - col
        fallback = (f'=YEAR(RC[{offset}])&"-"&MONTH(RC[{offset}])' if field == "月份"
                    else f'="WK-"&WEEKNUM(RC[{offset}],2)')
        formula = previous if isinstance(previous, str) and previous.startswith("=") else fallback
        ws.Range(ws.Cells(start, col), ws.Cells(end, col)).FormulaR1C1 = formula
    profit_col = headers.get("主单毛利")
    if profit_col is not None:
        previous_profit = ws.Cells(last_row, profit_col).FormulaR1C1 if last_row > header_row else None
        if not isinstance(previous_profit, str) or not previous_profit.startswith("="):
            raise ValueError(f"{ws.Name} 第 {last_row} 行“主单毛利”不是公式，无法下拉")
        ws.Range(ws.Cells(start, profit_col), ws.Cells(end, profit_col)).FormulaR1C1 = previous_profit
        LOG.info("%s：“主单毛利”沿用第 %s 行公式，填充第 %s 至 %s 行", ws.Name, last_row, start, end)
    else:
        LOG.warning("%s 目标表没有“主单毛利”列，跳过公式下拉", ws.Name)
    fee_cols = []
    if "应付-其他费用" in headers:
        total_col = headers["服务商总应付"]
        other_col = headers["应付-其他费用"]
        fee_cols = [col for field, col in headers.items()
                    if "应付" in field and field not in {"服务商总应付", "应付-其他费用"}]
        terms = ",".join(f"RC[{col - other_col}]" for col in sorted(fee_cols))
        formula = f"=RC[{total_col - other_col}]-SUM({terms})" if terms else f"=RC[{total_col - other_col}]"
        ws.Range(ws.Cells(start, other_col), ws.Cells(end, other_col)).FormulaR1C1 = formula
    elif europe:
        LOG.warning("%s 目标表没有“应付-其他费用”列，跳过该公式", ws.Name)
    if europe and "序号" in headers:
        col = headers["序号"]
        ws.Range(ws.Cells(start, col), ws.Cells(end, col)).Value2 = tuple(
            (row,) for row in range(start - header_row, end - header_row + 1)
        )
    ws.Range(ws.Cells(start, 1), ws.Cells(end, last_col)).Calculate()
    LOG.info("%s：从第 %s 行追加到第 %s 行，共 %s 条；其他费用扣除 %s 个应付列",
             ws.Name, start, end, len(records), len(fee_cols))
    return start, end


def merge_pair(excel: Any, target: ReportFile, sources: list[ReportFile], output_dir: Path) -> tuple[Path, int]:
    LOG.info("处理 %s 线：目标 %s，源文件 %s 个", target.line, target.path.name, len(sources))
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / target.path.name
    staged_name = target.path.with_name(f"{target.path.stem}_处理中{target.path.suffix}")
    staged_path = unique_output(staged_name, output_dir)
    shutil.copyfile(target.path, staged_path)
    workbook = None
    saved = False
    try:
        workbook = excel.Workbooks.Open(str(staged_path), UpdateLinks=0, ReadOnly=False,
                                        IgnoreReadOnlyRecommended=True)
        if workbook.ReadOnly:
            raise PermissionError(f"临时结果被 Excel 以只读方式打开：{staged_path}")
        sheets = [sheet for sheet in workbook.Worksheets
                  if "明细表" in sheet.Name and target.line in sheet.Name]
        if len(sheets) != 1:
            raise ValueError(f"{target.path.name} 中应有一个含“{target.line}”的“明细表”，实际 {len(sheets)} 个")
        ws = sheets[0]
        header_row, headers = destination_headers(ws)
        records = []
        for item in sources:
            try:
                records.extend(read_source(item.path, headers))
            except Exception:
                LOG.exception("读取失败，跳过源文件：%s", item.path)
        if not records:
            raise ValueError(f"{target.line}线没有可追加的数据")
        start, end = append_rows(ws, records, header_row, headers)
        workbook.Save()
        saved = True
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if not saved and staged_path.exists():
            staged_path.unlink()
    try:
        publish_output(staged_path, output_path)
    except Exception:
        if staged_path.exists():
            staged_path.unlink()
        raise
    LOG.info("已保存：%s", output_path)
    return output_path, end - start + 1


def merge_europe(excel: Any, target: ReportFile, sources: list[ReportFile],
                 output_dir: Path) -> tuple[Path, int]:
    LOG.info("处理欧线：目标 %s，源文件 %s 个，工作表 %s 个", target.path.name, len(sources), len(EUROPE_SHEETS))
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / target.path.name
    staged_name = target.path.with_name(f"{target.path.stem}_处理中{target.path.suffix}")
    staged_path = unique_output(staged_name, output_dir)
    shutil.copyfile(target.path, staged_path)
    workbook = None
    saved = False
    total = 0
    try:
        workbook = excel.Workbooks.Open(str(staged_path), UpdateLinks=0, ReadOnly=False,
                                        IgnoreReadOnlyRecommended=True)
        if workbook.ReadOnly:
            raise PermissionError(f"临时结果被 Excel 以只读方式打开：{staged_path}")
        target_sheets = {sheet.Name: sheet for sheet in workbook.Worksheets}
        missing_target = [name for name in EUROPE_SHEETS if name not in target_sheets]
        if missing_target:
            raise ValueError(f"利润报表缺少工作表：{'、'.join(missing_target)}")
        source_books = []
        try:
            for source in sources:
                LOG.info("读取欧线源文件：%s", source.path)
                book = openpyxl.load_workbook(source.path, read_only=True,
                                              data_only=True, keep_links=False)
                source_books.append((source, book))
            for name in EUROPE_SHEETS:
                ws = target_sheets[name]
                header_row, headers = destination_headers(ws)
                copy_fields = european_copy_fields(headers)
                records = []
                for source, book in source_books:
                    if name not in book.sheetnames:
                        LOG.warning("%s 缺少“%s”工作表，跳过该源表", source.path.name, name)
                        continue
                    records.extend(read_sheet_records(book, name, headers, copy_fields))
                if not records:
                    LOG.info("%s：没有可追加数据", name)
                    continue
                start, end = append_rows(ws, records, header_row, headers, europe=True)
                total += end - start + 1
        finally:
            for _, book in source_books:
                book.close()
        if not total:
            raise ValueError("欧线九张工作表均没有可追加数据")
        workbook.Save()
        saved = True
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if not saved and staged_path.exists():
            staged_path.unlink()
    try:
        publish_output(staged_path, output_path)
    except Exception:
        if staged_path.exists():
            staged_path.unlink()
        raise
    LOG.info("欧线已保存：%s，合计追加 %s 条", output_path, total)
    return output_path, total


def process(files: list[ReportFile]) -> tuple[list[tuple[Path, int]], list[str], Path | None]:
    pairs, skipped = get_pairs(files)
    if not pairs:
        return [], skipped or ["没有可配对的报表"], None
    day = datetime.now().strftime("%Y-%m-%d")
    log_path = setup_log(pairs[0][0].path.parent / day)
    LOG.info("开始处理。选中文件数：%s；各结果保存在对应利润报表原目录的 %s 文件夹", len(files), day)
    for message in skipped:
        LOG.warning(message)
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("缺少 pywin32，请安装后再运行") from exc
    results: list[tuple[Path, int]] = []
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        for target, sources in pairs:
            try:
                merge = merge_europe if target.line == "欧" else merge_pair
                results.append(merge(excel, target, sources, target.path.parent / day))
            except Exception as exc:
                LOG.exception("%s 线处理失败", target.line)
                skipped.append(f"{target.line}线处理失败：{exc}")
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    LOG.info("完成：产出 %s 个文件，合计追加 %s 条；跳过/失败 %s 项。日志：%s",
             len(results), sum(count for _, count in results), len(skipped), log_path)
    return results, skipped, log_path


def show_notice(title: str, message: str, *, error: bool = False) -> None:
    """在前台显示系统提示，避免提示被已隐藏的文件选择窗口挡住。"""
    flags = 0x00040000 | 0x00010000 | (0x00000010 if error else 0x00000040)
    ctypes.windll.user32.MessageBoxW(0, message, title, flags)


def run_gui() -> int:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        chosen = filedialog.askopenfilenames(
            parent=root,
            title="选择预估报表和利润报表（按 Ctrl 或 Shift 多选）",
            initialdir=str(ROOT if ROOT.is_dir() else Path.home()),
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not chosen:
            return 0
        # 文件选择结束后立刻关闭隐藏的 Tk 根窗口，后续不再给提示框设置隐藏窗口的 owner。
        root.destroy()
        root = None
        files: list[ReportFile] = []
        seen: set[str] = set()
        ignored: list[str] = []
        for raw_path in chosen:
            path = Path(raw_path)
            item = identify(path)
            if item is None:
                ignored.append(path.name)
                continue
            key = str(item.path).casefold()
            if key not in seen:
                files.append(item)
                seen.add(key)
        pairs, skipped = get_pairs(files)
        if not pairs:
            reason = "；".join(skipped) or "文件名需要包含线路和报表类型。"
            show_notice("无法配对", "请选择同一线路的预估报表和利润报表。" + chr(10) + reason, error=True)
            return 1
        try:
            outputs, skipped, log_path = process(files)
        except Exception as exc:
            show_notice("处理失败", str(exc), error=True)
            return 1
        lines = [f"已生成 {len(outputs)} 个文件，共追加 {sum(n for _, n in outputs)} 条数据。"]
        lines += [f"{path.name}：{count} 条" for path, count in outputs]
        if ignored:
            lines.append("未识别，已跳过：" + "、".join(ignored))
        if skipped:
            lines.append("跳过/失败：" + "；".join(skipped))
        if log_path:
            lines.append(f"日志：{log_path}")
        show_notice("处理完成", chr(10).join(lines))
        for folder in dict.fromkeys(path.parent for path, _ in outputs):
            os.startfile(str(folder))
        return 0 if outputs and not any("失败" in item for item in skipped) else 1
    finally:
        if root is not None:
            root.destroy()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--all", action="store_true", help="处理扫描到的所有报表，不打开选择窗口")
    parser.add_argument("--source-dir", type=Path, default=SOURCE_DIR, help="预估报表目录")
    parser.add_argument("--report-dir", type=Path, default=REPORT_DIR, help="利润报表目录")
    args = parser.parse_args(argv)
    try:
        if not args.all:
            return run_gui()
        files = discover(args.source_dir, args.report_dir)
        if not files:
            raise ValueError("没有找到文件名含“预估报表”或“利润报表”的 Excel 文件")
        outputs, skipped, log_path = process(files)
        print(f"完成：{len(outputs)} 个文件，{sum(n for _, n in outputs)} 条数据；日志：{log_path}")
        return 0 if outputs and not any("失败" in item for item in skipped) else 1
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
