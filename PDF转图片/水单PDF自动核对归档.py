# -*- coding: utf-8 -*-
"""按金额匹配报销/付款 PDF 与水单，并将唯一匹配项自动归档。

运行：python 水单PDF自动核对归档.py
依赖：PyMuPDF、Pillow、rapidocr-onnxruntime（扫描件或图片时才加载 OCR）。
"""
from __future__ import annotations

import csv
import os
import re
import shutil
import sys
import tempfile
import traceback
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import fitz
import cv2


SOURCE_SUFFIXES = {'.pdf'}
IMAGE_SUFFIXES = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}
MONEY_TOKEN = r'[-－]?\s*(?:CNY|USD)?\s*[¥￥$]?\s*[0-9][0-9,，]*(?:\.[0-9]{1,2})?'
NEGATIVE_MONEY_TOKEN = (r'[-－—–−]\s*(?:(?:CNY|人民币)\s*)?'
                        r'[¥￥]?\s*[0-9][0-9,，]*(?:\.[0-9]{1,2})?')
# 仅着色形如 54,466.00 或 -164.70 的两位小数金额，
# 不把处理序号、日期或 “1789....png” 这类文件名数字误标。
LOG_AMOUNT_TOKEN = r'(?<![A-Za-z0-9])[－—–−-]?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{1,2}(?![A-Za-z0-9])'
ANSI_BOLD_GREEN = '\033[1;92m'
ANSI_RESET = '\033[0m'


@dataclass(frozen=True)
class Record:
    path: Path
    amount: Decimal
    kind: str


def log(message: str) -> None:
    """实时输出；日志中的金额统一显示为亮绿色加粗。"""
    colored = re.sub(LOG_AMOUNT_TOKEN,
                     lambda match: ANSI_BOLD_GREEN + match.group(0) + ANSI_RESET,
                     str(message))
    print(colored, flush=True)


def enable_console_color() -> None:
    """为传统 Windows 控制台启用 ANSI 颜色，失败时仍保留正常日志。"""
    if os.name != 'nt':
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        mode = ctypes.c_ulong()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass


def display_path(path: Path) -> str:
    return str(path)


def compact(text: str) -> str:
    return re.sub(r'\s+', '', text or '').replace('，', ',').replace('－', '-')


def parse_amount(value: str) -> Decimal | None:
    """将 OCR/PDF 中带币种、千分位、负号的金额统一为两位 Decimal。"""
    if not value:
        return None
    clean = re.sub(r'[^0-9.\-]', '', value.replace('－', '-'))
    if clean.count('-') > 1 or clean.startswith('-') is False and '-' in clean:
        return None
    try:
        return Decimal(clean).quantize(Decimal('0.01'))
    except (InvalidOperation, ValueError):
        return None


def amounts_after(text: str, labels: tuple[str, ...], limit: int = 80) -> list[Decimal]:
    normalized = compact(text)
    values: list[Decimal] = []
    for label in labels:
        for match in re.finditer(re.escape(label) + r'[^0-9A-Za-z$¥￥－-]{0,12}(' + MONEY_TOKEN + r')', normalized, re.I):
            amount = parse_amount(match.group(1))
            if amount is not None:
                values.append(amount)
    return list(dict.fromkeys(values))


SOURCE_TABLE_AMOUNT_LABELS = ('报销金额', '费用金额', '借款金额', '申请金额')


def source_table_amounts(text: str) -> list[Decimal]:
    """取四类来源表格金额，适配 PDF 表格文字层与 OCR。

    PDF 表格的文字层有时按列、而非按视觉行输出；因此除了标题后的
    四个文本行，还在表头后的有限文本范围内兜底找带小数/千分位的金额。
    """
    lines = [line.strip() for line in (text or '').splitlines() if line.strip()]
    values: list[Decimal] = []
    for index, line in enumerate(lines):
        if not any(label in compact(line) for label in SOURCE_TABLE_AMOUNT_LABELS):
            continue
        # 正常阅读顺序：仅检查标题下方四行。
        below = compact(' '.join(lines[index + 1:index + 5]))
        for match in re.finditer(MONEY_TOKEN, below, re.I):
            token = match.group(0)
            # 日期、编号等纯整数不应作为报销金额；金额应有小数、千分位或币种符号。
            if not any(mark in token.upper() for mark in ('.', ',', 'CNY', '¥', '￥', '$')):
                continue
            amount = parse_amount(token)
            if amount is not None:
                values.append(amount)
                break

    # 表格 PDF 可能将表头与金额分列输出，例如“费用金额”与下方的 450.00元。
    # 在每个表头后的 300 个字符内兜底，仍要求金额有小数点或千分位，
    # 排除日期、单据张数等纯数字。
    normalized = compact(text)
    for label in SOURCE_TABLE_AMOUNT_LABELS:
        for header in re.finditer(re.escape(label) + r'(?:[（(]元?[）)])?', normalized):
            area = normalized[header.end():header.end() + 300]
            for match in re.finditer(MONEY_TOKEN, area, re.I):
                token = match.group(0)
                if '.' not in token and ',' not in token:
                    continue
                amount = parse_amount(token)
                if amount is not None:
                    values.append(amount)
                    break
    return list(dict.fromkeys(values))


def cny_receipt_amounts(text: str) -> list[tuple[Decimal, str]]:
    """返回人民币水单候选；类型顺序与业务规则一致。"""
    normalized = compact(text)
    found: list[tuple[Decimal, str]] = []
    # 私账水单，标签最明确，必须优先判断。
    for amount in amounts_after(normalized, ('汇款金额小写',)):
        found.append((amount, '人民币-私账'))
    # 类型 2：APP 支付。OCR 常把短横识别为全角横线、长横线或数学减号；
    # “金额”与负数也可能被分成多行，因此在金额字段后 100 个字符内查找。
    for label in re.finditer(r'金额', normalized):
        field_text = normalized[label.end():label.end() + 100]
        match = re.search(NEGATIVE_MONEY_TOKEN, field_text, re.I)
        if match:
            amount = parse_amount(match.group(0))
            if amount is not None:
                # APP 支付水单以负数表示支出，来源 PDF 的付款总额是正数。
                found.append((abs(amount), '人民币-APP'))
    # 微信/零钱通“支付成功”和支付宝“交易成功”截图，都将付款金额
    # 单独以大字放在页面中央，没有“金额”字段。
    if re.search(r'支付成功|交易成功|扫码付款|微信支付|支付宝|零钱通', normalized):
        for match in re.finditer(NEGATIVE_MONEY_TOKEN, normalized, re.I):
            amount = parse_amount(match.group(0))
            if amount is not None:
                found.append((abs(amount), '人民币-APP支付成功'))
    # 类型 1：常规水单。PDF/图片 OCR 的条目顺序不一定等于视觉顺序：
    # 有些水单会先读到“CNY 2,850.00”，最后才读到左侧的“金额”。
    # 因此保留原有“金额 -> CNY”规则，并增加“同图含金额字段 + CNY 金额”的兜底。
    for match in re.finditer(r'金额.{0,50}?CNY\s*([¥￥]?\s*[0-9][0-9,]*(?:\.[0-9]{1,2})?)', normalized, re.I):
        amount = parse_amount(match.group(1))
        if amount is not None:
            found.append((amount, '人民币-常规'))
    if '金额' in normalized:
        for match in re.finditer(r'C\s*N\s*Y\s*[¥￥]?\s*([0-9][0-9,，]*(?:\.[0-9]{1,2})?)', text or '', re.I):
            amount = parse_amount(match.group(1))
            if amount is not None:
                found.append((amount, '人民币-常规'))
    # 另一种普通银行水单直接显示“付款金额：54,466.00”，无需 CNY 标记。
    for amount in amounts_after(normalized, ('付款金额',)):
        found.append((amount, '人民币-常规付款金额'))
    # 类型 1 还会出现“金额（小写）：8,300.00”及“汇款金额：200.00元”。
    # 括号可因 PDF/OCR 变成全角或半角，compact 后均可统一识别。
    for amount in amounts_after(normalized, ('金额（小写）', '金额(小写)', '金额小写')):
        found.append((amount, '人民币-常规金额小写'))
    for amount in amounts_after(normalized, ('汇款金额',)):
        found.append((amount, '人民币-常规汇款金额'))
    return list(dict.fromkeys(found))


def usd_receipt_amounts(text: str, is_citic: bool, is_cmb: bool) -> list[tuple[Decimal, str]]:
    normalized = compact(text)
    found: list[tuple[Decimal, str]] = []
    if is_citic:
        for amount in amounts_after(normalized, ('现汇金额', '购汇金额')):
            found.append((amount, '美元-中信银行'))
    if is_cmb:
        for match in re.finditer(r'交易金额.{0,50}?USD\s*([$]?\s*[0-9][0-9,]*(?:\.[0-9]{1,2})?)', normalized, re.I):
            amount = parse_amount(match.group(1))
            if amount is not None:
                found.append((amount, '美元-招商银行'))
    return list(dict.fromkeys(found))


def get_ocr():
    from rapidocr_onnxruntime import RapidOCR
    return RapidOCR()


_OCR = None


def enhanced_ocr_image(image_path: Path) -> Path | None:
    """为单张识别失败的细小水单生成轻量 OCR 临时图。"""
    try:
        image = cv2.imread(str(image_path), cv2.IMREAD_GRAYSCALE)
        if image is None:
            return None
        # 仅用于单张漏识别图片：放大 1.5 倍、局部对比度增强、Otsu 二值化。
        # 对“费用金额”“汇款金额”这类细字加表格线，通常比原彩色截图更清晰。
        enlarged = cv2.resize(image, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(enlarged)
        _, enhanced = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        temp = Path(tempfile.gettempdir()) / f'water-slip-enhanced-{os.getpid()}-{image_path.stem}.png'
        cv2.imwrite(str(temp), enhanced)
        return temp
    except Exception:
        return None


def ocr_text(image_path: Path, retry_enhanced: bool = False) -> str:
    global _OCR
    if _OCR is None:
        log('正在加载 OCR 识别模型，首次加载可能需要几十秒…')
        _OCR = get_ocr()
        log('OCR 识别模型加载完成。')
    if not retry_enhanced:
        result, _ = _OCR(str(image_path))
        return '\n'.join(str(row[1]) for row in (result or []) if len(row) > 1)

    enhanced_path = enhanced_ocr_image(image_path)
    try:
        result, _ = _OCR(str(enhanced_path or image_path))
        return '\n'.join(str(row[1]) for row in (result or []) if len(row) > 1)
    finally:
        if enhanced_path:
            enhanced_path.unlink(missing_ok=True)


def pdf_text(path: Path, retry_enhanced: bool = False) -> str:
    """优先原生文字层；扫描 PDF 没有文字时逐页渲染后 OCR。"""
    doc = fitz.open(path)
    temporary: list[Path] = []
    try:
        native = '\n'.join(page.get_text('text') for page in doc)
        # 常规情况优先用 PDF 文字层；来源金额没取到而重试时，
        # 强制渲染整页后 OCR，避免重复读取同一份表格乱序文字层。
        if compact(native) and not retry_enhanced:
            return native
        for number, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            temp = Path(tempfile.gettempdir()) / f'water-slip-ocr-{os.getpid()}-{number}.png'
            pix.save(str(temp))
            temporary.append(temp)
        return '\n'.join(ocr_text(path, retry_enhanced=retry_enhanced) for path in temporary)
    finally:
        doc.close()
        for temp in temporary:
            temp.unlink(missing_ok=True)


def read_text(path: Path, retry_enhanced: bool = False) -> str:
    if path.suffix.lower() == '.pdf':
        return pdf_text(path, retry_enhanced=retry_enhanced)
    return ocr_text(path, retry_enhanced=retry_enhanced)


def source_records(pdf_dir: Path) -> tuple[list[Record], list[Record], list[dict[str, str]]]:
    cny: list[Record] = []
    usd: list[Record] = []
    report: list[dict[str, str]] = []
    # 只扫描用户选择的当前文件夹，不进入任何子文件夹。
    files = [path for path in sorted(pdf_dir.iterdir())
             if path.is_file() and path.suffix.lower() in SOURCE_SUFFIXES]
    log(f'第一步完成扫描（不扫描子文件夹）：发现 {len(files)} 份来源 PDF，开始识别金额。')
    for index, path in enumerate(files, 1):
        try:
            log(f'[PDF {index}/{len(files)}] 正在识别：{path.name}')
            text = read_text(path)
            relative_parts = path.relative_to(pdf_dir).parts
            is_usd_pdf = ('美金' in relative_parts or '美元' in relative_parts
                          or bool(re.search(r'\bUSD\b', text, re.I)))
            # 来源 PDF 支持付款总额/汇款金额，以及四类表格金额的右侧或下方取值。
            table_amount_values = source_table_amounts(text) if not is_usd_pdf else []
            cny_values = (amounts_after(text, ('付款总额', '汇款金额') + SOURCE_TABLE_AMOUNT_LABELS)
                          + table_amount_values) if not is_usd_pdf else []
            cny_values = list(dict.fromkeys(cny_values))
            usd_values = amounts_after(text, ('付款总额',)) if is_usd_pdf else []
            if not cny_values and not usd_values:
                # 扫描版付款申请中的细字表格先原图 OCR，失败后才走局部对比度增强。
                log('  原图未识别到来源金额，正在增强后重试…')
                text = read_text(path, retry_enhanced=True)
                is_usd_pdf = ('美金' in relative_parts or '美元' in relative_parts
                              or bool(re.search(r'\bUSD\b', text, re.I)))
                table_amount_values = source_table_amounts(text) if not is_usd_pdf else []
                cny_values = (amounts_after(text, ('付款总额', '汇款金额') + SOURCE_TABLE_AMOUNT_LABELS)
                              + table_amount_values) if not is_usd_pdf else []
                cny_values = list(dict.fromkeys(cny_values))
                usd_values = amounts_after(text, ('付款总额',)) if is_usd_pdf else []
            cny.extend(Record(path, value, '人民币') for value in cny_values)
            usd.extend(Record(path, value, '美元') for value in usd_values)
            if not cny_values and not usd_values:
                log('  未识别到待匹配金额。')
                report.append(row(path, '', '', '未识别到付款总额、报销金额或汇款金额'))
            else:
                values = cny_values or usd_values
                log('  识别金额：' + '、'.join(str(value) for value in values))
                if table_amount_values:
                    log('  表格金额候选：' + '、'.join(str(value) for value in table_amount_values))
        except Exception as exc:
            log(f'  识别失败：{exc}')
            report.append(row(path, '', '', f'源 PDF 读取失败：{exc}'))
    log(f'来源 PDF 识别完成：人民币候选 {len(cny)} 条，美元候选 {len(usd)} 条。')
    return cny, usd, report


def receipt_records(receipt_dir: Path) -> tuple[list[Record], list[Record], list[dict[str, str]]]:
    cny: list[Record] = []
    usd: list[Record] = []
    report: list[dict[str, str]] = []
    def is_receipt_file(path: Path) -> bool:
        if not path.is_file() or path.suffix.lower() not in IMAGE_SUFFIXES | SOURCE_SUFFIXES:
            return False
        # 同一目录可能保留了付款申请 PDF 的转图/截图副本，文件名例如
        # “张三提交的付款申请202609....png”。它是来源单据而非银行水单，
        # 不能因扩展名为 PNG 就参与水单金额匹配。
        if path.suffix.lower() in IMAGE_SUFFIXES and '提交的付款申请' in path.stem:
            return False
        # 人民币水单是图片；美元水单 PDF 必须直接放在所选目录，
        # 并由文件名或正文中的银行标识判断。
        if path.suffix.lower() != '.pdf':
            return True
        return '中信银行' in path.name or '招商银行' in path.name

    # 只扫描用户选择的当前文件夹，不进入任何子文件夹。
    files = [path for path in sorted(receipt_dir.iterdir()) if is_receipt_file(path)]
    log(f'第二步完成扫描（不扫描子文件夹）：发现 {len(files)} 份水单图片/PDF，开始 OCR 识别。')
    for index, path in enumerate(files, 1):
        try:
            log(f'[水单 {index}/{len(files)}] 正在识别：{path.name}')
            text = read_text(path)
            cny_values = cny_receipt_amounts(text)
            for amount, kind in cny_values:
                cny.append(Record(path, amount, kind))
            file_name = path.name.lower()
            is_citic = '中信银行' in path.name or 'citic' in file_name
            is_cmb = '招商银行' in path.name or 'cmb' in file_name
            usd_values = usd_receipt_amounts(text, is_citic, is_cmb)
            for amount, kind in usd_values:
                usd.append(Record(path, amount, kind))
            values = cny_values + usd_values
            if not values and path.suffix.lower() in IMAGE_SUFFIXES:
                log('  原图未识别到水单金额，正在轻量增强后重试…')
                text = ocr_text(path, retry_enhanced=True)
                cny_values = cny_receipt_amounts(text)
                usd_values = usd_receipt_amounts(text, is_citic, is_cmb)
                for amount, kind in cny_values:
                    cny.append(Record(path, amount, kind))
                for amount, kind in usd_values:
                    usd.append(Record(path, amount, kind))
                values = cny_values + usd_values
            if values:
                log('  识别金额：' + '、'.join(f'{kind} {amount}' for amount, kind in values))
            else:
                negative_values = [parse_amount(match.group(0))
                                   for match in re.finditer(NEGATIVE_MONEY_TOKEN, compact(text), re.I)]
                negative_values = [abs(value) for value in negative_values if value is not None]
                if negative_values:
                    log('  检测到负数金额 ' + '、'.join(str(value) for value in negative_values)
                        + '，但未能确认为“金额”字段或 APP 成功交易页面。')
                log('  未识别到符合水单规则的金额。')
        except Exception as exc:
            log(f'  识别失败：{exc}')
            report.append(row(path, '', '', f'水单读取失败：{exc}'))
    log(f'水单识别完成：人民币候选 {len(cny)} 条，美元候选 {len(usd)} 条。')
    return cny, usd, report


def row(source: Path | str, receipt: Path | str, amount: Decimal | str, result: str) -> dict[str, str]:
    return {'来源PDF': str(source), '水单': str(receipt), '金额': str(amount), '处理结果': result}


def unique_pairs(sources: list[Record], receipts: list[Record], report: list[dict[str, str]]) -> list[tuple[Record, Record]]:
    by_amount_sources: dict[Decimal, list[Record]] = defaultdict(list)
    by_amount_receipts: dict[Decimal, list[Record]] = defaultdict(list)
    for item in sources:
        by_amount_sources[item.amount].append(item)
    for item in receipts:
        by_amount_receipts[item.amount].append(item)
    pairs = []
    for amount in sorted(set(by_amount_sources) & set(by_amount_receipts)):
        left = list({item.path: item for item in by_amount_sources[amount]}.values())
        right = list({item.path: item for item in by_amount_receipts[amount]}.values())
        if len(left) == len(right) == 1:
            log(f'金额 {amount} 唯一匹配：{left[0].path.name} <-> {right[0].path.name}')
            pairs.append((left[0], right[0]))
        else:
            log(f'金额 {amount} 有 {len(left)} 份 PDF、{len(right)} 份水单候选，保留供人工核对。')
            report.append(row('; '.join(str(x.path) for x in left), '; '.join(str(x.path) for x in right), amount, '同金额存在多份候选，未自动移动'))
    return pairs


def safe_target(folder: Path, preferred_name: str) -> Path:
    candidate = folder / preferred_name
    if not candidate.exists():
        return candidate
    stem, suffix = Path(preferred_name).stem, Path(preferred_name).suffix
    index = 2
    while True:
        candidate = folder / f'{stem} ({index}){suffix}'
        if not candidate.exists():
            return candidate
        index += 1


def archive_pair(source: Record, receipt: Record, output_dir: Path) -> None:
    output_dir.mkdir(exist_ok=True)
    source_target = safe_target(output_dir, source.path.name)
    receipt_target = safe_target(output_dir, source.path.stem + receipt.path.suffix.lower())
    shutil.move(str(source.path), str(source_target))
    try:
        shutil.move(str(receipt.path), str(receipt_target))
    except Exception:
        # 归档必须成对：水单移动失败时将 PDF 退回原位。
        shutil.move(str(source_target), str(source.path))
        raise


def write_report(folder: Path, entries: list[dict[str, str]]) -> Path:
    report_path = folder / '水单匹配处理报告.csv'
    with report_path.open('w', encoding='utf-8-sig', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=('来源PDF', '水单', '金额', '处理结果'))
        writer.writeheader()
        writer.writerows(entries)
    return report_path


def choose_folder(root: Tk, title: str) -> Path | None:
    selected = filedialog.askdirectory(parent=root, title=title, mustexist=True)
    return Path(selected) if selected else None


def main() -> None:
    enable_console_color()
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        while True:
            pdf_dir = choose_folder(root, '第一步：选择存放报销/付款 PDF 的主文件夹')
            if not pdf_dir:
                return
            receipt_dir = choose_folder(root, '第二步：选择存放人民币/美元等水单的文件夹')
            if not receipt_dir:
                return
            log('=' * 60)
            log('开始水单 PDF 自动核对')
            log('来源 PDF 文件夹：' + display_path(pdf_dir))
            log('水单文件夹：' + display_path(receipt_dir))
            if pdf_dir == receipt_dir:
                log('提示：两个步骤选择的是同一文件夹，脚本会按扩展名区分 PDF 与图片。')
            log('=' * 60)

            cny_sources, usd_sources, report = source_records(pdf_dir)
            cny_receipts, usd_receipts, receipt_report = receipt_records(receipt_dir)
            report.extend(receipt_report)
            cny_pairs = unique_pairs(cny_sources, cny_receipts, report)
            usd_pairs = unique_pairs(usd_sources, usd_receipts, report)

            moved = 0
            for source, receipt in cny_pairs:
                archive_pair(source, receipt, pdf_dir / '水单正常匹配')
                log(f'已归档到 水单正常匹配：{source.path.name}')
                report.append(row(source.path.name, receipt.path.name, source.amount, f'已归档：水单正常匹配（{receipt.kind}）'))
                moved += 1
            for source, receipt in usd_pairs:
                archive_pair(source, receipt, pdf_dir / '美金正常')
                log(f'已归档到 美金正常：{source.path.name}')
                report.append(row(source.path.name, receipt.path.name, source.amount, f'已归档：美金正常（{receipt.kind}）'))
                moved += 1

            report_path = write_report(pdf_dir, report)
            log(f'处理完成：已归档 {moved} 对文件。')
            log('处理报告：' + display_path(report_path))
            continue_processing = messagebox.askyesno(
                '本批处理完成',
                f'已归档 {moved} 对文件。\n处理报告：\n{report_path}\n\n是否继续处理下一批？',
                parent=root,
            )
            if not continue_processing:
                log('用户选择不继续处理，按回车键关闭此窗口。')
                break
    except Exception as exc:
        traceback.print_exc()
        messagebox.showerror('水单 PDF 自动核对失败', str(exc))
    finally:
        root.destroy()
        if sys.stdin.isatty():
            try:
                input('按回车键关闭…')
            except (EOFError, KeyboardInterrupt):
                pass


if __name__ == '__main__':
    main()
