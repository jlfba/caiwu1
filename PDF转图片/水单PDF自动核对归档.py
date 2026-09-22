# -*- coding: utf-8 -*-
"""按金额匹配报销/付款 PDF 与水单，并将唯一匹配项自动归档。

运行：python 水单PDF自动核对归档.py
依赖：PyMuPDF、Pillow、rapidocr-onnxruntime（扫描件或图片时才加载 OCR）。
"""
from __future__ import annotations

import csv
from datetime import date
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
FOREIGN_CURRENCIES = ('USD', 'CAD', 'GBP', 'EUR')
MONEY_TOKEN = (r'[-－]?\s*(?:CNY|USD)?\s*[¥￥$]?\s*'
               r'(?:\d{1,3}(?:[,，.]\d{3})+[,，.]\d{1,2}'
               r'|\d{1,3}(?:[,，.]\d{3})+|\d+\.\d{1,2}|\d+)')
NEGATIVE_MONEY_TOKEN = (r'[-－—–−]\s*(?:(?:CNY|人民币)\s*)?'
                        r'[¥￥]?\s*(?:\d{1,3}(?:[,，.]\d{3})+[,，.]\d{1,2}'
                        r'|\d{1,3}(?:[,，.]\d{3})+|\d+\.\d{1,2}|\d+)')
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
    """将 OCR/PDF 中带币种、千分位、负号的金额统一为两位 Decimal。

    兼容 OCR 将千分位逗号误读为点：7.084.00 -> 7084.00。
    """
    if not value:
        return None
    clean = re.sub(r'[^0-9.,\-]', '', value.replace('－', '-').replace('，', ','))
    if clean.count('-') > 1 or clean.startswith('-') is False and '-' in clean:
        return None
    sign = '-' if clean.startswith('-') else ''
    digits = clean[1:] if sign else clean
    decimal_match = re.search(r'([.,])(\d{1,2})$', digits)
    if decimal_match:
        integer_part = re.sub(r'[.,]', '', digits[:decimal_match.start()])
        clean = sign + integer_part + '.' + decimal_match.group(2)
    else:
        clean = sign + re.sub(r'[.,]', '', digits)
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


def foreign_amounts_after(text: str, labels: tuple[str, ...], currency: str) -> list[Decimal]:
    """取银行 PDF 字段右侧指定币种金额，兼容币种在金额前或同一行。"""
    normalized = compact(text).upper()
    values: list[Decimal] = []
    for label in labels:
        pattern = (re.escape(compact(label)) + r'.{0,60}?' + re.escape(currency)
                   + r'\s*(' + MONEY_TOKEN + r')')
        for match in re.finditer(pattern, normalized, re.I):
            amount = parse_amount(match.group(1))
            if amount is not None:
                values.append(amount)
    return list(dict.fromkeys(values))


# “总费用金额”是该类付款申请实际需核对的总额；不再读取单项“费用金额”。
SOURCE_TABLE_AMOUNT_LABELS = ('报销金额', '总费用金额', '借款金额', '申请金额')


def source_table_amounts(text: str) -> list[Decimal]:
    """取四类来源表格总金额，适配 PDF 表格文字层与 OCR。

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

    return list(dict.fromkeys(values))


def pdf_table_amounts_by_position(path: Path) -> list[Decimal]:
    """按 PDF 坐标取金额表头正下方同列的最近金额，避免扫到日期等其他数字。"""
    values: list[Decimal] = []
    doc = fitz.open(path)
    try:
        for page in doc:
            words = page.get_text('words')  # x0, y0, x1, y1, word, block, line, word_no
            labels = [word for word in words if any(label in compact(word[4])
                      for label in SOURCE_TABLE_AMOUNT_LABELS)]
            for label in labels:
                lx0, ly0, lx1, ly1 = label[:4]
                candidates: list[tuple[float, Decimal]] = []
                for word in words:
                    wx0, wy0, wx1, _wy1, value = word[:5]
                    # 必须在表头下方，且与表头横向重叠或非常接近同一列。
                    if wy0 < ly1 - 2 or wx1 < lx0 - 12 or wx0 > lx1 + 12:
                        continue
                    amount = parse_amount(value)
                    if amount is not None and ('.' in value or ',' in value):
                        candidates.append((wy0 - ly1, amount))
                if candidates:
                    values.append(min(candidates, key=lambda item: item[0])[1])
    finally:
        doc.close()
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


def citic_foreign_amounts(text: str) -> list[tuple[Decimal, str]]:
    """中信银行外币水单：购汇金额 + 现汇金额（若有）作为同币种总金额。"""
    found: list[tuple[Decimal, str]] = []
    for currency in FOREIGN_CURRENCIES:
        purchase = foreign_amounts_after(text, ('购汇金额',), currency)
        spot = foreign_amounts_after(text, ('现汇金额',), currency)
        # 每个字段若 OCR 重复识别，只用第一笔。部分中信回单的同一笔
        # “购汇金额”会同时被 OCR 误关联到“现汇金额”；两者完全相等时
        # 必须只取一次，不能误算为双倍。只有两笔金额确实不同才相加。
        if purchase:
            total = purchase[0] if not spot or purchase[0] == spot[0] else purchase[0] + spot[0]
            found.append((total.quantize(Decimal('0.01')), f'中信银行-{currency}'))
        elif spot:
            found.append((spot[0], f'中信银行-{currency}'))
    return list(dict.fromkeys(found))


def cmb_foreign_amounts(text: str) -> list[tuple[Decimal, str]]:
    """招商银行外币回单：读取“买入币种/金额”右侧的外币金额。"""
    found: list[tuple[Decimal, str]] = []
    for currency in FOREIGN_CURRENCIES:
        values = foreign_amounts_after(text, ('买入币种/金额', '买入币种金额'), currency)
        if values:
            found.append((values[0], f'招商银行-{currency}'))
    return list(dict.fromkeys(found))


def spdb_foreign_amount_from_filename(path: Path) -> tuple[Decimal, str] | None:
    """浦发回单按约定从文件名取币种及金额，如“回单-富皇USD28577.pdf”。"""
    match = re.search(r'(USD|CAD|GBP|EUR)\s*[-_－]?\s*(' + MONEY_TOKEN + r')', path.stem, re.I)
    if not match:
        return None
    amount = parse_amount(match.group(2))
    if amount is None:
        return None
    return amount, f'浦发银行-{match.group(1).upper()}'


FOREIGN_BANKS = ('中信银行', '招商银行', '浦发银行')


def foreign_amount_label(text: str, bank_name: str) -> list[tuple[Decimal, str]]:
    """深层外币回单按银行格式读取 USD/CAD/GBP/EUR 金额。"""
    found: list[tuple[Decimal, str]] = []
    # 英文字段是三家银行 PDF 文字层中最稳定的部分；中文/OCR 结果作为兼容。
    labels = ('购汇金额', '购 汇 金 额', 'Amount of Purchase', 'AmountofPurchase')
    for currency in FOREIGN_CURRENCIES:
        values = foreign_amounts_after(text, labels, currency)
        if values:
            found.append((values[0], f'{bank_name}-{currency}'))
    return list(dict.fromkeys(found))


def matches_bank_receipt(text: str, bank_name: str) -> bool:
    normalized = compact(text).upper()
    if bank_name == '中信银行':
        return '中信银行' in text or 'CHINACITICBANK' in normalized
    if bank_name == '招商银行':
        return ('致：招商银行' in text or '致:招商银行' in text
                or '招商银行' in text or 'CHINAMERCHANTSBANK' in normalized)
    return ('上海浦东发展银行网上银行电子回单-借记回单' in text
            or '上海浦东发展银行' in text
            or 'SHANGHAIPUDONGDEVELOPMENTBANK' in normalized)


def customer_name_from_receipt_path(bank_root: Path, receipt_path: Path) -> str:
    """客户名取银行目录下一层，兼容“1-利盟”这类序号前缀。"""
    try:
        first_part = receipt_path.relative_to(bank_root).parts[0]
    except (ValueError, IndexError):
        first_part = '客户'
    name = re.sub(r'^\d+\s*[-_－、. ]*', '', first_part).strip()
    return name or '客户'


def amount_name(amount: Decimal) -> str:
    return format(amount.normalize(), 'f').rstrip('0').rstrip('.') or '0'


def payment_foreign_amount(path: Path) -> Decimal | None:
    """服务商付款申请只按正文付款总额读取。"""
    values = amounts_after(read_text(path), ('付款总额',))
    if not values:
        values = amounts_after(read_text(path, retry_enhanced=True), ('付款总额',))
    return values[0] if values else None


def organize_foreign_receipts(foreign_root: Path) -> tuple[int, int]:
    """整理三家银行深层回单及同级服务商付款申请，返回归档对数、冲突数。"""
    service_root = foreign_root.parent / '服务商'
    if foreign_root.name != '外币' or not service_root.is_dir():
        return 0, 0
    log('外币预整理开始：扫描三家银行的深层回单及同级 服务商 付款申请。')
    payments: list[Record] = []
    service_files = sorted(path for path in service_root.rglob('*.pdf') if path.is_file())
    log(f'外币预整理：发现服务商付款申请 PDF {len(service_files)} 份，开始读取付款总额。')
    for index, path in enumerate(service_files, 1):
        try:
            amount = payment_foreign_amount(path)
            if amount is not None:
                payments.append(Record(path, amount, '服务商付款申请'))
            else:
                log(f'[服务商 {index}/{len(service_files)}] 未识别付款总额：{path.name}')
        except Exception as exc:
            log(f'[服务商 {index}/{len(service_files)}] 读取失败：{path.name}，{exc}')

    receipts: list[Record] = []
    receipt_customers: dict[Path, str] = {}
    for bank_name in FOREIGN_BANKS:
        bank_root = foreign_root / bank_name
        if not bank_root.is_dir():
            continue
        # 仅扫尚未整理到日期目录的深层 PDF；已经规范命名的回单不重复处理。
        candidates = [path for path in sorted(bank_root.rglob('*.pdf'))
                      if path.is_file() and '水单正常匹配' not in path.parts
                      and not re.fullmatch(r'\d{6}', path.parent.name)]
        log(f'外币预整理：{bank_name} 深层 PDF {len(candidates)} 份，正在识别回单。')
        for index, path in enumerate(candidates, 1):
            try:
                text = read_text(path)
                if not matches_bank_receipt(text, bank_name):
                    continue
                if bank_name == '浦发银行':
                    value = spdb_foreign_amount_from_filename(path)
                    values = [value] if value else []
                else:
                    values = foreign_amount_label(text, bank_name)
                    if not values:
                        text = read_text(path, retry_enhanced=True)
                        values = foreign_amount_label(text, bank_name)
                if not values:
                    log(f'[{bank_name} {index}/{len(candidates)}] 已识别回单但未读取外币金额：{path.name}')
                    continue
                # 一份有效银行回单只允许一个币种金额参与自动整理。
                if len(values) != 1:
                    log(f'[{bank_name} {index}/{len(candidates)}] 回单出现多个外币金额，保留不移动：{path.name}')
                    continue
                amount, kind = values[0]
                receipts.append(Record(path, amount, kind))
                receipt_customers[path] = customer_name_from_receipt_path(bank_root, path)
                log(f'[{bank_name} {index}/{len(candidates)}] 回单金额：{kind} {amount}')
            except Exception as exc:
                log(f'[{bank_name} {index}/{len(candidates)}] 读取失败：{path.name}，{exc}')

    receipt_by_amount: dict[Decimal, list[Record]] = defaultdict(list)
    payment_by_amount: dict[Decimal, list[Record]] = defaultdict(list)
    for record in receipts:
        receipt_by_amount[record.amount].append(record)
    for record in payments:
        payment_by_amount[record.amount].append(record)
    today_folder = date.today().strftime('%y%m%d')
    moved = 0
    conflicts = 0
    for amount in sorted(set(receipt_by_amount) & set(payment_by_amount)):
        left = receipt_by_amount[amount]
        right = payment_by_amount[amount]
        if len(left) != 1 or len(right) != 1:
            conflicts += len(left) + len(right)
            log(f'外币预整理金额 {amount} 存在 {len(left)} 份回单、{len(right)} 份付款申请，未移动。')
            continue
        receipt, payment = left[0], right[0]
        bank_name, currency = receipt.kind.split('-', 1)
        customer = receipt_customers[receipt.path]
        approval = approval_number_from_filename(payment.path) or '审批编号待核对'
        destination = foreign_root / bank_name / today_folder
        receipt_name = f'回单-{customer}-{currency}-{amount_name(amount)}.pdf'
        payment_name = f'{customer}-{currency}-{amount_name(amount)}-{approval}.pdf'
        receipt_target = safe_target(destination, receipt_name)
        payment_target = safe_target(destination, payment_name)
        destination.mkdir(parents=True, exist_ok=True)
        try:
            shutil.move(str(receipt.path), str(receipt_target))
            try:
                shutil.move(str(payment.path), str(payment_target))
            except Exception:
                shutil.move(str(receipt_target), str(receipt.path))
                raise
            moved += 1
            log(f'外币预整理已归档：{receipt_target.name} <-> {payment_target.name}')
        except Exception as exc:
            log(f'外币预整理移动失败：{receipt.path.name}，{exc}')
    log(f'外币预整理完成：已集中归档 {moved} 对；同金额多候选未移动 {conflicts} 个文件。')
    return moved, conflicts


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
        # 对“总费用金额”“汇款金额”这类细字加表格线，通常比原彩色截图更清晰。
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


def source_records(pdf_dir: Path, receipt_dir: Path) -> tuple[list[Record], list[Record], list[dict[str, str]], list[Path]]:
    cny: list[Record] = []
    foreign: list[Record] = []
    report: list[dict[str, str]] = []
    # 人民币及普通付款申请只扫描用户选择的当前文件夹，不进入任何子文件夹。
    direct_files = [path for path in sorted(pdf_dir.iterdir())
                    if path.is_file() and path.suffix.lower() in SOURCE_SUFFIXES]
    # 中信、招商、浦发外币是固定例外：在“银行名称/日期文件夹”内，含“回单”的
    # PDF 是水单；同级其他 PDF（例如付款审核）是来源付款申请。
    bank_review_files: list[Path] = []
    bank_counts: dict[str, int] = {}
    for bank_name in ('中信银行', '招商银行', '浦发银行'):
        bank_root = receipt_dir / bank_name
        review_files = ([path for date_dir in sorted(bank_root.iterdir()) if date_dir.is_dir()
                         for path in sorted(date_dir.iterdir())
                         if path.is_file() and path.suffix.lower() == '.pdf'
                         and '回单' not in path.stem] if bank_root.is_dir() else [])
        bank_review_files.extend(review_files)
        bank_counts[bank_name] = len(review_files)
    files = direct_files + bank_review_files
    bank_review_set = set(bank_review_files)
    log('第一步完成扫描：当前目录来源 PDF ' + str(len(direct_files))
        + ' 份（不扫描子文件夹）；中信银行日期子文件夹付款审核 PDF '
        + str(bank_counts['中信银行']) + ' 份；招商银行日期子文件夹付款审核 PDF '
        + str(bank_counts['招商银行']) + ' 份；浦发银行日期子文件夹付款审核 PDF '
        + str(bank_counts['浦发银行']) + ' 份，开始识别金额。')
    for index, path in enumerate(files, 1):
        try:
            log(f'[PDF {index}/{len(files)}] 正在识别：{path.name}')
            text = read_text(path)
            is_foreign_pdf = path in bank_review_set or bool(re.search(r'\b(?:USD|CAD|GBP|EUR)\b', text, re.I))
            # 付款总额/汇款金额可取右侧值；四类表格总金额只按表头位置取值。
            table_amount_values = (pdf_table_amounts_by_position(path)
                                   + source_table_amounts(text)) if not is_foreign_pdf else []
            table_amount_values = list(dict.fromkeys(table_amount_values))
            cny_values = (amounts_after(text, ('付款总额', '汇款金额'))
                          + table_amount_values) if not is_foreign_pdf else []
            cny_values = list(dict.fromkeys(cny_values))
            # 外币付款申请常只有“付款总额”数值，不一定印出 USD/CAD 等币种；
            # 单独保存该字段，后续只与中信银行外币水单进行核对。
            foreign_values = amounts_after(text, ('付款总额',))
            if not cny_values and not foreign_values:
                # 扫描版付款申请中的细字表格先原图 OCR，失败后才走局部对比度增强。
                log('  原图未识别到来源金额，正在增强后重试…')
                text = read_text(path, retry_enhanced=True)
                is_foreign_pdf = path in bank_review_set or bool(re.search(r'\b(?:USD|CAD|GBP|EUR)\b', text, re.I))
                table_amount_values = source_table_amounts(text) if not is_foreign_pdf else []
                cny_values = (amounts_after(text, ('付款总额', '汇款金额'))
                              + table_amount_values) if not is_foreign_pdf else []
                cny_values = list(dict.fromkeys(cny_values))
                foreign_values = amounts_after(text, ('付款总额',))
            cny.extend(Record(path, value, '人民币') for value in cny_values)
            foreign.extend(Record(path, value, '外币付款申请') for value in foreign_values)
            if not cny_values and not foreign_values:
                log('  未识别到待匹配金额。')
                report.append(row(path, '', '', '未识别到付款总额、报销金额或汇款金额'))
            else:
                values = cny_values or foreign_values
                log('  识别金额：' + '、'.join(str(value) for value in values))
        except Exception as exc:
            log(f'  识别失败：{exc}')
            report.append(row(path, '', '', f'源 PDF 读取失败：{exc}'))
    log(f'来源 PDF 识别完成：人民币候选 {len(cny)} 条，外币付款候选 {len(foreign)} 条。')
    return cny, foreign, report, files


def receipt_records(receipt_dir: Path) -> tuple[list[Record], list[Record], list[dict[str, str]], list[Path]]:
    cny: list[Record] = []
    foreign: list[Record] = []
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

    # 人民币水单只扫描当前目录；中信、招商、浦发外币例外，按“银行/日期文件夹/PDF”读取。
    files = [path for path in sorted(receipt_dir.iterdir()) if is_receipt_file(path)]
    bank_receipt_files: dict[str, list[Path]] = {}
    for bank_name in ('中信银行', '招商银行', '浦发银行'):
        bank_root = receipt_dir / bank_name
        bank_receipt_files[bank_name] = ([path for date_dir in sorted(bank_root.iterdir()) if date_dir.is_dir()
                                         for path in sorted(date_dir.iterdir())
                                         if path.is_file() and path.suffix.lower() == '.pdf'
                                         and '回单' in path.stem] if bank_root.is_dir() else [])
    citic_files = bank_receipt_files['中信银行']
    cmb_files = bank_receipt_files['招商银行']
    spdb_files = bank_receipt_files['浦发银行']
    all_files = files + citic_files + cmb_files + spdb_files
    log(f'第二步完成扫描：当前目录水单 {len(files)} 份；中信银行日期子文件夹 PDF {len(citic_files)} 份；招商银行日期子文件夹 PDF {len(cmb_files)} 份；浦发银行日期子文件夹 PDF {len(spdb_files)} 份。')
    for index, path in enumerate(all_files, 1):
        try:
            label = ('中信外币' if path in citic_files else ('招商外币' if path in cmb_files
                     else ('浦发外币' if path in spdb_files else '水单')))
            log(f'[{label} {index}/{len(all_files)}] 正在识别：{path.name}')
            if path in citic_files:
                # 文件名只用来筛选“回单”；金额必须从回单 PDF 正文的
                # “购汇金额”及“现汇金额”字段获取，不能把文件名金额作为依据。
                text = read_text(path)
                citic_values = citic_foreign_amounts(text)
                if not citic_values:
                    log('  原图未识别到中信银行外币金额，正在增强后重试…')
                    text = read_text(path, retry_enhanced=True)
                    citic_values = citic_foreign_amounts(text)
                for amount, kind in citic_values:
                    foreign.append(Record(path, amount, kind))
                if citic_values:
                    log('  识别金额：' + '、'.join(f'{kind} {amount}' for amount, kind in citic_values))
                else:
                    log('  未识别到中信银行购汇金额/现汇金额的 USD、CAD、GBP 或 EUR。')
                continue
            if path in cmb_files:
                # 文件名只用来筛选“回单”；金额必须从招商回单 PDF 正文的
                # “买入币种/金额”字段读取，不能从文件名取得。
                text = read_text(path)
                cmb_values = cmb_foreign_amounts(text)
                if not cmb_values:
                    log('  原图未识别到招商银行外币金额，正在增强后重试…')
                    text = read_text(path, retry_enhanced=True)
                    cmb_values = cmb_foreign_amounts(text)
                for amount, kind in cmb_values:
                    foreign.append(Record(path, amount, kind))
                if cmb_values:
                    log('  识别金额：' + '、'.join(f'{kind} {amount}' for amount, kind in cmb_values))
                else:
                    log('  未识别到招商银行买入币种/金额的 USD、CAD、GBP 或 EUR。')
                continue
            if path in spdb_files:
                # 浦发按业务约定：文件名只要含“回单”，其金额从回单文件名的
                # 币种紧随金额处读取，例如“回单-富皇USD28577.pdf”。
                spdb_value = spdb_foreign_amount_from_filename(path)
                if spdb_value:
                    amount, kind = spdb_value
                    foreign.append(Record(path, amount, kind))
                    log(f'  按回单文件名识别金额：{kind} {amount}')
                else:
                    log('  未识别到浦发回单文件名中的 USD、CAD、GBP 或 EUR 金额。')
                continue
            text = read_text(path)
            cny_values = cny_receipt_amounts(text)
            for amount, kind in cny_values:
                cny.append(Record(path, amount, kind))
            file_name = path.name.lower()
            is_citic = '中信银行' in path.name or 'citic' in file_name
            is_cmb = '招商银行' in path.name or 'cmb' in file_name
            usd_values = usd_receipt_amounts(text, is_citic, is_cmb)
            for amount, kind in usd_values:
                foreign.append(Record(path, amount, kind))
            values = cny_values + usd_values
            if not values and path.suffix.lower() in IMAGE_SUFFIXES:
                log('  原图未识别到水单金额，正在轻量增强后重试…')
                text = ocr_text(path, retry_enhanced=True)
                cny_values = cny_receipt_amounts(text)
                usd_values = usd_receipt_amounts(text, is_citic, is_cmb)
                for amount, kind in cny_values:
                    cny.append(Record(path, amount, kind))
                for amount, kind in usd_values:
                    foreign.append(Record(path, amount, kind))
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
    log(f'水单识别完成：人民币候选 {len(cny)} 条，外币候选 {len(foreign)} 条。')
    return cny, foreign, report, all_files


def row(source: Path | str, receipt: Path | str, amount: Decimal | str, result: str) -> dict[str, str]:
    return {'来源PDF': str(source), '水单': str(receipt), '金额': str(amount), '处理结果': result}


def unique_pairs(sources: list[Record], receipts: list[Record], report: list[dict[str, str]]) -> tuple[list[tuple[Record, Record]], set[Record]]:
    by_amount_sources: dict[Decimal, list[Record]] = defaultdict(list)
    by_amount_receipts: dict[Decimal, list[Record]] = defaultdict(list)
    for item in sources:
        by_amount_sources[item.amount].append(item)
    for item in receipts:
        by_amount_receipts[item.amount].append(item)
    pairs = []
    conflict_records: set[Record] = set()
    for amount in sorted(set(by_amount_sources) & set(by_amount_receipts)):
        left = list({item.path: item for item in by_amount_sources[amount]}.values())
        right = list({item.path: item for item in by_amount_receipts[amount]}.values())
        if len(left) == len(right) == 1:
            log(f'金额 {amount} 唯一匹配：{left[0].path.name} <-> {right[0].path.name}')
            pairs.append((left[0], right[0]))
        else:
            log(f'金额 {amount} 有 {len(left)} 份 PDF、{len(right)} 份水单候选，保留供人工核对。')
            report.append(row('; '.join(str(x.path) for x in left), '; '.join(str(x.path) for x in right), amount, '同金额存在多份候选，未自动移动'))
            # 只有两边都存在、却无法唯一对应的同金额冲突项才进入异常文件夹。
            conflict_records.update(left)
            conflict_records.update(right)
    return pairs, conflict_records


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


def approval_number_from_filename(path: Path) -> str | None:
    """取得付款申请文件名中以 2026 开头的审批编号。"""
    match = re.search(r'2026\d+', path.stem)
    return match.group(0) if match else None


def archive_pair(source: Record, receipt: Record, output_dir: Path) -> tuple[Path, Path]:
    output_dir.mkdir(exist_ok=True)
    source_target = safe_target(output_dir, source.path.name)
    # 正常匹配保留付款申请和水单/回单的原文件名。回单仅在末尾追加
    # 付款申请文件名中的 2026 审批编号，既能关联又不会把回单改成付款申请名。
    approval_number = approval_number_from_filename(source.path)
    receipt_stem = receipt.path.stem
    if approval_number and approval_number not in receipt_stem:
        receipt_stem += '-' + approval_number
    receipt_name = receipt_stem + receipt.path.suffix.lower()
    receipt_target = safe_target(output_dir, receipt_name)
    shutil.move(str(source.path), str(source_target))
    try:
        shutil.move(str(receipt.path), str(receipt_target))
    except Exception:
        # 归档必须成对：水单移动失败时将 PDF 退回原位。
        shutil.move(str(source_target), str(source.path))
        raise
    return source_target, receipt_target


def pair_output_dir(source: Record, receipt: Record, default_dir: Path) -> Path:
    """银行日期目录内的同级付款审核和回单，归档在该日期目录中。"""
    if (source.kind == '外币付款申请' and any(bank in receipt.kind
            for bank in ('中信银行-', '招商银行-', '浦发银行-'))
            and source.path.parent == receipt.path.parent):
        return source.path.parent / '水单正常匹配'
    return default_dir / '水单正常匹配'


def archive_exception(path: Path, output_dir: Path, preferred_name: str | None = None) -> Path:
    """将同金额冲突文件移至异常目录，保留或指定安全文件名。"""
    output_dir.mkdir(exist_ok=True)
    target = safe_target(output_dir, preferred_name or path.name)
    shutil.move(str(path), str(target))
    return target


def foreign_conflict_name(record: Record, index: int) -> str:
    """外币同金额冲突项：保留原名主体，只在扩展名前追加序号。"""
    prefix = record.path.stem or '外币水单'
    return f'{prefix}.{index}{record.path.suffix.lower()}'


def write_report(folder: Path, entries: list[dict[str, str]]) -> Path:
    report_path = folder / '水单匹配处理报告.csv'
    with report_path.open('w', encoding='utf-8-sig', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=('来源PDF', '水单', '金额', '处理结果'))
        writer.writeheader()
        writer.writerows(entries)
    return report_path


def choose_spreadsheet(root: Tk) -> Path | None:
    selected = filedialog.askopenfilename(
        parent=root,
        title='第二步：选择人民币金额核对表格',
        filetypes=(('Excel 表格', '*.xlsx *.xlsm *.xls'),),
    )
    return Path(selected) if selected else None


def is_rmb_payment_pdf(path: Path) -> bool:
    """第二步仅核对人民币：外币目录中的付款审核 PDF 不参与表格金额匹配。"""
    return '外币' not in path.parts


def spreadsheet_approval_number(path: Path) -> str | None:
    """表格核对使用付款申请文件名中以 20260 开头的编号。"""
    match = re.search(r'20260\d+', path.stem)
    return match.group(0) if match else None


def payment_amount_candidates(path: Path) -> list[Decimal]:
    """第二步只读取付款总额或报销金额，不采用其他金额字段。"""
    text = read_text(path)
    values = amounts_after(text, ('付款总额', '报销金额'))
    # 文字层中“报销金额”可能位于表头、金额在其下方；仅直接读取不到
    # 付款总额/报销金额时才按表格坐标兜底，避免重复打开 PDF。
    if not values and '报销金额' in compact(text):
        values.extend(pdf_table_amounts_by_position(path))
        values.extend(source_table_amounts(text))
    return list(dict.fromkeys(values))


def payment_handling_date(path: Path) -> str:
    """从付款申请底部“杨舒媛 + 已付/已办理”记录提取月日。"""
    text = read_text(path)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # 自底向上定位，以适配文字层把同一行拆成多行的情况。
    for index in range(len(lines) - 1, -1, -1):
        nearby = ' '.join(lines[max(0, index - 2):min(len(lines), index + 4)])
        if '杨舒媛' not in nearby or not re.search(r'已付|已办理', nearby):
            continue
        match = re.search(r'(20\d{2})[年./-]\s*(\d{1,2})[月./-]\s*(\d{1,2})', nearby)
        if match:
            return f'{int(match.group(2))}月{int(match.group(3))}号'
    return '日期待核对'


def excel_cell_amount(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    return parse_amount(str(value))


def excel_document_number(value: object) -> str:
    return re.sub(r'\.0$', '', str(value or '').strip())


def checked_spreadsheet_path(source: Path) -> Path:
    return safe_target(source.parent, source.stem + '-已核对' + source.suffix.lower())


def reconcile_rmb_spreadsheet(spreadsheet: Path, payment_pdfs: list[Path]) -> tuple[Path, int, int]:
    """将本轮人民币付款申请同表格单据号/支出逐笔核对，并另存结果表。"""
    import win32com.client

    output = checked_spreadsheet_path(spreadsheet)
    shutil.copy2(spreadsheet, output)
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    checked = 0
    abnormal = 0
    try:
        workbook = excel.Workbooks.Open(str(output.resolve()))
        sheets: list[tuple[object, int, int, int, int]] = []
        for sheet in workbook.Worksheets:
            used = sheet.UsedRange
            first_row, first_col = int(used.Row), int(used.Column)
            last_row = first_row + int(used.Rows.Count) - 1
            last_col = first_col + int(used.Columns.Count) - 1
            for row_number in range(first_row, min(last_row, first_row + 20) + 1):
                headers = {compact(str(sheet.Cells(row_number, col).Value or '')): col
                           for col in range(first_col, last_col + 1)}
                if '单据号' not in headers or '支出' not in headers:
                    continue
                note_col = headers.get('核对记录备注')
                if not note_col:
                    note_col = last_col + 1
                    sheet.Cells(row_number, note_col).Value = '核对记录备注'
                sheets.append((sheet, row_number, headers['单据号'], headers['支出'], note_col))
                break
        if not sheets:
            raise ValueError('表格中未找到同一表头行的“单据号”和“支出”列。')

        for pdf_path in payment_pdfs:
            approval = spreadsheet_approval_number(pdf_path)
            if not approval:
                log(f'[表格核对] 跳过：{pdf_path.name}，文件名未找到以 20260 开头的审批编号。')
                continue
            amounts = payment_amount_candidates(pdf_path)
            if not amounts:
                log(f'[表格核对] 跳过：{pdf_path.name}，未识别到付款总额或报销金额。')
                continue
            matched_rows: list[tuple[object, int, int, Decimal]] = []
            for sheet, header_row, document_col, expense_col, note_col in sheets:
                last_row = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1
                for row_number in range(header_row + 1, last_row + 1):
                    if excel_document_number(sheet.Cells(row_number, document_col).Value) != approval:
                        continue
                    matched_rows.append((sheet, row_number, note_col,
                                         excel_cell_amount(sheet.Cells(row_number, expense_col).Value) or Decimal('0.00')))
            if not matched_rows:
                log(f'[表格核对] 单据号 {approval} 未在表格中找到，未写备注。')
                continue
            total = sum((item[3] for item in matched_rows), Decimal('0.00')).quantize(Decimal('0.01'))
            matched_amount = next((amount for amount in amounts if amount == total), None)
            if matched_amount is None:
                for sheet, row_number, note_col, _amount in matched_rows:
                    sheet.Cells(row_number, note_col).Value = '金额异常'
                abnormal += len(matched_rows)
                log(f'[表格核对] 单据号 {approval}：PDF 金额与支出合计 {total} 不一致，已写金额异常。')
                continue
            date_text = payment_handling_date(pdf_path)
            if len(matched_rows) == 1:
                note = f'已核对---{date_text}已打印付款申请及回单'
            else:
                note = f'已核对---{date_text}已打印付款申请及回单，总计{total:.2f}元'
            for sheet, row_number, note_col, _amount in matched_rows:
                sheet.Cells(row_number, note_col).Value = note
            checked += len(matched_rows)
            log(f'[表格核对] 单据号 {approval}：已核对 {len(matched_rows)} 行，金额 {matched_amount}。')
        workbook.Save()
        return output, checked, abnormal
    except Exception:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
            workbook = None
        output.unlink(missing_ok=True)
        raise
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def choose_folder(root: Tk, title: str) -> Path | None:
    selected = filedialog.askdirectory(parent=root, title=title, mustexist=True)
    return Path(selected) if selected else None


def foreign_batch_roots(parent: Path) -> list[Path]:
    """在共同上级下找每个可独立处理的“外币 + 服务商”批次。"""
    candidates = [parent] + [path for path in parent.rglob('外币') if path.is_dir()]
    roots = [path for path in candidates
             if path.name == '外币' and (path.parent / '服务商').is_dir()]
    return sorted(set(roots))


def run_foreign_batch_queue(root: Tk) -> None:
    """一次选共同上级，按队列执行多批外币预整理。"""
    while True:
        parent = choose_folder(root, '批量外币整理：选择所有批次共同的上级文件夹')
        if not parent:
            return
        queue = foreign_batch_roots(parent)
        if not queue:
            messagebox.showinfo('未发现外币批次',
                                '未找到同时具备“外币”与同级“服务商”文件夹的批次。',
                                parent=root)
        else:
            log(f'批量外币整理：发现 {len(queue)} 个批次，开始排队处理。')
            total_moved = 0
            total_conflicts = 0
            for index, foreign_root in enumerate(queue, 1):
                log(f'批量外币 [{index}/{len(queue)}]：{display_path(foreign_root)}')
                moved, conflicts = organize_foreign_receipts(foreign_root)
                total_moved += moved
                total_conflicts += conflicts
            log(f'批量外币整理完成：处理 {len(queue)} 个批次，归档 {total_moved} 对；多候选未移动 {total_conflicts} 个文件。')
            messagebox.showinfo('批量外币整理完成',
                                f'已处理 {len(queue)} 个批次。\n'
                                f'已集中归档 {total_moved} 对文件。\n'
                                f'同金额多候选未移动 {total_conflicts} 个文件。',
                                parent=root)
        if not messagebox.askyesno('继续批量整理', '是否继续选择另一组共同上级文件夹？', parent=root):
            return


def main() -> None:
    enable_console_color()
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        batch_mode = messagebox.askyesnocancel(
            '选择处理模式',
            '是否批量整理多个外币文件夹？\n\n'
            '选择“是”：只需选择共同上级，自动排队处理每个“外币 + 服务商”批次。\n'
            '选择“否”：进入原有的单批 PDF / 水单核对流程。\n'
            '选择“取消”：退出。',
            parent=root,
        )
        if batch_mode is None:
            return
        if batch_mode:
            run_foreign_batch_queue(root)
            return
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

            organize_foreign_receipts(receipt_dir)
            cny_sources, foreign_sources, report, _source_files = source_records(pdf_dir, receipt_dir)
            cny_receipts, foreign_receipts, receipt_report, _receipt_files = receipt_records(receipt_dir)
            report.extend(receipt_report)
            cny_pairs, cny_conflicts = unique_pairs(cny_sources, cny_receipts, report)
            foreign_pairs, foreign_conflicts = unique_pairs(foreign_sources, foreign_receipts, report)

            moved = 0
            archived_payment_pdfs: list[Path] = []
            for source, receipt in cny_pairs:
                archived_source, _archived_receipt = archive_pair(source, receipt, pair_output_dir(source, receipt, pdf_dir))
                if is_rmb_payment_pdf(archived_source):
                    archived_payment_pdfs.append(archived_source)
                log(f'已归档到 水单正常匹配：{source.path.name}')
                report.append(row(source.path.name, receipt.path.name, source.amount, f'已归档：水单正常匹配（{receipt.kind}）'))
                moved += 1
            moved_source_paths = {source.path for source, _receipt in cny_pairs}
            for source, receipt in foreign_pairs:
                if source.path in moved_source_paths:
                    # 同一来源 PDF 已按人民币正常归档，不允许重复移动。
                    continue
                archived_source, _archived_receipt = archive_pair(source, receipt, pair_output_dir(source, receipt, pdf_dir))
                if is_rmb_payment_pdf(archived_source):
                    archived_payment_pdfs.append(archived_source)
                log(f'已归档到 水单正常匹配：{source.path.name}')
                report.append(row(source.path.name, receipt.path.name, source.amount, f'已归档：水单正常匹配（{receipt.kind}）'))
                moved += 1

            # 仅“同金额两边都有多份候选”的冲突项进入异常文件夹。
            # 只有 PDF 或只有水单、未识别金额等文件保持原目录不动。
            cny_exception_paths = {record.path for record in cny_conflicts}
            foreign_exception_records = [record for record in foreign_conflicts if record.path.exists()]
            exception_files = [path for path in cny_exception_paths if path.exists()]
            exception_count = 0
            for path in dict.fromkeys(exception_files):
                archive_exception(path, pdf_dir / '异常文件夹')
                log(f'已移至 异常文件夹：{path.name}')
                report.append(row(path.name, '', '', '同金额多份候选，已移至异常文件夹待进一步验证'))
                exception_count += 1
            # 外币同金额冲突保留原始名称，并按“原名.序号”区分，
            # 例如：回单-富皇-USD-6666.1.pdf、回单-富皇-USD-6666.2.pdf。
            foreign_by_path = {record.path: record for record in foreign_exception_records}
            for index, record in enumerate(foreign_by_path.values(), 1):
                if record.path in moved_source_paths:
                    continue
                target = archive_exception(record.path, pdf_dir / '异常', foreign_conflict_name(record, index))
                log(f'已移至 异常：{target.name}')
                report.append(row(record.path.name, target.name, record.amount, '外币同金额多候选，已重命名并移至异常待进一步验证'))
                exception_count += 1

            report_path = write_report(pdf_dir, report)
            log(f'处理完成：已归档 {moved} 对文件。')
            log(f'待进一步验证：已移至异常文件夹 {exception_count} 个文件。')
            log('处理报告：' + display_path(report_path))
            if archived_payment_pdfs:
                spreadsheet = choose_spreadsheet(root)
                if spreadsheet:
                    log('开始第二步：人民币付款申请与表格金额核对。')
                    output, checked, abnormal = reconcile_rmb_spreadsheet(spreadsheet, archived_payment_pdfs)
                    log(f'第二步完成：已核对 {checked} 行，金额异常 {abnormal} 行。')
                    log('表格核对结果：' + display_path(output))
                else:
                    log('第二步已跳过：未选择表格。')
            else:
                log('第二步跳过：本批没有已归档的人民币付款申请 PDF。')
            continue_processing = messagebox.askyesno(
                '本批处理完成',
                f'已归档 {moved} 对文件。\n已移至异常文件夹 {exception_count} 个文件。\n处理报告：\n{report_path}\n\n是否继续处理下一批？',
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
