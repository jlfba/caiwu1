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


SOURCE_SUFFIXES = {'.pdf'}
IMAGE_SUFFIXES = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp'}
MONEY_TOKEN = r'[-－]?s*(?:CNY|USD)?s*[¥￥$]?s*[0-9][0-9,，]*(?:\.[0-9]{1,2})?'


@dataclass(frozen=True)
class Record:
    path: Path
    amount: Decimal
    kind: str


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


def cny_receipt_amounts(text: str) -> list[tuple[Decimal, str]]:
    """返回人民币水单候选；类型顺序与业务规则一致。"""
    normalized = compact(text)
    found: list[tuple[Decimal, str]] = []
    # 类型 3：私账水单，标签最明确，必须优先判断。
    for amount in amounts_after(normalized, ('汇款金额小写',)):
        found.append((amount, '人民币-私账'))
    # 类型 2：APP 支付，金额值本身带负号。
    for match in re.finditer(r'金额[^0-9－-]{0,12}([－-]\s*[¥￥]?\s*[0-9][0-9,]*(?:\.[0-9]{1,2})?)', normalized):
        amount = parse_amount(match.group(1))
        if amount is not None:
            # APP 支付水单通常用负数表示支出，来源 PDF 的付款总额为正数。
            found.append((abs(amount), '人民币-APP'))
    # 类型 1：常规水单，只接受金额字段附近明确出现 CNY 的值。
    for match in re.finditer(r'金额.{0,50}?CNY\s*([¥￥]?\s*[0-9][0-9,]*(?:\.[0-9]{1,2})?)', normalized, re.I):
        amount = parse_amount(match.group(1))
        if amount is not None:
            found.append((amount, '人民币-常规'))
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


def ocr_text(image_path: Path) -> str:
    global _OCR
    if _OCR is None:
        _OCR = get_ocr()
    result, _ = _OCR(str(image_path))
    return '\n'.join(str(row[1]) for row in (result or []) if len(row) > 1)


def pdf_text(path: Path) -> str:
    """优先原生文字层；扫描 PDF 没有文字时逐页渲染后 OCR。"""
    doc = fitz.open(path)
    temporary: list[Path] = []
    try:
        native = '\n'.join(page.get_text('text') for page in doc)
        if compact(native):
            return native
        for number, page in enumerate(doc, 1):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            temp = Path(tempfile.gettempdir()) / f'water-slip-ocr-{os.getpid()}-{number}.png'
            pix.save(str(temp))
            temporary.append(temp)
        return '\n'.join(ocr_text(path) for path in temporary)
    finally:
        doc.close()
        for temp in temporary:
            temp.unlink(missing_ok=True)


def read_text(path: Path) -> str:
    return pdf_text(path) if path.suffix.lower() == '.pdf' else ocr_text(path)


def source_records(pdf_dir: Path) -> tuple[list[Record], list[Record], list[dict[str, str]]]:
    cny: list[Record] = []
    usd: list[Record] = []
    report: list[dict[str, str]] = []
    for path in sorted(pdf_dir.rglob('*')):
        if not path.is_file() or path.suffix.lower() not in SOURCE_SUFFIXES:
            continue
        if {'水单正常匹配', '美金正常'} & set(path.relative_to(pdf_dir).parts):
            continue
        try:
            text = read_text(path)
            relative_parts = path.relative_to(pdf_dir).parts
            is_usd_pdf = ('美金' in relative_parts or '美元' in relative_parts
                          or bool(re.search(r'\bUSD\b', text, re.I)))
            cny_values = amounts_after(text, ('付款总额', '报销金额', '汇款金额')) if not is_usd_pdf else []
            usd_values = amounts_after(text, ('付款总额',)) if is_usd_pdf else []
            cny.extend(Record(path, value, '人民币') for value in cny_values)
            usd.extend(Record(path, value, '美元') for value in usd_values)
            if not cny_values and not usd_values:
                report.append(row(path, '', '', '未识别到付款总额、报销金额或汇款金额'))
        except Exception as exc:
            report.append(row(path, '', '', f'源 PDF 读取失败：{exc}'))
    return cny, usd, report


def receipt_records(receipt_dir: Path) -> tuple[list[Record], list[Record], list[dict[str, str]]]:
    cny: list[Record] = []
    usd: list[Record] = []
    report: list[dict[str, str]] = []
    for path in sorted(receipt_dir.rglob('*')):
        if not path.is_file() or path.suffix.lower() not in IMAGE_SUFFIXES | SOURCE_SUFFIXES:
            continue
        try:
            text = read_text(path)
            for amount, kind in cny_receipt_amounts(text):
                cny.append(Record(path, amount, kind))
            parts = {part.lower() for part in path.parts}
            is_citic = '中信银行' in path.parts or 'citic' in parts
            is_cmb = '招商银行' in path.parts or 'cmb' in parts
            for amount, kind in usd_receipt_amounts(text, is_citic, is_cmb):
                usd.append(Record(path, amount, kind))
        except Exception as exc:
            report.append(row(path, '', '', f'水单读取失败：{exc}'))
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
            pairs.append((left[0], right[0]))
        else:
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
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        pdf_dir = choose_folder(root, '第一步：选择存放报销/付款 PDF 的主文件夹')
        if not pdf_dir:
            return
        receipt_dir = choose_folder(root, '第二步：选择存放人民币/美元等水单的文件夹')
        if not receipt_dir:
            return

        cny_sources, usd_sources, report = source_records(pdf_dir)
        cny_receipts, usd_receipts, receipt_report = receipt_records(receipt_dir)
        report.extend(receipt_report)
        cny_pairs = unique_pairs(cny_sources, cny_receipts, report)
        usd_pairs = unique_pairs(usd_sources, usd_receipts, report)

        moved = 0
        for source, receipt in cny_pairs:
            archive_pair(source, receipt, pdf_dir / '水单正常匹配')
            report.append(row(source.path.name, receipt.path.name, source.amount, f'已归档：水单正常匹配（{receipt.kind}）'))
            moved += 1
        for source, receipt in usd_pairs:
            archive_pair(source, receipt, pdf_dir / '美金正常')
            report.append(row(source.path.name, receipt.path.name, source.amount, f'已归档：美金正常（{receipt.kind}）'))
            moved += 1

        report_path = write_report(pdf_dir, report)
        messagebox.showinfo('水单 PDF 自动核对完成', f'已归档 {moved} 对文件。\n处理报告：\n{report_path}')
    except Exception as exc:
        traceback.print_exc()
        messagebox.showerror('水单 PDF 自动核对失败', str(exc))
    finally:
        root.destroy()


if __name__ == '__main__':
    main()
