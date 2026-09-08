# -*- coding: utf-8 -*-
"""后台任务队列：单 worker 线程串行处理，避免 PaddleOCR 单例并发问题。

任务流程：create_task 存文件并入队 → worker 逐任务处理 → 前端轮询 get_task → 完成后 download_path 取结果。
"""
import os
import queue
import shutil
import threading
import time
import uuid

import processor
import report_processor
import report_csv_cli

_TMP_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.tmp')

_TASKS = {}          # task_id -> dict（状态/进度/文件名）
_QUEUE = queue.Queue()
_LOCK = threading.Lock()


def _make_task_id():
    return uuid.uuid4().hex[:12]


def create_task(pdf_files, mode, inv_type, layout='v', start_cell='A1',
                template=None, sheet_name=''):
    """创建后台任务。pdf_files: [(原始文件名, bytes), ...]；template: (原始文件名, bytes) 或 None。
    返回 task_id。"""
    task_id = _make_task_id()
    task_dir = os.path.join(_TMP_ROOT, task_id)
    in_dir = os.path.join(task_dir, 'in')
    out_dir = os.path.join(task_dir, 'out')
    os.makedirs(in_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)

    # 保存上传文件：加序号前缀防重名
    saved = []
    for i, (orig_name, data) in enumerate(pdf_files, 1):
        safe = processor.sanitize_filename(orig_name)
        if not safe.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
            safe += '.pdf'
        path = os.path.join(in_dir, '%03d_%s' % (i, safe))
        with open(path, 'wb') as f:
            f.write(data)
        saved.append(path)

    # 保存可选表格模板
    template_path = None
    if template:
        tname = processor.sanitize_filename(template[0])
        if not tname.lower().endswith(('.xlsx', '.xlsm')):
            tname += '.xlsx'
        template_path = os.path.join(in_dir, 'template_' + tname)
        with open(template_path, 'wb') as f:
            f.write(template[1])

    task = {
        'id': task_id,
        'dir': task_dir,
        'out_dir': out_dir,
        'status': 'pending',
        'current': 0,
        'total': 0,
        'message': '等待处理…',
        'filename': '',
        'error': '',
        'logs': [],
        'created': time.time(),
    }
    with _LOCK:
        _TASKS[task_id] = task
    _QUEUE.put((task_id, saved, mode, inv_type, layout, start_cell,
                template_path, sheet_name))
    return task_id


def create_report_task(filename, data, sheet_name, report_profile='no_receivable'):
    """创建报表组 Excel 处理任务。"""
    task_id = _make_task_id()
    task_dir = os.path.join(_TMP_ROOT, task_id)
    in_dir = os.path.join(task_dir, 'in')
    out_dir = os.path.join(task_dir, 'out')
    os.makedirs(in_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)
    safe = processor.sanitize_filename(filename)
    input_path = os.path.join(in_dir, safe)
    with open(input_path, 'wb') as file:
        file.write(data)
    max_step = 9 if report_profile == 'no_salesperson_cost' else 10
    task = {
        'id': task_id, 'dir': task_dir, 'out_dir': out_dir,
        'status': 'pending', 'current': 0, 'total': max_step, 'step': 1, 'max_step': max_step,
        'input_path': input_path, 'sheet_name': sheet_name, 'report_profile': report_profile,
        'source_path': input_path, 'csv_work': os.path.join(out_dir, 'csv_work'),
        'result_path': '', 'web_auto': True,
        'message': '等待处理…', 'filename': '', 'error': '', 'logs': [],
        'created': time.time(), 'elapsed_seconds': 0, 'processing_started_at': None,
    }
    with _LOCK:
        _TASKS[task_id] = task
    _QUEUE.put((task_id, [input_path], '4step', '', 'v', 'A1', None, sheet_name))
    return task_id


def continue_report_task(task_id):
    with _LOCK:
        task = _TASKS.get(task_id)
        if not task or task.get('status') != 'paused':
            return False
        if task.get('step', 1) >= task.get('max_step', 10):
            return False
        task['step'] += 1
        task['status'] = 'pending'
        task['current'] = 0
        task['message'] = '等待继续处理…'
    _QUEUE.put((task_id, [task['input_path']], '4step', '', 'v', 'A1', None, task['sheet_name']))
    return True


def get_task(task_id):
    """返回任务状态的副本；不存在返回 None。"""
    with _LOCK:
        t = _TASKS.get(task_id)
        if not t:
            return None
        result = dict(t)
        elapsed = result.get('elapsed_seconds', 0)
        if result.get('processing_started_at') is not None:
            elapsed += int(time.time() - result['processing_started_at'])
        result['elapsed_seconds'] = elapsed
        return result


def download_path(task_id):
    """处理完成返回结果文件绝对路径，否则返回 None。"""
    task = get_task(task_id)
    if not task or task['status'] not in ('done', 'paused') or not task['filename']:
        return None
    path = task.get('result_path') or os.path.join(task['dir'], 'out', task['filename'])
    return path if os.path.isfile(path) else None


def _worker():
    """单 worker：FIFO 串行处理，进度写回任务表。"""
    while True:
        task_id, pdfs, mode, inv_type, layout, start_cell, template_path, sheet_name = _QUEUE.get()
        task = _TASKS.get(task_id)
        if task is None:
            continue
        task['status'] = 'processing'
        task['processing_started_at'] = time.time()
        task['message'] = '开始处理…'
        task['logs'].append('[开始] 已接收处理任务')

        def progress(cur, tot, msg):
            if mode == '4step':
                task['current'] = task.get('step', cur)
                task['total'] = task.get('max_step', 10)
            else:
                task['current'] = cur
                task['total'] = tot
            task['message'] = msg
            task['logs'].append('[%02d/%02d] %s' % (cur, tot, msg))

        try:
            if mode == '4step':
                result = None; result_name = ''
                first_step = task.get('step', 1)
                for report_step in range(first_step, task.get('max_step', 10) + 1):
                    task['step'] = report_step
                    task['current'] = report_step
                    task['message'] = '正在执行步骤 %d/%d' % (report_step, task.get('max_step', 10))
                    task['logs'].append('[开始] 正在执行步骤 %d/%d' % (report_step, task.get('max_step', 10)))
                    result, result_name = report_csv_cli.run_web_csv_step(task, report_step, progress)
                    task['input_path'] = result
            elif mode == '4':
                extension = os.path.splitext(pdfs[0])[1].lower()
                output = os.path.join(task['out_dir'], '无应收明细处理结果' + extension)
                result = report_processor.process_report(
                    pdfs[0], sheet_name, output, progress)
            elif mode == '1':
                result = processor.process_mode1(
                    pdfs, task['out_dir'], progress,
                    layout=layout, start_cell=start_cell,
                    template_path=template_path, sheet_name=sheet_name)
            elif mode == '2':
                result = processor.process_receipt_mode2(pdfs, task['out_dir'], progress)
            elif mode == '5':
                result = processor.process_shao_meilin(pdfs, task['out_dir'], progress)
            elif mode == '6':
                result = processor.process_payment_receipts(pdfs, task['out_dir'], progress, inv_type)
            elif mode == '9':
                result = processor.process_shao_ordinary_invoice(pdfs, task['out_dir'], progress)
            else:
                result = processor.process_mode2(pdfs, task['out_dir'], inv_type, progress)
            task['filename'] = result_name if mode == '4step' else os.path.basename(result)
            if mode == '4step':
                task['result_path'] = result
            if mode == '4step':
                task['status'] = 'done'
            else:
                task['status'] = 'done'
            task['total'] = task['total'] or 1
            task['current'] = task['total']
            task['message'] = ('步骤 %d 完成，等待确认继续' % task['step']) if mode == '4step' and task['status'] == 'paused' else '处理完成'
            task['logs'].append('[完成] 输出文件已生成：%s' % task['filename'])
        except Exception as e:
            task['status'] = 'error'
            task['error'] = str(e)
            task['message'] = '处理失败：%s' % e
            task['logs'].append('[失败] %s' % e)
        finally:
            if mode == '4step' and task.get('processing_started_at') is not None:
                task['elapsed_seconds'] = task.get('elapsed_seconds', 0) + int(
                    time.time() - task['processing_started_at'])
                task['processing_started_at'] = None


def cleanup_old_tmp(older_than=24 * 3600):
    """清理超过 older_than 秒的旧任务临时目录。"""
    if not os.path.isdir(_TMP_ROOT):
        return
    now = time.time()
    for name in os.listdir(_TMP_ROOT):
        p = os.path.join(_TMP_ROOT, name)
        try:
            if now - os.path.getmtime(p) > older_than:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass


def start_worker():
    """启动单 worker 线程（幂等）。"""
    cleanup_old_tmp()
    t = threading.Thread(target=_worker, daemon=True)
    t.start()
