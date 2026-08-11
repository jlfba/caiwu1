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

_TMP_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.tmp')

_TASKS = {}          # task_id -> dict（状态/进度/文件名）
_QUEUE = queue.Queue()
_LOCK = threading.Lock()


def _make_task_id():
    return uuid.uuid4().hex[:12]


def create_task(pdf_files, mode, inv_type, layout='v', start_cell='A1'):
    """创建后台任务。pdf_files: [(原始文件名, bytes), ...]。返回 task_id。"""
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
        if not safe.lower().endswith('.pdf'):
            safe += '.pdf'
        path = os.path.join(in_dir, '%03d_%s' % (i, safe))
        with open(path, 'wb') as f:
            f.write(data)
        saved.append(path)

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
        'created': time.time(),
    }
    with _LOCK:
        _TASKS[task_id] = task
    _QUEUE.put((task_id, saved, mode, inv_type, layout, start_cell))
    return task_id


def get_task(task_id):
    """返回任务状态的副本；不存在返回 None。"""
    with _LOCK:
        t = _TASKS.get(task_id)
        return dict(t) if t else None


def download_path(task_id):
    """处理完成返回结果文件绝对路径，否则返回 None。"""
    task = get_task(task_id)
    if not task or task['status'] != 'done' or not task['filename']:
        return None
    path = os.path.join(task['dir'], 'out', task['filename'])
    return path if os.path.isfile(path) else None


def _worker():
    """单 worker：FIFO 串行处理，进度写回任务表。"""
    while True:
        task_id, pdfs, mode, inv_type, layout, start_cell = _QUEUE.get()
        task = _TASKS.get(task_id)
        if task is None:
            continue
        task['status'] = 'processing'
        task['message'] = '开始处理…'

        def progress(cur, tot, msg):
            task['current'] = cur
            task['total'] = tot
            task['message'] = msg

        try:
            if mode == '1':
                result = processor.process_mode1(pdfs, task['out_dir'], progress,
                                                 layout=layout, start_cell=start_cell)
            else:
                result = processor.process_mode2(pdfs, task['out_dir'], inv_type, progress)
            task['filename'] = os.path.basename(result)
            task['status'] = 'done'
            task['total'] = task['total'] or 1
            task['current'] = task['total']
            task['message'] = '处理完成'
        except Exception as e:
            task['status'] = 'error'
            task['error'] = str(e)
            task['message'] = '处理失败：%s' % e


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
