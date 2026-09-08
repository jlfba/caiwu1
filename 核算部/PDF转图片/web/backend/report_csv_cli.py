# -*- coding: utf-8 -*-
"""报表组高速终端版：Excel 只读一次，中间全程 CSV，最后导出 XLSX。"""
import argparse, csv, os, re, sys, zipfile
from collections import Counter
from openpyxl import Workbook, load_workbook

EXCLUDED = ('风驰-数据同步', 'YX订舱', '风驰-卖柜', '李雪原', '龙行-清关', '李增韬')
STEPS = (
    '删除操作状态签入', '筛选应收单价小于等于1', '删除客户简称关键词',
    '删除业务员华南KA', '删除备注J000、无应收、免费补发',
    '删除配仓单号刘丹整柜', '删除整柜且应收金额大于10000',
    '新增无应收分类', '生成透视表', '导出最终CSV',
)
REPORT_PROFILES = {
    'no_receivable': {
        'detail': '无应收明细', 'pivot': '无应收明细透视表', 'zip': '无应收明细-步骤结果.zip',
        'category': '无应收&金额异常', 'unit': '应收单价', 'category_zero': '无应收',
        'required_extra': (),
    },
    'no_salesperson_cost': {
        'detail': '无业务员成本明细', 'pivot': '无业务员成本明细透视表', 'zip': '无业务员成本明细-步骤结果.zip',
        'category': '业务员成本/实际成本', 'unit': '业务员成本单价', 'category_zero': '业务员成本/实际成本',
        'required_extra': ('业务员成本单价',),
    },
}

def text(v): return '' if v is None else str(v).strip()
def number(v):
    try: return float(re.sub(r'[^0-9.\-]', '', text(v).replace(',', '')) or 0)
    except ValueError: return 0.0
def clean_path(v):
    v = v.strip()
    if v.startswith('&'): v = v[1:].strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"": v = v[1:-1]
    return v
def remark_remove(v):
    s = re.sub(r'[\s\u00a0\u3000]+', '', text(v))
    return bool(re.match(r'^J000', s, re.I) or '无应收' in s or '免费补发' in s)
def choose_sheet(path):
    w = load_workbook(path, read_only=True, data_only=True)
    names = w.sheetnames; w.close()
    for i, n in enumerate(names, 1): print(f'{i}. {n}')
    while True:
        v = input('选择工作表序号或名称：').strip()
        if v.isdigit() and 1 <= int(v) <= len(names): return names[int(v)-1]
        if v in names: return v
def read_sheet(path, sheet, csv_path, totals_path, profile='no_receivable'):
    w = load_workbook(path, read_only=True, data_only=True); ws = w[sheet]
    rows = ws.iter_rows(values_only=True); header = [text(x) for x in next(rows)]
    idx = {x:i for i,x in enumerate(header)}; required = ('操作状态','客户简称','业务员','自定义备注','销售产品','应收金额','客户所属机构','运单号', REPORT_PROFILES[profile]['unit'])
    if profile == 'no_receivable': required += ('配仓单号',)
    missing = [x for x in required if x not in idx]
    if missing: w.close(); raise ValueError('缺少列：' + '、'.join(missing))
    totals = Counter()
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        out = csv.writer(f); out.writerow(header)
        for row in rows:
            vals = [text(x) for x in row]; out.writerow(vals)
            org = vals[idx['客户所属机构']]
            if org: totals[org] += 1
    with open(totals_path, 'w', newline='', encoding='utf-8-sig') as f:
        out = csv.writer(f); out.writerow(['客户所属机构','总票数']); out.writerows(sorted(totals.items()))
    w.close(); return header, idx
def process_step(src, dst, log_path, step, header, idx, profile='no_receivable'):
    rules = {
      1: (lambda v: text(v[idx['应收单价']]) != '' and number(v[idx['应收单价']]) <= 1) if profile == 'no_receivable' else (lambda v: text(v[idx['业务员成本单价']]) != '' and number(v[idx['业务员成本单价']]) < 1),
      2: (lambda v: '签入' not in v[idx['操作状态']]) if profile == 'no_salesperson_cost' else (lambda v: text(v[idx['应收单价']]) != '' and number(v[idx['应收单价']]) <= 1),
      3: lambda v: not any(k in text(v[idx['客户简称']]) for k in EXCLUDED),
      4: lambda v: '华南KA' not in text(v[idx['业务员']]),
      5: lambda v: not remark_remove(v[idx['自定义备注']]),
      6: (lambda v: not ('整柜' in text(v[idx['销售产品']]) and number(v[idx['应收金额']]) > 10000)) if profile == 'no_salesperson_cost' else (lambda v: '刘丹整柜' not in text(v[idx['配仓单号']])),
      7: lambda v: not ('整柜' in text(v[idx['销售产品']]) and number(v[idx['应收金额']]) > 10000),
    }
    keep_rule = rules.get(step, lambda v: True)
    with open(src, newline='', encoding='utf-8-sig') as fi, open(dst, 'w', newline='', encoding='utf-8-sig') as fo, open(log_path, 'w', newline='', encoding='utf-8-sig') as fl:
        reader = csv.reader(fi); out = csv.writer(fo); log = csv.writer(fl); h = next(reader); out.writerow(h); log.writerow(h); kept=removed=read=0
        for v in reader:
            read += 1
            if keep_rule(v): out.writerow(v); kept += 1
            else: log.writerow(v); removed += 1
            if read % 100000 == 0: print(f'  已读取 {read} 行，保留 {kept} 行，删除 {removed} 行', flush=True)
    return kept, removed
def add_category(src, dst, idx, profile='no_receivable'):
    with open(src, newline='', encoding='utf-8-sig') as fi, open(dst, 'w', newline='', encoding='utf-8-sig') as fo:
        reader=csv.reader(fi); out=csv.writer(fo); h=next(reader); col = 48 if profile == 'no_receivable' else len(h); h = h + [''] * max(0, col + 1 - len(h)); h[col] = REPORT_PROFILES[profile]['category']; out.writerow(h)
        for v in reader:
            price=number(v[idx[REPORT_PROFILES[profile]['unit']]]); amount=number(v[idx['应收金额']]); v = v + [''] * max(0, col + 1 - len(v)); v[col] = '金额异常' if amount>0 else (REPORT_PROFILES[profile]['category_zero'] if price==0 else ''); out.writerow(v)
def export_xlsx(csv_path, output, log_dir, totals_path, idx, progress=None, profile='no_receivable'):
    config = REPORT_PROFILES[profile]
    wb=Workbook(write_only=True); detail=wb.create_sheet(config['detail'])
    groups={}; totals={}
    if progress: progress(1, 1, '步骤 10/10：正在写入最终无应收明细')
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        r=csv.reader(f); h=next(r); detail.append(h)
        for row_number, v in enumerate(r, 1):
            detail.append(v); org=text(v[idx['客户所属机构']]); cat=text(v[48] if profile == 'no_receivable' else v[-1]); tracking=text(v[idx['运单号']]);
            if org: groups.setdefault(org, Counter()); groups[org][cat] += 1
            if progress and row_number % 100000 == 0:
                progress(1, 1, f'步骤 10/10：最终明细已写入 {row_number} 行')
    with open(totals_path, newline='', encoding='utf-8-sig') as f:
        for i,row in enumerate(csv.reader(f)):
            if i: totals[text(row[0])] = int(number(row[1]))
    cats=[x for x in (config['category_zero'],'金额异常','') if any(g[x] for g in groups.values())]; pivot=wb.create_sheet(config['pivot']); pivot.append(['客户所属机构']+[('空白' if not x else x) for x in cats]+['合计','总票数','占比'])
    for org in sorted(groups):
        nums=[groups[org][x] for x in cats]; total=sum(nums); den=totals.get(org,0); pivot.append([org]+nums+[total,den,total/den if den else 0])
    if progress: progress(1, 1, '步骤 10/10：正在写入透视表和各步骤删除记录')
    # 删除记录保留在处理目录中供核查，不写入最终 XLSX。
    if progress: progress(1, 1, '步骤 10/10：正在压缩保存最终总表')
    wb.save(output); wb.close()

def export_csv_results(csv_path, output_dir, log_dir, totals_path, idx):
    detail_output = os.path.join(output_dir, '无应收明细-最终结果.csv')
    with open(csv_path, 'rb') as source, open(detail_output, 'wb') as target:
        target.write(source.read())
    groups = {}
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f); header = next(reader)
        for row in reader:
            org = text(row[idx['客户所属机构']])
            tracking = text(row[idx['运单号']])
            category = text(row[48]) if len(row) > 48 else ''
            if org and tracking:
                groups.setdefault(org, Counter())[category] += 1
    totals = {}
    with open(totals_path, newline='', encoding='utf-8-sig') as f:
        for i, row in enumerate(csv.reader(f)):
            if i and len(row) >= 2:
                totals[text(row[0])] = int(number(row[1]))
    cats = [x for x in ('无应收', '金额异常', '') if any(g[x] for g in groups.values())]
    pivot_output = os.path.join(output_dir, '无应收明细透视表.csv')
    with open(pivot_output, 'w', newline='', encoding='utf-8-sig') as f:
        out = csv.writer(f)
        out.writerow(['客户所属机构'] + [('空白' if not x else x) for x in cats] + ['合计', '总票数', '占比'])
        for org in sorted(groups):
            nums = [groups[org][x] for x in cats]
            total = sum(nums); denominator = totals.get(org, 0)
            out.writerow([org] + nums + [total, denominator, total / denominator if denominator else 0])
    return detail_output, pivot_output

def export_combined_csv(csv_path, output, totals_path, idx, progress=None):
    """将明细和透视结果串联到一个 CSV，用于最终速度/体积测试。"""
    groups = {}
    with open(output, 'w', newline='', encoding='utf-8-sig') as out_file:
        out = csv.writer(out_file)
        # 先扫描明细并汇总透视数据，最终 CSV 将透视表放在最上方。
        with open(csv_path, newline='', encoding='utf-8-sig') as source:
            reader = csv.reader(source); header = next(reader)
            detail_rows = []
            for row_number, row in enumerate(reader, 1):
                detail_rows.append(row)
                org = text(row[idx['客户所属机构']])
                tracking = text(row[idx['运单号']])
                category = text(row[48]) if len(row) > 48 else ''
                if org and tracking:
                    groups.setdefault(org, Counter())[category] += 1
                if progress and row_number % 100000 == 0:
                    progress(1, 1, f'步骤 10/10：正在汇总透视数据，已读取 {row_number} 行')
        totals = {}
        with open(totals_path, newline='', encoding='utf-8-sig') as totals_file:
            for index, row in enumerate(csv.reader(totals_file)):
                if index and len(row) >= 2:
                    totals[text(row[0])] = int(number(row[1]))
        categories = [x for x in ('无应收', '金额异常', '')
                      if any(group[x] for group in groups.values())]
        out.writerow(['【无应收明细透视表】'])
        out.writerow(['客户所属机构'] + [('空白' if not x else x) for x in categories]
                     + ['合计', '总票数', '占比'])
        for org in sorted(groups):
            counts = [groups[org][key] for key in categories]
            total = sum(counts); denominator = totals.get(org, 0)
            ratio = total / denominator if denominator else 0
            out.writerow([org] + counts + [total, denominator, f'{ratio:.2%}'])
        out.writerow([]); out.writerow([]); out.writerow(['【无应收明细】']); out.writerow(header)
        out.writerows(detail_rows)
        return output
        totals = {}
        with open(totals_path, newline='', encoding='utf-8-sig') as totals_file:
            for index, row in enumerate(csv.reader(totals_file)):
                if index and len(row) >= 2:
                    totals[text(row[0])] = int(number(row[1]))
        categories = [x for x in ('无应收', '金额异常', '')
                      if any(group[x] for group in groups.values())]
        out.writerow([])
        out.writerow([])
        out.writerow(['【无应收明细透视表】'])
        out.writerow(['客户所属机构'] + [('空白' if not x else x) for x in categories]
                     + ['合计', '总票数', '占比'])
        for org in sorted(groups):
            counts = [groups[org][key] for key in categories]
            total = sum(counts)
            denominator = totals.get(org, 0)
            out.writerow([org] + counts + [total, denominator,
                                           total / denominator if denominator else 0])
    return output

def export_step_xlsx(csv_path, output):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet('无应收明细')
    with open(csv_path, newline='', encoding='utf-8-sig') as source:
        for row in csv.reader(source):
            ws.append(row)
    wb.save(output); wb.close()

def export_pivot_xlsx(csv_path, output, totals_path, idx):
    groups = {}
    with open(csv_path, newline='', encoding='utf-8-sig') as source:
        reader = csv.reader(source); header = next(reader)
        for row in reader:
            org = text(row[idx['客户所属机构']]); tracking = text(row[idx['运单号']])
            category = text(row[48]) if len(row) > 48 else ''
            if org and tracking:
                groups.setdefault(org, Counter())[category] += 1
    totals = {}
    with open(totals_path, newline='', encoding='utf-8-sig') as source:
        for n, row in enumerate(csv.reader(source)):
            if n and len(row) >= 2: totals[text(row[0])] = int(number(row[1]))
    wb = Workbook(write_only=True); ws = wb.create_sheet('无应收明细透视表')
    categories = [x for x in ('无应收', '金额异常', '') if any(g[x] for g in groups.values())]
    ws.append(['客户所属机构'] + [('空白' if not x else x) for x in categories] + ['合计', '总票数', '占比'])
    for org in sorted(groups):
        counts = [groups[org][x] for x in categories]; total = sum(counts); denominator = totals.get(org, 0)
        ws.append([org] + counts + [total, denominator, total / denominator if denominator else 0])
    wb.save(output); wb.close()

def export_step_package(work, output_dir, totals_path, idx, progress=None, profile='no_receivable'):
    config = REPORT_PROFILES[profile]
    package_dir = os.path.join(output_dir, config['detail'] + '-步骤结果')
    os.makedirs(package_dir, exist_ok=True)
    files = []
    for step in range(1, 9 if profile == 'no_salesperson_cost' else 10):
        csv_path = os.path.join(work, f'step{step}.csv')
        if not os.path.isfile(csv_path): continue
        output = os.path.join(package_dir, f"{config['detail']}-步骤{step}.csv")
        if progress: progress(1, 1, f'正在整理步骤 CSV：步骤 {step}/9')
        with open(csv_path, 'rb') as source, open(output, 'wb') as target:
            target.write(source.read())
        files.append(output)
    final_xlsx = os.path.join(package_dir, config['detail'] + '-最终结果.xlsx')
    if progress: progress(1, 1, '正在生成最终明细和透视表 XLSX')
    final_step = 8 if profile == 'no_salesperson_cost' else 9
    export_xlsx(os.path.join(work, f'step{final_step}.csv'), final_xlsx, work, totals_path, idx, profile=profile)
    files.append(final_xlsx)
    zip_path = os.path.join(output_dir, config['zip'])
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as archive:
        for file_path in files:
            archive.write(file_path, os.path.basename(file_path))
    return zip_path

def export_xlsx_bundle(work_map, output, progress=None):
    wb = Workbook(write_only=True)
    for profile, info in work_map.items():
        config = REPORT_PROFILES[profile]
        csv_path, totals_path, idx, category_col = info
        detail = wb.create_sheet(config['detail'])
        groups = {}
        with open(csv_path, newline='', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            header = next(reader)
            detail.append(header)
            for row in reader:
                detail.append(row)
                org = text(row[idx['客户所属机构']])
                tracking = text(row[idx['运单号']])
                if org and tracking:
                    groups.setdefault(org, Counter())[text(row[category_col])] += 1
        totals = {}
        with open(totals_path, newline='', encoding='utf-8-sig') as f:
            for n, row in enumerate(csv.reader(f)):
                if n and len(row) >= 2:
                    totals[text(row[0])] = int(number(row[1]))
        cats = [x for x in (config['category_zero'], '金额异常', '') if any(g[x] for g in groups.values())]
        pivot = wb.create_sheet(config['pivot'])
        pivot.append(['客户所属机构'] + [('空白' if not x else x) for x in cats] + ['合计', '总票数', '占比'])
        for org in sorted(groups):
            counts = [groups[org][x] for x in cats]
            total = sum(counts)
            denominator = totals.get(org, 0)
            pivot.append([org] + counts + [total, denominator, total / denominator if denominator else 0])
    if progress:
        progress(1, 1, '正在保存包含两个明细表和两个透视表的最终 XLSX')
    wb.save(output)
    wb.close()

def run_web_csv_bundle(task, progress=None):
    root = os.path.join(task['out_dir'], 'csv_work')
    work_map = {}
    for profile in ('no_receivable', 'no_salesperson_cost'):
        config = REPORT_PROFILES[profile]
        work = os.path.join(root, profile)
        os.makedirs(work, exist_ok=True)
        csv0 = os.path.join(work, 'step0.csv')
        totals = os.path.join(work, 'totals.csv')
        header, idx = read_sheet(task['source_path'], task['sheet_name'], csv0, totals, profile)
        current = csv0
        max_filter = 7 if profile == 'no_receivable' else 6
        for step in range(1, max_filter + 1):
            nxt = os.path.join(work, f'step{step}.csv')
            kept, removed = process_step(current, nxt, os.path.join(work, f'临时删除_步骤{step}.csv'), step, header, idx, profile)
            if progress:
                progress(1, 1, f'{config["detail"]}：步骤 {step} 完成，保留 {kept} 行，删除 {removed} 行')
            current = nxt
        category_step = 8 if profile == 'no_receivable' else 7
        category_path = os.path.join(work, f'step{category_step}.csv')
        add_category(current, category_path, idx, profile)
        current = category_path
        final_step = 9 if profile == 'no_receivable' else 8
        final_csv = os.path.join(work, f'step{final_step}.csv')
        if current != final_csv:
            with open(current, 'rb') as source, open(final_csv, 'wb') as target:
                target.write(source.read())
        category_col = 48 if profile == 'no_receivable' else len(header)
        work_map[profile] = (final_csv, totals, idx, category_col)
    package_dir = os.path.join(task['out_dir'], '报表组步骤结果')
    os.makedirs(package_dir, exist_ok=True)
    files = []
    for profile, info in work_map.items():
        config = REPORT_PROFILES[profile]
        work = os.path.dirname(info[0])
        for step in range(1, 10):
            path = os.path.join(work, f'step{step}.csv')
            if os.path.isfile(path):
                out = os.path.join(package_dir, f'{config["detail"]}-步骤{step}.csv')
                with open(path, 'rb') as source, open(out, 'wb') as target:
                    target.write(source.read())
                files.append(out)
    final_xlsx = os.path.join(package_dir, '报表组-最终结果.xlsx')
    export_xlsx_bundle(work_map, final_xlsx, progress)
    files.append(final_xlsx)
    zip_path = os.path.join(task['out_dir'], '报表组-步骤结果.zip')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as archive:
        for path in files:
            archive.write(path, os.path.basename(path))
    return zip_path, os.path.basename(zip_path)

def run_web_csv_step(task, step, progress=None):
    """网页报表任务的 CSV 分步入口；只在最后一步导出 XLSX。"""
    profile = task.get('report_profile', 'no_receivable')
    work = task['csv_work']
    os.makedirs(work, exist_ok=True)
    source = task['source_path']
    if step == 1 and not os.path.isfile(os.path.join(work, 'step0.csv')):
        if progress: progress(1, 1, '正在读取所选工作表并转换为 CSV')
        header, idx = read_sheet(source, task['sheet_name'], os.path.join(work, 'step0.csv'), os.path.join(work, 'totals.csv'), profile)
        task['csv_header'] = header; task['csv_idx'] = idx
    else:
        with open(os.path.join(work, 'step0.csv'), newline='', encoding='utf-8-sig') as f:
            task['csv_header'] = next(csv.reader(f))
        task['csv_idx'] = {x:i for i,x in enumerate(task['csv_header'])}
    current = os.path.join(work, f'step{step-1}.csv')
    if step == 1: current = os.path.join(work, 'step0.csv')
    if step == 10: current = os.path.join(work, 'step9.csv')
    max_filter_step = 6 if profile == 'no_salesperson_cost' else 7
    if step <= max_filter_step:
        nxt = os.path.join(work, f'step{step}.csv')
        kept, removed = process_step(current, nxt, os.path.join(work, f'临时删除_步骤{step}.csv'), step, task['csv_header'], task['csv_idx'], profile)
        if progress: progress(1, 1, f'步骤 {step}/10：已完成，保留 {kept} 行，删除 {removed} 行')
        return nxt, f"{REPORT_PROFILES[profile]['detail']}-步骤{step}.csv"
    category_step = 7 if profile == 'no_salesperson_cost' else 8
    if step == category_step:
        nxt = os.path.join(work, f'step{step}.csv'); add_category(current, nxt, task['csv_idx'], profile)
        if progress: progress(1, 1, f'步骤 {step}/10：已新增分类列')
        return nxt, f"{REPORT_PROFILES[profile]['detail']}-步骤{step}.csv"
    pivot_step = 8 if profile == 'no_salesperson_cost' else 9
    if step == pivot_step:
        if progress: progress(1, 1, f'步骤 {step}/10：整理最终明细 CSV，透视表将在最终总表中生成')
        nxt = os.path.join(work, f'step{step}.csv')
        if current != nxt:
            with open(current, 'rb') as source, open(nxt, 'wb') as target:
                target.write(source.read())
        return nxt, f"{REPORT_PROFILES[profile]['detail']}-步骤{step}.csv"
    final = os.path.join(task['out_dir'], '无应收明细-最终总表.xlsx')
    final_csv = os.path.join(task['out_dir'], '无应收明细-最终测试.csv')
    package = export_step_package(work, task['out_dir'], os.path.join(work, 'totals.csv'), task['csv_idx'], progress, profile)
    if progress: progress(1, 1, '步骤 10/10：已生成步骤 CSV 和最终 XLSX 结果包')
    return package, os.path.basename(package)
def main():
    p=argparse.ArgumentParser(description='报表组 CSV 高速终端版'); p.add_argument('input',nargs='?'); p.add_argument('-s','--sheet'); p.add_argument('-o','--output-dir'); p.add_argument('--no-pause',action='store_true'); a=p.parse_args()
    source=clean_path(a.input or input('请输入 Excel 文件路径：')); source=os.path.abspath(source)
    if not os.path.isfile(source): print('文件不存在：'+source,file=sys.stderr); return 2
    sheet=a.sheet or choose_sheet(source); outdir=os.path.abspath(a.output_dir or os.path.dirname(source)); os.makedirs(outdir,exist_ok=True); work=os.path.join(outdir,'.report_csv_work'); os.makedirs(work,exist_ok=True)
    current=os.path.join(work,'step0.csv'); totals=os.path.join(work,'totals.csv'); print('首次读取 Excel 并转换 CSV…'); header,idx=read_sheet(source,sheet,current,totals)
    for step in range(1,10):
        print(f'\n===== 步骤 {step}/10：{STEPS[step-1]} ====='); nxt=os.path.join(work,f'step{step}.csv')
        if step <= 7: kept,removed=process_step(current,nxt,os.path.join(work,f'临时删除_步骤{step}.csv'),step,header,idx); print(f'完成：保留 {kept} 行，删除 {removed} 行')
        elif step == 8: add_category(current,nxt,idx); print('完成：已新增无应收分类列')
        else: print('完成：透视数据将在最终导出时生成')
        if step < 9 and not a.no_pause and input('按回车继续，输入 q 退出：').strip().lower()=='q': return 0
        current=nxt if os.path.exists(nxt) else current
    print('\n===== 步骤 10/10：导出最终总表 XLSX =====')
    final_xlsx = os.path.join(outdir, '无应收明细-最终总表.xlsx')
    export_xlsx(current, final_xlsx, work, totals, idx)
    print('最终总表 Excel（含明细和透视表工作表）：' + final_xlsx)
    print('删除记录和中间 CSV 保存在：' + work)
    return 0
if __name__=='__main__': raise SystemExit(main())
