# -*- coding: utf-8 -*-
"""报表组高速终端版：Excel 只读一次，中间全程 CSV，最后导出 XLSX。"""
import argparse, csv, os, re, sys
from collections import Counter
from openpyxl import Workbook, load_workbook

EXCLUDED = ('风驰-数据同步', 'YX订舱', '风驰-卖柜', '李雪原', '龙行-清关', '李增韬')
STEPS = (
    '删除操作状态签入', '筛选应收单价小于等于1', '删除客户简称关键词',
    '删除业务员华南KA', '删除备注J000、无应收、免费补发',
    '删除配仓单号刘丹整柜', '删除整柜且应收金额大于10000',
    '新增无应收分类', '生成透视表', '导出最终CSV',
)

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
def read_sheet(path, sheet, csv_path, totals_path):
    w = load_workbook(path, read_only=True, data_only=True); ws = w[sheet]
    rows = ws.iter_rows(values_only=True); header = [text(x) for x in next(rows)]
    idx = {x:i for i,x in enumerate(header)}; required = ('操作状态','应收单价','客户简称','业务员','自定义备注','配仓单号','销售产品','应收金额','客户所属机构','运单号')
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
def process_step(src, dst, log_path, step, header, idx):
    rules = {
      1: lambda v: '签入' not in v[idx['操作状态']],
      2: lambda v: text(v[idx['应收单价']]) != '' and number(v[idx['应收单价']]) <= 1,
      3: lambda v: not any(k in text(v[idx['客户简称']]) for k in EXCLUDED),
      4: lambda v: '华南KA' not in text(v[idx['业务员']]),
      5: lambda v: not remark_remove(v[idx['自定义备注']]),
      6: lambda v: '刘丹整柜' not in text(v[idx['配仓单号']]),
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
def add_category(src, dst, idx):
    with open(src, newline='', encoding='utf-8-sig') as fi, open(dst, 'w', newline='', encoding='utf-8-sig') as fo:
        reader=csv.reader(fi); out=csv.writer(fo); h=next(reader); h = h + [''] * max(0, 49 - len(h)); h[48] = '无应收&金额异常'; out.writerow(h)
        for v in reader:
            price=number(v[idx['应收单价']]); amount=number(v[idx['应收金额']]); v = v + [''] * max(0, 49 - len(v)); v[48] = '金额异常' if amount>0 else ('无应收' if price==0 else ''); out.writerow(v)
def export_xlsx(csv_path, output, log_dir, totals_path, idx, progress=None):
    wb=Workbook(write_only=True); detail=wb.create_sheet('无应收明细')
    groups={}; totals={}
    if progress: progress(1, 1, '步骤 10/10：正在写入最终无应收明细')
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        r=csv.reader(f); h=next(r); detail.append(h)
        for row_number, v in enumerate(r, 1):
            detail.append(v); org=text(v[idx['客户所属机构']]); cat=text(v[48]); tracking=text(v[idx['运单号']]);
            if org: groups.setdefault(org, Counter()); groups[org][cat] += 1
            if progress and row_number % 100000 == 0:
                progress(1, 1, f'步骤 10/10：最终明细已写入 {row_number} 行')
    with open(totals_path, newline='', encoding='utf-8-sig') as f:
        for i,row in enumerate(csv.reader(f)):
            if i: totals[text(row[0])] = int(number(row[1]))
    cats=[x for x in ('无应收','金额异常','') if any(g[x] for g in groups.values())]; pivot=wb.create_sheet('无应收明细透视表'); pivot.append(['客户所属机构']+[('空白' if not x else x) for x in cats]+['合计','总票数','占比'])
    for org in sorted(groups):
        nums=[groups[org][x] for x in cats]; total=sum(nums); den=totals.get(org,0); pivot.append([org]+nums+[total,den,total/den if den else 0])
    if progress: progress(1, 1, '步骤 10/10：正在写入透视表和各步骤删除记录')
    for name in sorted(os.listdir(log_dir)):
        if name.startswith('临时删除_步骤') and name.endswith('.csv'):
            ws=wb.create_sheet(name[:-4]);
            with open(os.path.join(log_dir,name),newline='',encoding='utf-8-sig') as f:
                for row in csv.reader(f): ws.append(row)
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

def run_web_csv_step(task, step, progress=None):
    """网页报表任务的 CSV 分步入口；只在最后一步导出 XLSX。"""
    work = task['csv_work']
    os.makedirs(work, exist_ok=True)
    source = task['source_path']
    if step == 1 and not os.path.isfile(os.path.join(work, 'step0.csv')):
        if progress: progress(1, 1, '正在读取所选工作表并转换为 CSV')
        header, idx = read_sheet(source, task['sheet_name'], os.path.join(work, 'step0.csv'), os.path.join(work, 'totals.csv'))
        task['csv_header'] = header; task['csv_idx'] = idx
    else:
        with open(os.path.join(work, 'step0.csv'), newline='', encoding='utf-8-sig') as f:
            task['csv_header'] = next(csv.reader(f))
        task['csv_idx'] = {x:i for i,x in enumerate(task['csv_header'])}
    current = os.path.join(work, f'step{step-1}.csv')
    if step == 1: current = os.path.join(work, 'step0.csv')
    if step == 10: current = os.path.join(work, 'step9.csv')
    if step <= 7:
        nxt = os.path.join(work, f'step{step}.csv')
        kept, removed = process_step(current, nxt, os.path.join(work, f'临时删除_步骤{step}.csv'), step, task['csv_header'], task['csv_idx'])
        if progress: progress(1, 1, f'步骤 {step}/10：已完成，保留 {kept} 行，删除 {removed} 行')
        return nxt, f'无应收明细-步骤{step}.csv'
    if step == 8:
        nxt = os.path.join(work, 'step8.csv'); add_category(current, nxt, task['csv_idx'])
        if progress: progress(1, 1, '步骤 8/10：已新增无应收&金额异常列')
        return nxt, '无应收明细-步骤8.csv'
    if step == 9:
        if progress: progress(1, 1, '步骤 9/10：透视表将在最终总表中生成')
        nxt = os.path.join(work, 'step9.csv')
        if current != nxt:
            with open(current, 'rb') as source, open(nxt, 'wb') as target:
                target.write(source.read())
        return nxt, '无应收明细-步骤9.csv'
    final = os.path.join(task['out_dir'], '无应收明细-最终总表.xlsx')
    export_xlsx(current, final, work, os.path.join(work, 'totals.csv'), task['csv_idx'], progress)
    if progress: progress(1, 1, '步骤 10/10：已生成包含明细和透视表的最终总表')
    return final, os.path.basename(final)
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
