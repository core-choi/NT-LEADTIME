#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NT Service Center - Lead Time Dashboard Builder

Folder structure:
  data/
    2025/01~12/   (each: 서비스_리드타임_현황.xlsx, RoReport.xlsx)
    2026/01~12/
    ...

Auto-detects: latest month = current, previous month = prev, prev year = trend
"""

import pandas as pd
import re
import os
import sys
from datetime import datetime, timedelta, timezone

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

CI_MODE = '--ci' in sys.argv
KST = timezone(timedelta(hours=9))

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, 'data')
OUTPUT_DIR = os.path.join(BASE, 'output')
TEMPLATE = os.path.join(BASE, 'template.html')

STYPE_MAP = {'입고 Proactive-RITA': 'RITA', '미입고 Proactive-RITA': 'RITA'}
STYPES = ['Maintenance', 'Recall/TC', '경수리', '중수리', '소음', '진단', '사고수리', 'RITA']
STYPE_RO_MAP = {'입고 Proactive-RITA': 'RITA', '미입고 Proactive-RITA': 'RITA', 'Recall/TC': 'Recall'}
RO_STYPES = ['Maintenance', 'Recall', '경수리', '중수리', '소음', '진단', '사고수리', 'RITA']
BRANCHES = ['전주', '평택', '군산', '목포', '서산']
INC_LABELS = ['상담완료', '가정산 완료', 'End Control 요청', '정비시작', '상담시작']


def log(msg):
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode('ascii', 'replace').decode('ascii'))


# ============================================================
# Data discovery - find all year/month folders with data
# ============================================================
def find_all_months(data_dir):
    """Scan data/ for year/month folders, return sorted list of (year, month, path)"""
    results = []
    if not os.path.exists(data_dir):
        return results
    for year_name in sorted(os.listdir(data_dir)):
        year_path = os.path.join(data_dir, year_name)
        if not os.path.isdir(year_path):
            continue
        try:
            year = int(year_name)
        except ValueError:
            continue
        for month_name in sorted(os.listdir(year_path)):
            month_path = os.path.join(year_path, month_name)
            if not os.path.isdir(month_path):
                continue
            try:
                month = int(month_name)
                if month < 1 or month > 12:
                    continue
            except ValueError:
                continue
            # Check if folder has xlsx files
            has_lt = any('리드타임' in f or '서비스_리드타임' in f for f in os.listdir(month_path) if f.endswith('.xlsx'))
            has_ro = any('ro' in f.lower() and f.endswith('.xlsx') for f in os.listdir(month_path))
            if has_lt or has_ro:
                results.append((year, month, month_path))
    return results


def find_current_and_prev(all_months):
    """From sorted month list, find current (latest) and previous"""
    if not all_months:
        return None, None
    current = all_months[-1]  # latest
    prev = all_months[-2] if len(all_months) >= 2 else None
    return current, prev


def find_trend_year(all_months, current):
    """Find 12 months of data for trend (previous year of current)"""
    if not current:
        return None
    target_year = current[0] - 1
    trend_months = [m for m in all_months if m[0] == target_year]
    if trend_months:
        return target_year
    # Also check current year (for partial year trend)
    curr_months = [m for m in all_months if m[0] == current[0]]
    if len(curr_months) >= 2:
        return current[0]
    return None


# ============================================================
# File readers
# ============================================================
def read_lt(folder):
    if not os.path.exists(folder):
        return None
    for f in os.listdir(folder):
        if ('리드타임' in f or '서비스_리드타임' in f) and f.endswith('.xlsx') and not f.startswith('~$'):
            return pd.read_excel(os.path.join(folder, f))
    return None


def read_ro(folder):
    if not os.path.exists(folder):
        return None
    for f in os.listdir(folder):
        if 'ro' in f.lower() and f.endswith('.xlsx') and not f.startswith('~$'):
            df = pd.read_excel(os.path.join(folder, f), header=1)
            df.columns = df.columns.str.strip()
            return df
    return None


# ============================================================
# Data processing
# ============================================================
def process_lt_to_js(df):
    lines = []
    for _, r in df.iterrows():
        bn = str(r['지점명']).replace('AS_', '')
        st = STYPE_MAP.get(str(r['서비스타입']), str(r['서비스타입']))
        lines.append(
            f'  {{branch:"{bn}",type:"{st}",ro:{int(r["RO건수"])},'
            f'slt:{float(r["서비스L/T"]):.2f},rlt:{float(r["예약 L/T"]):.2f},'
            f'clt:{float(r["고객대기 L/T"]):.2f},colt:{float(r["상담 L/T"]):.2f},'
            f'wlt:{float(r["정비대기 L/T"]):.2f},mlt:{float(r["정비 L/T"]):.2f},'
            f'dlt:{float(r["출고대기 L/T"]):.2f},blt:{float(r["정산 L/T"]):.2f}}}'
        )
    return '[\n' + ',\n'.join(lines) + '\n]'


def process_ro(df):
    total = len(df)
    invoiced = len(df[df['RO상태'] == '인보이스 완료'])
    cancelled = len(df[df['RO상태'] == 'RO취소'])
    incomplete = total - invoiced - cancelled
    by_branch = {}
    for b in BRANCHES:
        bdf = df[df['AS지점'].astype(str).str.replace('AS_', '', regex=False) == b]
        bt = len(bdf); bi = len(bdf[bdf['RO상태'] == '인보이스 완료']); bc = len(bdf[bdf['RO상태'] == 'RO취소'])
        by_branch[b] = {'total': bt, 'invoiced': bi, 'cancelled': bc, 'incomplete': bt - bi - bc}
    inc_df = df[(df['RO상태'] != '인보이스 완료') & (df['RO상태'] != 'RO취소')]
    inc_detail = {}
    for lbl in INC_LABELS:
        cnt = len(inc_df[inc_df['RO상태'] == lbl])
        if cnt > 0:
            inc_detail[lbl] = cnt
    inc_by_branch = {}
    for b in BRANCHES:
        bdf = inc_df[inc_df['AS지점'].astype(str).str.replace('AS_', '', regex=False) == b]
        bd = {}
        for lbl in INC_LABELS:
            bd[lbl] = len(bdf[bdf['RO상태'] == lbl])
        inc_by_branch[b] = bd
    df_ex = df[df['RO상태'] != 'RO취소'].copy()
    df_ex['stype_key'] = df_ex['서비스타입'].map(lambda x: STYPE_RO_MAP.get(str(x), str(x)))
    by_stype = {}
    for b in BRANCHES:
        bdf = df_ex[df_ex['AS지점'].astype(str).str.replace('AS_', '', regex=False) == b]
        bs = {}
        for st in RO_STYPES:
            sdf = bdf[bdf['stype_key'] == st]
            inv = len(sdf[sdf['RO상태'] == '인보이스 완료'])
            bs[st] = {'inv': inv, 'inc': len(sdf) - inv}
        by_stype[b] = bs
    max_date = pd.to_datetime(df['RO갱신일시'], errors='coerce').max()
    update_str = max_date.strftime('%Y-%m-%d %H:%M') if pd.notna(max_date) else datetime.now(KST).strftime('%Y-%m-%d %H:%M')
    return {
        'total': total, 'invoiced': invoiced, 'cancelled': cancelled,
        'incomplete': incomplete, 'incDetail': inc_detail,
        'byBranch': by_branch, 'incByBranch': inc_by_branch,
        'byStype': by_stype, 'updateDate': update_str
    }


def ro_to_js(ro_data, var_name='RO_STATUS'):
    lines = [f'const {var_name} = {{']
    lines.append(f'  total:{ro_data["total"]},invoiced:{ro_data["invoiced"]},cancelled:{ro_data["cancelled"]},incomplete:{ro_data["incomplete"]},')
    inc_parts = ','.join(f'"{k}":{v}' for k, v in ro_data['incDetail'].items())
    lines.append(f'  incDetail:{{{inc_parts}}},')
    bb = []
    for b in BRANCHES:
        d = ro_data['byBranch'].get(b, {'total':0,'invoiced':0,'cancelled':0,'incomplete':0})
        bb.append(f'    "{b}":{{total:{d["total"]},invoiced:{d["invoiced"]},cancelled:{d["cancelled"]},incomplete:{d["incomplete"]}}}')
    lines.append('  byBranch:{\n' + ',\n'.join(bb) + '\n  },')
    ib = []
    for b in BRANCHES:
        bd = ro_data['incByBranch'].get(b, {})
        items = ','.join(f'"{k}":{bd.get(k,0)}' for k in INC_LABELS)
        ib.append(f'    "{b}":{{{items}}}')
    lines.append('  incByBranch:{\n' + ',\n'.join(ib) + '\n  },')
    bs = []
    for b in BRANCHES:
        bd = ro_data['byStype'].get(b, {})
        items = ','.join(f'{k}:{{inv:{bd.get(k,{}).get("inv",0)},inc:{bd.get(k,{}).get("inc",0)}}}' for k in RO_STYPES)
        bs.append(f'    "{b}":{{{items}}}')
    lines.append('  byStype:{\n' + ',\n'.join(bs) + '\n  }')
    lines.append('\n};')
    return '\n'.join(lines)


def process_yearly_trend(data_dir, year):
    stype_monthly = {st: [None]*12 for st in STYPES}
    found = []
    for mi in range(12):
        m = f'{mi+1:02d}'
        m_path = os.path.join(data_dir, str(year), m)
        lt = read_lt(m_path)
        if lt is None:
            continue
        found.append(m)
        for st in STYPES:
            mapped = [st] + [k for k, v in STYPE_MAP.items() if v == st]
            rows = lt[lt['서비스타입'].isin(mapped)]
            if len(rows) == 0:
                continue
            total_ro = rows['RO건수'].sum()
            if total_ro == 0:
                continue
            stype_monthly[st][mi] = round((rows['서비스L/T'] * rows['RO건수']).sum() / total_ro / 24, 2)
    return stype_monthly, found


def trend_to_js(stype_monthly):
    lines = ['const TREND_DATA = {']
    lines.append('  months:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],')
    lines.append('  series:[')
    for st in STYPES:
        vals = ','.join('null' if v is None else str(v) for v in stype_monthly[st])
        lines.append(f'    {{type:"{st}",v:[{vals}]}},')
    lines.append('  ]')
    lines.append('\n};')
    return '\n'.join(lines)


# ============================================================
# Main build
# ============================================================
def build():
    log('=' * 55)
    log('  NT Dashboard - Auto Build')
    log('=' * 55)

    if not os.path.exists(TEMPLATE):
        log(f'\n[ERROR] template.html not found: {TEMPLATE}')
        return False

    # Discover data
    log('\n--- Scanning data folders ---')
    all_months = find_all_months(DATA_DIR)

    if not all_months:
        log('[ERROR] No data found in data/ folder!')
        log(f'  Expected: data/2026/03/ with xlsx files')
        return False

    for y, m, p in all_months:
        files = [f for f in os.listdir(p) if f.endswith('.xlsx') and not f.startswith('~$')]
        log(f'  {y}/{m:02d}: {", ".join(files)}')

    # Determine current and previous
    current, prev = find_current_and_prev(all_months)
    log(f'\n  >> Current: {current[0]}/{current[1]:02d}')
    if prev:
        log(f'  >> Previous: {prev[0]}/{prev[1]:02d}')
    else:
        log(f'  >> Previous: none')

    # Read current month
    curr_lt = read_lt(current[2])
    curr_ro_df = read_ro(current[2])
    if curr_lt is None or curr_ro_df is None:
        log(f'[ERROR] Current month data incomplete: {current[2]}')
        if curr_lt is None:
            log(f'  Missing: 서비스_리드타임_현황.xlsx')
        if curr_ro_df is None:
            log(f'  Missing: RoReport.xlsx')
        return False
    log(f'\n[OK] Current ({current[0]}/{current[1]:02d}): LT {len(curr_lt)} rows, RO {len(curr_ro_df)} rows')

    # Read previous month
    has_prev = False
    prev_lt = None
    prev_ro_df = None
    if prev:
        prev_lt = read_lt(prev[2])
        prev_ro_df = read_ro(prev[2])
        has_prev = prev_lt is not None and prev_ro_df is not None
        if has_prev:
            log(f'[OK] Previous ({prev[0]}/{prev[1]:02d}): LT {len(prev_lt)} rows, RO {len(prev_ro_df)} rows')
        else:
            log(f'[WARN] Previous month data incomplete')

    # Trend year
    trend_year = find_trend_year(all_months, current)
    has_trend = False
    found_trend = []
    if trend_year:
        stype_monthly, found_trend = process_yearly_trend(DATA_DIR, trend_year)
        has_trend = len(found_trend) > 0
        if has_trend:
            log(f'[OK] Trend ({trend_year}): {len(found_trend)} months')

    # Process
    curr_ro = process_ro(curr_ro_df)
    log(f'\n[DATA] RO: total {curr_ro["total"]}, done {curr_ro["invoiced"]}, pending {curr_ro["incomplete"]}')

    js_raw = f'const RAW = {process_lt_to_js(curr_lt)};'
    js_ro = ro_to_js(curr_ro, 'RO_STATUS')

    if has_prev:
        prev_ro = process_ro(prev_ro_df)
        js_prev_raw = f'const PREV_RAW = {process_lt_to_js(prev_lt)};'
        bb = []
        for b in BRANCHES:
            d = prev_ro['byBranch'].get(b, {'total':0,'invoiced':0,'cancelled':0,'incomplete':0})
            bb.append(f'    "{b}":{{total:{d["total"]},invoiced:{d["invoiced"]},cancelled:{d["cancelled"]},incomplete:{d["incomplete"]}}}')
        js_prev_ro = (
            f'const PREV_RO = {{\n'
            f'  total:{prev_ro["total"]},invoiced:{prev_ro["invoiced"]},'
            f'cancelled:{prev_ro["cancelled"]},incomplete:{prev_ro["incomplete"]},\n'
            f'  byBranch:{{\n' + ',\n'.join(bb) + '\n  }\n};'
        )
        prev_month_label = f'{prev[1]}월'
        log(f'  Previous label: {prev_month_label}')
    else:
        js_prev_raw = 'const PREV_RAW = [];'
        js_prev_ro = 'const PREV_RO = {\n  total:0,invoiced:0,cancelled:0,incomplete:0,\n  byBranch:{}\n};'
        prev_month_label = '전월'

    # Read template
    with open(TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()

    # Replace data
    html = re.sub(r'const RAW = \[.*?\];', js_raw, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const PREV_RAW = \[.*?\];', js_prev_raw, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const RO_STATUS = \{.*?\n\};', js_ro, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const PREV_RO = \{.*?\n\};', js_prev_ro, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const PREV_MONTH_LABEL = ".*?";', f'const PREV_MONTH_LABEL = "{prev_month_label}";', html, count=1)

    if has_trend:
        js_trend = trend_to_js(stype_monthly)
        html = re.sub(r'const TREND_DATA = \{.*?\n\};', js_trend, html, count=1, flags=re.DOTALL)
        html = re.sub(r'\d{4}년 서비스타입별 총서비스 L/T 월간 변화', f'{trend_year}년 서비스타입별 총서비스 L/T 월간 변화', html)

    # Build timestamp (KST)
    build_time = datetime.now(KST).strftime('%Y-%m-%d %H:%M')
    html = html.replace('BUILD_TIMESTAMP', build_time)
    html = re.sub(r'<span>\d{4}-\d{2}-\d{2} \d{2}:\d{2} [^<]*</span>', f'<span>{build_time} 업데이트</span>', html, count=1)

    # Output
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for name in ['index.html', '서비스_리드타임_대시보드.html']:
        with open(os.path.join(OUTPUT_DIR, name), 'w', encoding='utf-8') as f:
            f.write(html)

    log(f'\n[OK] Build complete!')
    log(f'  Current: {current[0]}/{current[1]:02d}')
    if has_prev:
        log(f'  vs Prev:  {prev[0]}/{prev[1]:02d}')
    log(f'  Time: {build_time} KST')
    log(f'  Output: {OUTPUT_DIR}')
    log('=' * 55)
    return True


if __name__ == '__main__':
    try:
        success = build()
        if not success:
            log('\n[HELP] Expected folder structure:')
            log('  data/')
            log('    2025/01~12/  (서비스_리드타임_현황.xlsx, RoReport.xlsx)')
            log('    2026/01~12/')
    except Exception as e:
        log(f'\n[ERROR] Build failed: {e}')
        import traceback
        traceback.print_exc()

    if not CI_MODE:
        input('\nPress Enter to close...')
