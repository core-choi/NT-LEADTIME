#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NT 서비스센터 - 리드타임 대시보드 자동 빌드
GitHub Actions 및 로컬 실행 모두 지원
"""

import pandas as pd
import re
import os
import sys
from datetime import datetime

# Windows 콘솔 UTF-8
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    try:
        os.system('chcp 65001 >nul 2>&1')
    except Exception:
        pass

CI_MODE = '--ci' in sys.argv

BASE = os.path.dirname(os.path.abspath(__file__))
PREV_DIR = os.path.join(BASE, 'data', '전월')
CURR_DIR = os.path.join(BASE, 'data', '현월')
YEAR_DIR = os.path.join(BASE, 'data', '전년')
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


def check_environment():
    log('\n--- 환경 진단 ---')
    log(f'  build.py 위치: {BASE}')
    log(f'  template.html: {"있음" if os.path.exists(TEMPLATE) else "없음 !!!"}')
    log(f'  data 폴더: {"있음" if os.path.exists(os.path.join(BASE, "data")) else "없음 !!!"}')
    for name, path in [('현월', CURR_DIR), ('전월', PREV_DIR), ('전년', YEAR_DIR)]:
        exists = os.path.exists(path)
        log(f'  data/{name}/: {"있음" if exists else "없음"}')
        if exists:
            files = [f for f in os.listdir(path) if not f.startswith('.')]
            if files:
                for f in files:
                    log(f'    - {f}')
            else:
                log(f'    (비어있음)')
    log('--- 진단 끝 ---\n')


def read_lt(folder):
    path = os.path.join(folder, '서비스_리드타임_현황.xlsx')
    if not os.path.exists(path):
        if os.path.exists(folder):
            for f in os.listdir(folder):
                if '리드타임' in f and f.endswith('.xlsx'):
                    return pd.read_excel(os.path.join(folder, f))
        return None
    return pd.read_excel(path)


def read_ro(folder):
    path = os.path.join(folder, 'RoReport.xlsx')
    if not os.path.exists(path):
        if os.path.exists(folder):
            for f in os.listdir(folder):
                if 'ro' in f.lower() and f.endswith('.xlsx'):
                    df = pd.read_excel(os.path.join(folder, f), header=1)
                    df.columns = df.columns.str.strip()
                    return df
        return None
    df = pd.read_excel(path, header=1)
    df.columns = df.columns.str.strip()
    return df


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
        bt = len(bdf)
        bi = len(bdf[bdf['RO상태'] == '인보이스 완료'])
        bc = len(bdf[bdf['RO상태'] == 'RO취소'])
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
    update_str = max_date.strftime('%Y-%m-%d %H:%M') if pd.notna(max_date) else datetime.now().strftime('%Y-%m-%d %H:%M')
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
    bb_parts = []
    for b in BRANCHES:
        d = ro_data['byBranch'].get(b, {'total': 0, 'invoiced': 0, 'cancelled': 0, 'incomplete': 0})
        bb_parts.append(f'    "{b}":{{total:{d["total"]},invoiced:{d["invoiced"]},cancelled:{d["cancelled"]},incomplete:{d["incomplete"]}}}')
    lines.append('  byBranch:{\n' + ',\n'.join(bb_parts) + '\n  },')
    ib_parts = []
    for b in BRANCHES:
        bd = ro_data['incByBranch'].get(b, {})
        items = ','.join(f'"{k}":{bd.get(k, 0)}' for k in INC_LABELS)
        ib_parts.append(f'    "{b}":{{{items}}}')
    lines.append('  incByBranch:{\n' + ',\n'.join(ib_parts) + '\n  },')
    bs_parts = []
    for b in BRANCHES:
        bd = ro_data['byStype'].get(b, {})
        items = ','.join(f'{k}:{{inv:{bd.get(k, {}).get("inv", 0)},inc:{bd.get(k, {}).get("inc", 0)}}}' for k in RO_STYPES)
        bs_parts.append(f'    "{b}":{{{items}}}')
    lines.append('  byStype:{\n' + ',\n'.join(bs_parts) + '\n  }')
    lines.append('\n};')
    return '\n'.join(lines)


def process_yearly_trend(year_dir):
    stype_monthly = {st: [None]*12 for st in STYPES}
    found_months = []
    for mi in range(12):
        m = f'{mi+1:02d}'
        lt_path = os.path.join(year_dir, m, '서비스_리드타임_현황.xlsx')
        if not os.path.exists(lt_path):
            m_dir = os.path.join(year_dir, m)
            if os.path.exists(m_dir):
                for f in os.listdir(m_dir):
                    if '리드타임' in f and f.endswith('.xlsx'):
                        lt_path = os.path.join(m_dir, f)
                        break
                else:
                    continue
            else:
                continue
        found_months.append(m)
        df = pd.read_excel(lt_path)
        for st in STYPES:
            mapped_types = [st]
            for orig, mapped in STYPE_MAP.items():
                if mapped == st:
                    mapped_types.append(orig)
            rows = df[df['서비스타입'].isin(mapped_types)]
            if len(rows) == 0:
                continue
            total_ro = rows['RO건수'].sum()
            if total_ro == 0:
                continue
            wavg_slt = (rows['서비스L/T'] * rows['RO건수']).sum() / total_ro / 24
            stype_monthly[st][mi] = round(wavg_slt, 2)
    return stype_monthly, found_months


def trend_to_js(stype_monthly):
    lines = ['const TREND_DATA = {']
    lines.append('  months:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],')
    lines.append('  series:[')
    for st in STYPES:
        vals = stype_monthly[st]
        val_str = ','.join('null' if v is None else str(v) for v in vals)
        lines.append(f'    {{type:"{st}",v:[{val_str}]}},')
    lines.append('  ]')
    lines.append('\n};')
    return '\n'.join(lines)


def extract_prev_month(prev_ro_df):
    for col in ['RO발행일시', 'RO갱신일시', '차량접수일시', '예약일시']:
        if col in prev_ro_df.columns:
            try:
                dates = pd.to_datetime(prev_ro_df[col], errors='coerce').dropna()
                if len(dates) > 0:
                    month_num = dates.dt.month.mode()
                    if len(month_num) > 0:
                        return f'{int(month_num.iloc[0])}월'
            except Exception:
                continue
    return '전월'


def build():
    log('=' * 55)
    log('  NT 서비스센터 - 리드타임 대시보드 빌드')
    log('=' * 55)
    check_environment()

    if not os.path.exists(TEMPLATE):
        log(f'[ERROR] template.html이 없습니다: {TEMPLATE}')
        return False

    curr_lt = read_lt(CURR_DIR)
    curr_ro_df = read_ro(CURR_DIR)
    if curr_lt is None or curr_ro_df is None:
        log(f'[ERROR] 현월 데이터가 없습니다! 폴더: {CURR_DIR}')
        return False
    log(f'[OK] 현월: L/T {len(curr_lt)}행, RO {len(curr_ro_df)}행')

    prev_lt = read_lt(PREV_DIR)
    prev_ro_df = read_ro(PREV_DIR)
    has_prev = prev_lt is not None and prev_ro_df is not None
    if has_prev:
        log(f'[OK] 전월: L/T {len(prev_lt)}행, RO {len(prev_ro_df)}행')
    else:
        log(f'[WARN] 전월 데이터 없음')

    has_year = os.path.exists(YEAR_DIR)
    found_months = []
    if has_year:
        stype_monthly, found_months = process_yearly_trend(YEAR_DIR)
        if found_months:
            log(f'[OK] 전년 트렌드: {len(found_months)}개월')
        else:
            has_year = False

    curr_ro = process_ro(curr_ro_df)
    log(f'[DATA] RO: 총 {curr_ro["total"]}건, 완료 {curr_ro["invoiced"]}건')

    js_raw = f'const RAW = {process_lt_to_js(curr_lt)};'
    js_ro = ro_to_js(curr_ro, 'RO_STATUS')

    if has_prev:
        prev_ro = process_ro(prev_ro_df)
        js_prev_raw = f'const PREV_RAW = {process_lt_to_js(prev_lt)};'
        bb_parts = []
        for b in BRANCHES:
            d = prev_ro['byBranch'].get(b, {'total': 0, 'invoiced': 0, 'cancelled': 0, 'incomplete': 0})
            bb_parts.append(f'    "{b}":{{total:{d["total"]},invoiced:{d["invoiced"]},cancelled:{d["cancelled"]},incomplete:{d["incomplete"]}}}')
        js_prev_ro = (
            f'const PREV_RO = {{\n'
            f'  total:{prev_ro["total"]},invoiced:{prev_ro["invoiced"]},'
            f'cancelled:{prev_ro["cancelled"]},incomplete:{prev_ro["incomplete"]},\n'
            f'  byBranch:{{\n' + ',\n'.join(bb_parts) + '\n  }\n};'
        )
        prev_month_label = extract_prev_month(prev_ro_df)
        log(f'   전월 기준: {prev_month_label}')
    else:
        js_prev_raw = 'const PREV_RAW = [];'
        js_prev_ro = 'const PREV_RO = {\n  total:0,invoiced:0,cancelled:0,incomplete:0,\n  byBranch:{}\n};'
        prev_month_label = '전월'

    with open(TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()

    html = re.sub(r'const RAW = \[.*?\];', js_raw, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const PREV_RAW = \[.*?\];', js_prev_raw, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const RO_STATUS = \{.*?\n\};', js_ro, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const PREV_RO = \{.*?\n\};', js_prev_ro, html, count=1, flags=re.DOTALL)
    html = re.sub(r'const PREV_MONTH_LABEL = ".*?";', f'const PREV_MONTH_LABEL = "{prev_month_label}";', html, count=1)

    if has_year and found_months:
        js_trend = trend_to_js(stype_monthly)
        html = re.sub(r'const TREND_DATA = \{.*?\n\};', js_trend, html, count=1, flags=re.DOTALL)
        trend_year = datetime.now().year - 1
        html = re.sub(r'\d{4}년 서비스타입별 총서비스 L/T 월간 변화', f'{trend_year}년 서비스타입별 총서비스 L/T 월간 변화', html)

    # Build timestamp (KST = UTC+9)
    from datetime import timedelta, timezone
    kst = timezone(timedelta(hours=9))
    build_time = datetime.now(kst).strftime('%Y-%m-%d %H:%M')

    # Handle both old and new template format
    html = html.replace('BUILD_TIMESTAMP', build_time)
    html = re.sub(
        r'<span>\d{4}-\d{2}-\d{2} \d{2}:\d{2} [^<]*</span>',
        f'<span>{build_time} 업데이트</span>',
        html, count=1
    )

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # GitHub Pages용 index.html + 로컬용 대시보드.html 둘 다 생성
    index_path = os.path.join(OUTPUT_DIR, 'index.html')
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(html)

    dash_path = os.path.join(OUTPUT_DIR, '서비스_리드타임_대시보드.html')
    with open(dash_path, 'w', encoding='utf-8') as f:
        f.write(html)

    log(f'\n[OK] 대시보드 생성 완료!')
    log(f'   - {index_path}')
    log(f'   - {dash_path}')
    log('=' * 55)
    return True


if __name__ == '__main__':
    try:
        success = build()
        if not success:
            log('\n[HELP] 올바른 폴더 구조:')
            log('  build.py 와 같은 폴더에:')
            log('  +-- template.html')
            log('  +-- data/')
            log('      +-- 현월/ (서비스_리드타임_현황.xlsx, RoReport.xlsx)')
            log('      +-- 전월/ (서비스_리드타임_현황.xlsx, RoReport.xlsx)')
            log('      +-- 전년/01~12/ (월별 서비스_리드타임_현황.xlsx)')
            if not CI_MODE:
                sys.exit(1)
    except Exception as e:
        log(f'\n[ERROR] 빌드 실패: {e}')
        import traceback
        traceback.print_exc()
        if not CI_MODE:
            input('\n엔터를 누르면 종료합니다...')
        sys.exit(1)

    if not CI_MODE:
        input('\n엔터를 누르면 종료합니다...')
