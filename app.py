from flask import Flask, render_template, request, jsonify
import openpyxl
import os
import re
from datetime import datetime

app = Flask(__name__)

EXCEL_PATH = os.environ.get(
    'EXCEL_PATH',
    r'C:\Users\Hong\Downloads\★★제휴정산_입금내역★★.xlsx'
)

_cache = {}

# 파트너명 → 시트명 매핑
PARTNER_SHEETS = {
    'KG이니시스':       '1.KG이니시스',
    'GTF':             '3.GTF',
    '리터놀':           '4.리터놀',
    '버클 (선발행)':    '5. 버클',
    '비플':             '6.비플',
    '야마토':           '7.야마토(YSM)',
    '채널코퍼레이션':   '11.채널코퍼레이션',
    '체큐리티(나이스)': '12.체큐리티(나이스)',
    '카카오모빌리티':   '13.카카오모빌리티',
    '토스페이먼츠':     '14.토스페이먼츠',
    '패스트박스':       '15.패스트박스',
    '스토어X':          '17.스토어X',
    '두발히어로':       '19.두발히어로',
    '엑심베이':         '20.엑심베이',
}


def parse_month_label(raw):
    if not raw:
        return ''
    m = re.search(r'(\d{4})\.(\d+)월', str(raw))
    if m:
        return f"{m.group(2)}월"
    m2 = re.search(r'(\d+)월', str(raw))
    if m2:
        return f"{m2.group(1)}월"
    return str(raw)[:10]


def load_status_data():
    if 'status' in _cache:
        return _cache['status']

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb['입금여부']

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {'months': [], 'partners': []}

    header = rows[0]
    months = [parse_month_label(header[i]) for i in range(1, 13)]

    partners = []
    for row in rows[1:]:
        if not row[0] or not str(row[0]).strip():
            continue
        name = str(row[0]).strip()
        note = str(row[13]).strip() if len(row) > 13 and row[13] else ''
        note = '' if note == 'None' else note

        statuses = []
        for i in range(1, 13):
            val = row[i] if i < len(row) else None
            v = str(val).strip() if val is not None else ''
            if not v or v == 'None':
                statuses.append({'type': 'pending', 'label': ''})
            elif '✔' in v:
                statuses.append({'type': 'paid', 'label': '✔'})
            elif v.lower() == 'x':
                statuses.append({'type': 'refused', 'label': '✕'})
            else:
                statuses.append({'type': 'na', 'label': v[:8]})

        partners.append({
            'name': name,
            'statuses': statuses,
            'note': note,
            'has_sheet': name in PARTNER_SHEETS,
        })

    result = {'months': months, 'partners': partners}
    _cache['status'] = result
    return result


def load_partner_detail(partner_name):
    sheet_name = PARTNER_SHEETS.get(partner_name)
    if not sheet_name:
        return None

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    if sheet_name not in wb.sheetnames:
        return None

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    records = []
    for row in rows:
        if not any(row):
            continue
        month_val = str(row[0]).strip() if row[0] else ''
        if not re.match(r'^\d+월$', month_val):
            continue

        # 컬럼 2 이후에서 숫자 찾기
        amounts = []
        texts = []
        for cell in row[1:]:
            if cell is None:
                continue
            try:
                amt = float(cell)
                if amt > 0:
                    amounts.append(int(amt))
            except (TypeError, ValueError):
                s = str(cell).strip()
                if s and s != 'None':
                    texts.append(s)

        total = sum(amounts) if amounts else 0
        records.append({
            'month': month_val,
            'amounts': amounts,
            'texts': texts,
            'total': total,
        })

    return records


@app.route('/')
def index():
    data = load_status_data()
    partners = data['partners']
    months = data['months']

    # 현재 월 기준 통계 (0-indexed)
    now_month = datetime.now().month
    cur_idx = min(now_month - 1, 11)

    stats = []
    for i, m in enumerate(months):
        paid = sum(1 for p in partners if p['statuses'][i]['type'] == 'paid')
        refused = sum(1 for p in partners if p['statuses'][i]['type'] == 'refused')
        pending = sum(1 for p in partners if p['statuses'][i]['type'] == 'pending')
        na = sum(1 for p in partners if p['statuses'][i]['type'] == 'na')
        stats.append({'month': m, 'paid': paid, 'refused': refused,
                      'pending': pending, 'na': na})

    return render_template('index.html',
                           months=months,
                           partners=partners,
                           stats=stats,
                           cur_idx=cur_idx,
                           total=len(partners))


@app.route('/partner/<path:name>')
def partner_detail(name):
    data = load_status_data()
    partner = next((p for p in data['partners'] if p['name'] == name), None)
    if not partner:
        return "파트너를 찾을 수 없습니다.", 404

    records = load_partner_detail(name) or []
    grand_total = sum(r['total'] for r in records)
    return render_template('partner.html',
                           partner=partner,
                           months=data['months'],
                           records=records,
                           grand_total=grand_total)


@app.route('/api/reload')
def reload_cache():
    _cache.clear()
    return jsonify({'ok': True, 'message': '캐시가 초기화되었습니다.'})


if __name__ == '__main__':
    app.run(debug=True, port=5001)
