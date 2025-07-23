import io
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict, deque

XLSX_FILE = 'juicereel_jun_25.xlsx'

SHEETS = {
    'All': 'worksheets/sheet1.xml',
    'Live': 'worksheets/sheet2.xml',
}

# load shared strings
with zipfile.ZipFile(XLSX_FILE) as zf:
    root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    SHARED = [t.text or '' for t in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')]

def iter_rows(zf, sheet):
    data = zf.read(f'xl/{SHEETS[sheet]}')
    for _, elem in ET.iterparse(io.BytesIO(data), events=('end',)):
        if elem.tag.endswith('row'):
            row = []
            for c in elem.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                val = v.text if v is not None else ''
                if c.get('t') == 's' and val:
                    val = SHARED[int(val)]
                row.append(val)
            yield row
            elem.clear()

def monthly_weighted_avg(zf, sheet):
    rows = iter_rows(zf, sheet)
    header = next(rows)
    idx = {h: i for i, h in enumerate(header)}
    stats = defaultdict(lambda: {'ws': 0.0, 'cnt': 0.0})
    for row in rows:
        ym = (int(row[idx['bet_year']]), int(row[idx['bet_month']]))
        count = float(row[idx['count_of_bets']])
        val = row[idx['avg_leg_count_inclusiveofstraightbets']]
        avg = float(val) if val else 0.0
        stats[ym]['ws'] += avg * count
        stats[ym]['cnt'] += count
    return {ym: d['ws'] / d['cnt'] if d['cnt'] else 0.0 for ym, d in stats.items()}

def trailing_average(series, window=12):
    q = deque()
    total = 0.0
    result = {}
    for ym in sorted(series):
        avg = series[ym]
        q.append(avg)
        total += avg
        if len(q) > window:
            total -= q.popleft()
        result[ym] = total / len(q)
    return result

def top_operators(zf, header):
    idx = {h: i for i, h in enumerate(header)}
    totals = defaultdict(float)
    for row in iter_rows(zf, 'All'):
        if row == header:
            continue
        name = row[idx['name']]
        handle = float(row[idx['total_bet_handle']]) if row[idx['total_bet_handle']] else 0.0
        totals[name] += handle
    top = sorted(totals.items(), key=lambda x: x[1], reverse=True)[:10]
    return [n for n, _ in top]

def operator_trailing(zf, header, operators):
    idx = {h: i for i, h in enumerate(header)}
    stats = defaultdict(lambda: defaultdict(lambda: {'ws': 0.0, 'cnt': 0.0}))
    rows = iter_rows(zf, 'All')
    next(rows)  # skip header again
    for row in rows:
        name = row[idx['name']]
        if name not in operators:
            continue
        ym = (int(row[idx['bet_year']]), int(row[idx['bet_month']]))
        count = float(row[idx['count_of_bets']])
        val = row[idx['avg_leg_count_inclusiveofstraightbets']]
        avg = float(val) if val else 0.0
        stats[name][ym]['ws'] += avg * count
        stats[name][ym]['cnt'] += count
    trailing = {}
    for op in operators:
        series = {ym: d['ws'] / d['cnt'] for ym, d in stats[op].items()}
        trailing[op] = trailing_average(series)
    return trailing

if __name__ == '__main__':
    with zipfile.ZipFile(XLSX_FILE) as zf:
        all_month = monthly_weighted_avg(zf, 'All')
        live_month = monthly_weighted_avg(zf, 'Live')
        nonlive_month = {ym: all_month[ym] - live_month.get(ym, 0.0) for ym in all_month}

        all_trail = trailing_average(all_month)
        live_trail = trailing_average(live_month)
        non_trail = trailing_average(nonlive_month)

        print('Trailing 12-month averages (All vs Live vs Non-live) - last 3 months:')
        for ym in sorted(all_trail)[-3:]:
            print(ym, {
                'all': round(all_trail[ym], 3),
                'live': round(live_trail.get(ym, 0.0), 3),
                'non_live': round(non_trail.get(ym, 0.0), 3),
            })

        rows = iter_rows(zf, 'All')
        header = next(rows)
        top_ops = top_operators(zf, header)
        print('\nTop operators:', ', '.join(top_ops))
        op_trailing = operator_trailing(zf, header, top_ops)
        for op in top_ops:
            series = op_trailing[op]
            months = sorted(series)[-3:]
            print(f"\n{op} trailing averages:")
            for ym in months:
                print(ym, round(series[ym], 3))

