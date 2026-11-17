#!/usr/bin/env python3
"""Build a small Moody's-like demo: equations + baseline (CSV) + alternative (XLSX).

The alternative scenario: oil price +30% shock fading out, policy rate +100bp,
a judgmental cut to consumer spending (add-factor) in 2026, and a California
house-price adjustment. Uses scenario_compare's own parser to simulate.
"""
import math
import os
import random
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import scenario_compare as sc  # noqa: E402

N_GEO = int(sys.argv[1]) if len(sys.argv) > 1 else 50
N_FILL = int(sys.argv[2]) if len(sys.argv) > 2 else 800
OUT = sys.argv[3] if len(sys.argv) > 3 else 'demo'
STYLE = sys.argv[4] if len(sys.argv) > 4 else 'dot'   # 'dot': FGDP$.IUSA   'underscore': FGDP$_IUSA


def NM(n):
    return n.replace('.', '_') if STYLE == 'underscore' else n


def EQ(txt):
    if STYLE != 'underscore':
        return txt
    import re as _re
    return _re.sub(r'([A-Za-z$]+[A-Za-z0-9$]*)\.(I[A-Z0-9]+)', r'\1_\2', txt)
os.makedirs(OUT, exist_ok=True)
random.seed(7)

STATES = ('AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY '
          'NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY').split()
geos = ['I' + s for s in STATES][:N_GEO]
i = 0
while len(geos) < N_GEO:
    geos.append('IX%04d' % i)
    i += 1

periods = ['%dQ%d' % (y, q) for y in range(2012, 2031) for q in range(1, 5)]
T = len(periods)
F0 = periods.index('2026Q1')

NAT = [  # (equation text, solve kind)
    ("DLOG(FCPI.IUSA) = 0.005 + 0.02*DLOG(FOIL) + 0.3*DLOG(FCPI.IUSA(-1))", 'dlog'),
    ("FRGT10Y.IUSA = 0.8 + 0.6*FRFED + 0.15*@PCY(FCPI.IUSA)", 'level'),
    ("DLOG(FYPDPI$.IUSA) = 0.002 + 0.5*DLOG(FGDP$.IUSA(-1)) - 0.3*DLOG(FCPI.IUSA)", 'dlog'),
    ("DLOG(FC$.IUSA) = 0.001 + 0.7*DLOG(FYPDPI$.IUSA) - 0.002*D(FRGT10Y.IUSA)", 'dlog'),
    ("DLOG(FIR$.IUSA) = 0.002 - 0.01*D(FRGT10Y.IUSA) + 0.6*DLOG(FGDP$.IUSA(-1))", 'dlog'),
    ("FNX$.IUSA = -300 - 1.5*FOIL", 'level'),
    ("@IDENTITY FGDP$.IUSA = FC$.IUSA + FIR$.IUSA + FGOV$.IUSA + FNX$.IUSA", 'level'),
    ("FLBR.IUSA = FLBR.IUSA(-1) - 25*(DLOG(FGDP$.IUSA) - 0.005)", 'level'),
    ("DLOG(FHPI.IUSA) = 0.004 + 0.5*DLOG(FYPDPI$.IUSA) - 0.005*D(FRGT10Y.IUSA)", 'dlog'),
]
GEO = [
    ("DLOG(FYPDPI$.{g}) = 0.0005 + 1.1*DLOG(FYPDPI$.IUSA)", 'dlog'),
    ("DLOG(FEMP.{g}) = -0.001 + 0.8*DLOG(FGDP$.IUSA) + 0.3*DLOG(FPOP.{g})", 'dlog'),
    ("FLBR.{g} = FLBR.{g}(-1) + 0.9*D(FLBR.IUSA)", 'level'),
    ("DLOG(FHPI.{g}) = 0.002 + 0.9*DLOG(FHPI.IUSA) + 0.2*DLOG(FYPDPI$.{g})", 'dlog'),
    ("@IDENTITY FYPC$.{g} = FYPDPI$.{g} / FPOP.{g}", 'level'),
    ("LOG(FRET$.{g}/FCPI.IUSA) = 0.3 + 0.9*LOG(FYPDPI$.{g}/FCPI.IUSA)", 'logratio'),
]
DESC = {'FCPI': 'Consumer price index', 'FRGT10Y': '10-year Treasury yield, %', 'FYPDPI$': 'Real disposable income',
        'FC$': 'Real consumer spending', 'FIR$': 'Real business investment', 'FNX$': 'Real net exports',
        'FGDP$': 'Real GDP', 'FLBR': 'Unemployment rate, %', 'FHPI': 'House price index', 'FOIL': 'Oil price, WTI $/bbl',
        'FRFED': 'Fed funds rate, %', 'FGOV$': 'Real government spending', 'FPOP': 'Population', 'FEMP': 'Employment',
        'FYPC$': 'Real income per capita', 'FRET$': 'Retail sales'}

eqs = []
for txt, kind in NAT:
    eqs.append((txt, kind))
for g in geos:
    for txt, kind in GEO:
        eqs.append((txt.format(g=g), kind))
parsed = []
for txt, kind in eqs:
    e = sc.parse_equation_line(txt, 'demo')
    assert e and e.evaluable, txt
    parsed.append((e, kind))

nan = float('nan')


def init_exog(shock):
    d = {}
    oil, ff = [], []
    for t in range(T):
        o = 70 + 8 * math.sin(t / 5.0)
        f = 2.0 + 1.5 * math.sin(t / 9.0)
        if shock and t >= F0:
            k = t - F0
            o *= 1 + 0.30 * (0.85 ** k)
            f += 1.0 if k < 12 else max(0, 1.0 - 0.25 * (k - 11))
        oil.append(o)
        ff.append(f)
    d['FOIL'] = oil
    d['FRFED'] = ff
    d['FGOV$.IUSA'] = [3500 * 1.004 ** t for t in range(T)]
    for g in geos:
        base = random.Random(g).uniform(0.5, 40) * 1e3
        d['FPOP.' + g] = [base * 1.0015 ** t for t in range(T)]
    for j in range(N_FILL):
        r = random.Random(j)
        b0 = r.uniform(10, 1000)
        d['FX%04d.IUSA' % j] = [b0 * (1 + 0.003 * t) + r.gauss(0, 1) for t in range(T)]
    return d


START = {'FCPI': 230, 'FRGT10Y': 2.5, 'FYPDPI$': 12000, 'FC$': 11000, 'FIR$': 2800, 'FNX$': -400, 'FGDP$': 17000,
         'FLBR': 5.0, 'FHPI': 180, 'FEMP': 2000, 'FRET$': 400}


def simulate(shock):
    d = init_exog(shock)
    for e, _ in parsed:
        d.setdefault(e.target, [nan] * T)
    rnd = random.Random(11)
    resid = {e.target: [rnd.gauss(0, 0.002) if k in ('dlog',) else rnd.gauss(0, 0.02) for _ in range(T)]
             for e, k in parsed}
    for e, k in parsed:
        if k in ('level', 'logratio') and 'IDENTITY' in e.text.upper():
            resid[e.target] = [0.0] * T
        if k == 'logratio':
            resid[e.target] = [rnd.gauss(0, 0.002) for _ in range(T)]
    if shock:  # judgmental adjustments
        for t in range(F0, F0 + 4):
            resid['FC$.IUSA'][t] -= 0.004
        if 'FHPI.ICA' in resid:
            for t in range(F0 + 1, F0 + 3):
                resid['FHPI.ICA'][t] -= 0.01

    def get(n):
        return d[n]
    for t in range(T):
        for e, k in parsed:
            y = d[e.target]
            if t < 6:
                stem = e.target.split('.')[0]
                y[t] = START.get(stem, 100.0) * (1 + 0.01 * t)
                if e.target.startswith('FYPC$'):
                    y[t] = d['FYPDPI$.' + e.target.split('.')[1]][t] / d['FPOP.' + e.target.split('.')[1]][t]
                continue
            r = sc.evaluate(e.rhs, get, t, t, 4)[0]
            u = resid[e.target][t]
            if k == 'dlog':
                y[t] = y[t - 1] * math.exp(r + u)
            elif k == 'level':
                y[t] = r + u
            elif k == 'logratio':
                y[t] = d['FCPI.IUSA'][t] * math.exp(r + u)
    return d


A = simulate(False)
B = simulate(True)
# add-factor series (Moody's style <name>_A): zero in the baseline, the judgmental adjustment in the alternative
for nm, rng, v in [('FC$.IUSA', range(F0, F0 + 4), -0.004), ('FHPI.ICA', range(F0 + 1, F0 + 3), -0.01)]:
    if nm in A:
        A[nm + '_A'] = [0.0] * T
        B[nm + '_A'] = [v if t in rng else 0.0 for t in range(T)]
names = sorted(A)


def desc(n):
    stem, _, geo = n.partition('.')
    if n.endswith('_A'):
        return 'Add-factor for ' + n[:-2]
    base = DESC.get(stem, 'Filler series %s' % stem)
    return '%s, %s' % (base, geo or 'US') if geo else base


# Scenario A: CSV, series in rows (Mnemonic, Description, 2012Q1, ...)
with open(os.path.join(OUT, 'baseline.csv'), 'w') as f:
    f.write('Mnemonic,Description,' + ','.join(periods) + '\n')
    for n in names:
        f.write('%s,"%s",' % (NM(n), desc(n)) + ','.join('%.6f' % v for v in A[n]) + '\n')


# Scenario B: XLSX, Data Buffet style (description + mnemonic rows, dates down column A)
def xlsx(path, rows):
    def cell(r, c, v):
        col = ''
        c += 1
        while c:
            c, rem = divmod(c - 1, 26)
            col = chr(65 + rem) + col
        ref = '%s%d' % (col, r + 1)
        if isinstance(v, (int, float)):
            return '<c r="%s"><v>%r</v></c>' % (ref, v)
        v = str(v).replace('&', '&amp;').replace('<', '&lt;')
        return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, v)
    body = ''.join('<row r="%d">%s</row>' % (r + 1, ''.join(cell(r, c, v) for c, v in enumerate(row)))
                   for r, row in enumerate(rows))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>%s</sheetData></worksheet>' % body)


rows = [['Description:'] + [desc(n) for n in names], ['Mnemonic:'] + [NM(n) for n in names], ['Source:'] + ['Moody\'s demo'] * len(names)]
for t, p in enumerate(periods):
    rows.append([p] + [round(B[n][t], 6) for n in names])
xlsx(os.path.join(OUT, 'alt_scenario.xlsx'), rows)

with open(os.path.join(OUT, 'equations.txt'), 'w') as f:
    f.write("' Demo model in EViews syntax (substituted coefficients)\n:EQ_IGNORED\n")
    for txt, _ in eqs:
        if txt.startswith('DLOG(FC$.IUSA)'):  # split over two lines on purpose
            l, r = txt.split(' - 0.002')
            f.write(EQ(l) + '\n   - 0.002' + EQ(r) + '\n')
        else:
            f.write(EQ(txt) + '\n')
    f.write(EQ("eq_fc.ls dlog(fc$.iusa) c dlog(fypdpi$.iusa) d(frgt10y.iusa)\n"))  # spec duplicate
    f.write(EQ("@ADD(V) FC$.IUSA FC$.IUSA_A\n"))
    f.write("smpl 2026q1 2030q4\n")
print('demo written to', OUT, '-', len(names), 'series,', len(eqs), 'equations')
