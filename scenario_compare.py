#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scenario_compare.py - make sense of two Moody's (EViews-style) macro scenarios.

Standard library only (Python 3.8+). No pip installs, no openpyxl/pandas.

Commands
  report    interactive HTML report: key changes, where the scenario starts,
            driver decomposition for each variable, path tracing X -> Y
  top       print the biggest changes (optionally save all stats to CSV)
  explain   contribution of each input (and the add-factor) to the gap of VAR
  addfactors  formula vs add-factor share of the gap, for every equation
  impact    can X affect Y, and by how much? (re-solves the equations with only X changed)
  attribute split the gap of Y across every scenario shock upstream of it
  path      the channels through which X reaches Y
  serve     the report in your browser, plus exact impact / attribution on demand
  info      what was loaded (layout detected, periods, equations parsed)

Examples
  python scenario_compare.py report  --a base.csv --b s3.csv --eq model.txt -o compare.html
  python scenario_compare.py top     --a base.csv --b s3.csv -n 40
  python scenario_compare.py explain FGDP$.IUSA --a base.csv --b s3.csv --eq model.txt
  python scenario_compare.py path FOIL FHPI.IOH --a base.csv --b s3.csv --eq model.txt

Run any command with -h for all options.
"""
import argparse
import csv
import datetime
import json
import math
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from array import array
from collections import Counter, defaultdict, deque

NAN = float('nan')
NA_STRINGS = {'NA', 'N/A', '#N/A', 'NAN', 'ND', 'NULL', 'NONE', '-', '.', '..',
              '#VALUE!', '#DIV/0!', '#REF!', '#NUM!'}

try:
    csv.field_size_limit(min(sys.maxsize, 2 ** 31 - 1))
except Exception:  # pragma: no cover
    pass


def _num(cell):
    """Parse a numeric cell; missing / text -> NaN."""
    if cell is None:
        return NAN
    if isinstance(cell, (int, float)):
        v = float(cell)
        return v if math.isfinite(v) else NAN
    t = str(cell).strip()
    if not t or t.upper() in NA_STRINGS:
        return NAN
    t = t.replace(',', '')
    try:
        v = float(t)
    except ValueError:
        return NAN
    return v if math.isfinite(v) else NAN


def _clean(c):
    return '' if c is None else str(c).strip()


def _isnan(v):
    return v != v


# --------------------------------------------------------------------------
# Periods
# --------------------------------------------------------------------------
MONTHS = {m: i + 1 for i, m in enumerate(
    'JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}
_RE_Q1 = re.compile(r'^(\d{4})\s*[-\s]?\s*[Qq]\s*([1-4])$')
_RE_Q2 = re.compile(r'^[Qq]([1-4])[\s\-/:]*(\d{4})$')
_RE_M1 = re.compile(r'^(\d{4})\s*[Mm]\s*(\d{1,2})$')
_RE_YM = re.compile(r'^(\d{4})-(\d{1,2})$')
_RE_COLON = re.compile(r'^(\d{4}):(\d{1,2})$')
_RE_ISO = re.compile(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T].*)?$')
_RE_US = re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:\s.*)?$')
_RE_Y = re.compile(r'^(\d{4})\s*[Aa]?$')
_RE_MON = re.compile(r'^([A-Za-z]{3})[a-z]*[\s\-]+(\d{2}|\d{4})$')


def _date_key(y, mo, d):
    try:
        datetime.date(y, mo, d)
    except ValueError:
        return None
    return ((y, 5, mo * 100 + d), '%04d-%02d-%02d' % (y, mo, d))


def parse_period(cell, allow_serial=False):
    """Return (sortkey, label) if the cell looks like a date/period, else None."""
    if cell is None:
        return None
    t = str(cell).strip()
    if not t or len(t) > 30:
        return None
    m = _RE_Q1.match(t)
    if m:
        y, q = int(m.group(1)), int(m.group(2))
        return ((y, 2, q), '%dQ%d' % (y, q))
    m = _RE_Q2.match(t)
    if m:
        q, y = int(m.group(1)), int(m.group(2))
        return ((y, 2, q), '%dQ%d' % (y, q))
    m = _RE_M1.match(t) or _RE_YM.match(t)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return ((y, 3, mo), '%dM%02d' % (y, mo))
        return None
    m = _RE_COLON.match(t)
    if m:
        y, s = int(m.group(1)), int(m.group(2))
        if 1 <= s <= 12:
            return ((y, 4, s), '%d:%d' % (y, s))
        return None
    m = _RE_ISO.match(t)
    if m:
        return _date_key(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = _RE_US.match(t)
    if m:
        mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000 if y < 50 else 1900
        return _date_key(y, mo, d)
    m = _RE_MON.match(t)
    if m and m.group(1).upper() in MONTHS:
        y = int(m.group(2))
        if y < 100:
            y += 2000 if y < 50 else 1900
        mo = MONTHS[m.group(1).upper()]
        return ((y, 3, mo), '%dM%02d' % (y, mo))
    m = _RE_Y.match(t)
    if m:
        y = int(m.group(1))
        if 1800 <= y <= 2200:
            return ((y, 1, 0), str(y))
        return None
    if allow_serial:  # Excel date serial number (xlsx headers)
        try:
            v = float(t)
        except ValueError:
            return None
        if 15000 < v < 80000 and abs(v - round(v)) < 1e-9:
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(round(v)))
            return _date_key(d.year, d.month, d.day)
    return None


def infer_ppy(keys):
    """Periods per year, used by @PCHY/@PCY."""
    if not keys:
        return 4
    kind = Counter(k[1] for k in keys).most_common(1)[0][0]
    if kind == 2:
        return 4
    if kind == 3:
        return 12
    if kind == 1:
        return 1
    if kind == 4:
        return 12 if max(k[2] for k in keys if k[1] == 4) > 4 else 4
    ds = sorted(datetime.date(k[0], k[2] // 100, k[2] % 100) for k in keys if k[1] == 5)
    if len(ds) < 2:
        return 4
    gaps = sorted((b - a).days for a, b in zip(ds, ds[1:]))
    g = gaps[len(gaps) // 2]
    if g < 45:
        return 12
    if g < 200:
        return 4
    return 1


# --------------------------------------------------------------------------
# Reading files (CSV / TSV / XLSX) with the standard library
# --------------------------------------------------------------------------
_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx(path, sheet=None):
    """Minimal .xlsx reader (zipfile + xml). Returns list of rows of strings."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if 'xl/sharedStrings.xml' in names:
            for _, el in ET.iterparse(z.open('xl/sharedStrings.xml')):
                if el.tag == _NS + 'si':
                    shared.append(''.join(t.text or '' for t in el.iter(_NS + 't')))
                    el.clear()
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'): r.get('Target') for r in rels}
        sheets = []
        for s in wb.iter(_NS + 'sheet'):
            tgt = (relmap.get(s.get(_RNS + 'id')) or '').lstrip('/')
            if not tgt.startswith('xl/'):
                tgt = 'xl/' + tgt
            sheets.append((s.get('name'), tgt))
        if not sheets:
            raise ValueError('%s: no worksheets found' % path)
        target = sheets[0][1]
        if sheet is not None:
            found = None
            for i, (nm, tg) in enumerate(sheets):
                if nm.lower() == str(sheet).lower() or str(i + 1) == str(sheet):
                    found = tg
            if not found:
                raise ValueError('%s: sheet %r not found; sheets are %s'
                                 % (path, sheet, [s[0] for s in sheets]))
            target = found
        grid = []
        for _, el in ET.iterparse(z.open(target)):
            if el.tag != _NS + 'row':
                continue
            row = []
            for c in el.findall(_NS + 'c'):
                ref = c.get('r')
                j = _col_index(ref) if ref else len(row)
                t = c.get('t')
                val = ''
                if t == 's':
                    v = c.find(_NS + 'v')
                    if v is not None and v.text is not None:
                        val = shared[int(v.text)]
                elif t == 'inlineStr':
                    is_ = c.find(_NS + 'is')
                    if is_ is not None:
                        val = ''.join(x.text or '' for x in is_.iter(_NS + 't'))
                else:
                    v = c.find(_NS + 'v')
                    if v is not None and v.text is not None:
                        val = v.text
                        if t not in ('str', 'e', 'b'):
                            try:
                                f = float(val)
                                val = str(int(f)) if f.is_integer() and abs(f) < 1e15 else repr(f)
                            except ValueError:
                                pass
                if j >= len(row):
                    row.extend([''] * (j - len(row) + 1))
                row[j] = val
            r = el.get('r')
            ri = int(r) - 1 if r else len(grid)
            while len(grid) < ri:
                grid.append([])
            grid.append(row)
            el.clear()
    return grid


def read_grid(path, sheet=None):
    """Return (rows, is_xlsx)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        return read_xlsx(path, sheet), True
    if ext == '.xls':
        raise SystemExit('%s: old .xls format is not supported - save it as .xlsx or .csv' % path)
    with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        sample = f.read(65536)
        f.seek(0)
        delim = ','
        if ext in ('.tsv', '.tab'):
            delim = '\t'
        else:
            try:
                delim = csv.Sniffer().sniff(sample, delimiters=',\t;|').delimiter
            except csv.Error:
                pass
        return [row for row in csv.reader(f, delimiter=delim)], False


# --------------------------------------------------------------------------
# Scenario data
# --------------------------------------------------------------------------
NAME_HDRS = {'mnemonic', 'mnemonics', 'series', 'variable', 'name', 'code', 'ticker',
             'id', 'series id', 'series_id', 'mnem', 'series name', 'var'}


def _is_desc_hdr(h):
    h = h.lower().rstrip(':').strip()
    return ('desc' in h or 'title' in h or h in ('concept', 'label', 'long name', 'longname'))


class Scenario:
    def __init__(self, label):
        self.label = label
        self.blocks = []        # list of lists of period sortkeys
        self.series = {}        # NAME -> (block id, array of floats)
        self.desc = {}
        self.keys = {}          # sortkey -> label
        self.notes = []
        self.dups = 0

    def put(self, name, bid, arr, desc=''):
        name = name.strip().upper()
        if not name:
            return
        if name in self.series:
            self.dups += 1
        self.series[name] = (bid, arr)
        if desc and name not in self.desc:
            self.desc[name] = desc


def _pick_name_col(grid, start, candidates, hdr):
    for j in candidates:
        if j < len(hdr) and hdr[j] in NAME_HDRS:
            return j
    best, bestscore = None, -1
    sample = grid[start:start + 60]
    for j in candidates:
        score = 0
        for r in sample:
            c = _clean(r[j]) if j < len(r) else ''
            if c and _isnan(_num(c)):
                score += 1
        if score > bestscore:
            best, bestscore = j, score
    return best


def _increasing(cands):
    """Keep period cells that form an increasing run (rejects numbers that merely look like dates)."""
    out, last = [], None
    for j, p in cands:
        if not p:
            continue
        if last is None or p[0] > last:
            out.append((j, p))
            last = p[0]
    # a real date axis has (almost) no out-of-order cells
    n = sum(1 for _, p in cands if p)
    return out if n and len(out) >= 0.9 * n else []


def _ingest(sc, grid, serial, path):
    nonempty = [i for i, r in enumerate(grid) if any(_clean(c) for c in r)]
    if not nonempty:
        raise SystemExit('%s: file is empty' % path)
    h = nonempty[0]
    hdr = [_clean(c).lower() for c in grid[h]]

    def find(cands):
        for j, c in enumerate(hdr):
            if c in cands:
                return j
        return -1
    vcol = find({'value', 'val', 'data', 'obs_value', 'amount'})
    dcol = find({'date', 'period', 'time', 'obs', 'datetime', 'quarter', 'month', 'year'})
    scol = find(NAME_HDRS)
    if vcol >= 0 and dcol >= 0 and scol >= 0 and len([c for c in hdr if c]) <= 10:
        return _ingest_long(sc, grid, h, scol, dcol, vcol, serial, path)

    rows_hdr = None
    for i in nonempty[:25]:
        cols = _increasing([(j, parse_period(c, serial)) for j, c in enumerate(grid[i])])
        if len(cols) >= 3:
            rows_hdr = (i, cols)
            break
    cols_best = None
    for c in range(4):
        idx = _increasing([(i, parse_period(r[c], serial) if c < len(r) else None)
                           for i, r in enumerate(grid)])
        if len(idx) >= 3 and (cols_best is None or len(idx) > len(cols_best[1])):
            cols_best = (c, idx)
    nr = len(rows_hdr[1]) if rows_hdr else 0
    nc = len(cols_best[1]) if cols_best else 0
    if nr == 0 and nc == 0:
        raise SystemExit('%s: could not find a row or column of dates/periods '
                         '(e.g. 2026Q1, 2026-03-31, 2026:1).' % path)
    if nr >= nc:
        _ingest_rows(sc, grid, rows_hdr[0], rows_hdr[1], path)
    else:
        _ingest_cols(sc, grid, cols_best[0], cols_best[1], path)


def _ingest_rows(sc, grid, i, cols, path):
    for _, p in cols:
        sc.keys[p[0]] = p[1]
    datecols = set(j for j, _ in cols)
    hdr = [_clean(c).lower() for c in grid[i]]
    width = max(len(r) for r in grid[i:i + 50])
    nondate = [j for j in range(width) if j not in datecols]
    if not nondate:
        raise SystemExit('%s: no column with series names found' % path)
    namecol = _pick_name_col(grid, i + 1, nondate, hdr)
    desccol = next((j for j in nondate if j < len(hdr) and _is_desc_hdr(hdr[j])), None)
    bid = len(sc.blocks)
    sc.blocks.append([p[0] for _, p in cols])
    n = 0
    jlist = [j for j, _ in cols]
    for r in grid[i + 1:]:
        name = _clean(r[namecol]) if namecol < len(r) else ''
        if not name:
            continue
        L = len(r)
        arr = array('d', (_num(r[j]) if j < L else NAN for j in jlist))
        if all(_isnan(v) for v in arr):
            continue
        desc = _clean(r[desccol]) if desccol is not None and desccol < L else ''
        sc.put(name, bid, arr, desc)
        n += 1
    sc.notes.append('%s: series in rows, %d series x %d periods (%s .. %s), names from column %d'
                    % (os.path.basename(path), n, len(cols), cols[0][1][1], cols[-1][1][1], namecol + 1))


def _ingest_cols(sc, grid, c, idx, path):
    for _, p in idx:
        sc.keys[p[0]] = p[1]
    first = idx[0][0]
    pre = [i for i in range(first) if any(_clean(x) for x in grid[i])]
    if not pre:
        raise SystemExit('%s: dates are in column %d but there is no header row with series names'
                         % (path, c + 1))
    namerow, descrow = None, None
    for i in pre:
        lab = _clean(grid[i][c] if c < len(grid[i]) else '').lower().rstrip(':').strip()
        if not lab and c > 0:
            lab = _clean(grid[i][0]).lower().rstrip(':').strip()
        if lab in NAME_HDRS and namerow is None:
            namerow = i
        elif _is_desc_hdr(lab) and descrow is None:
            descrow = i
    if namerow is None:
        best = -1
        for i in pre:
            sc_ = sum(1 for j, x in enumerate(grid[i]) if j > c and _clean(x) and _isnan(_num(x)))
            if sc_ > best:
                best, namerow = sc_, i
    names = grid[namerow]
    descs = grid[descrow] if descrow is not None else []
    bid = len(sc.blocks)
    sc.blocks.append([p[0] for _, p in idx])
    rows = [grid[i] for i, _ in idx]
    n = 0
    for j in range(c + 1, len(names)):
        name = _clean(names[j])
        if not name:
            continue
        arr = array('d', (_num(r[j]) if j < len(r) else NAN for r in rows))
        if all(_isnan(v) for v in arr):
            continue
        sc.put(name, bid, arr, _clean(descs[j]) if j < len(descs) else '')
        n += 1
    sc.notes.append('%s: series in columns, %d series x %d periods (%s .. %s), names from row %d'
                    % (os.path.basename(path), n, len(idx), idx[0][1][1], idx[-1][1][1], namerow + 1))


def _ingest_long(sc, grid, h, scol, dcol, vcol, serial, path):
    data = defaultdict(dict)
    for r in grid[h + 1:]:
        if max(scol, dcol, vcol) >= len(r):
            continue
        p = parse_period(r[dcol], serial)
        name = _clean(r[scol])
        if not p or not name:
            continue
        sc.keys[p[0]] = p[1]
        data[name][p[0]] = _num(r[vcol])
    keys = sorted(set(k for d in data.values() for k in d))
    bid = len(sc.blocks)
    sc.blocks.append(keys)
    for name, d in data.items():
        sc.put(name, bid, array('d', (d.get(k, NAN) for k in keys)))
    sc.notes.append('%s: long format (series, date, value), %d series' % (os.path.basename(path), len(data)))


def load_scenario(paths, label, sheet=None):
    sc = Scenario(label)
    for p in paths:
        grid, serial = read_grid(p, sheet)
        _ingest(sc, grid, serial, p)
    if sc.dups:
        sc.notes.append('%d duplicate series names (last one kept)' % sc.dups)
    return sc


def load_meta(path):
    grid, _ = read_grid(path)
    hdr = [_clean(c).lower() for c in grid[0]]
    nc = next((j for j, h in enumerate(hdr) if h in NAME_HDRS), 0)
    dc = next((j for j, h in enumerate(hdr) if _is_desc_hdr(h)), 1)
    out = {}
    for r in grid[1:]:
        if nc < len(r) and dc < len(r) and _clean(r[nc]):
            out[_clean(r[nc]).upper()] = _clean(r[dc])
    return out


# --------------------------------------------------------------------------
# EViews-style equations: tokenizer, parser, evaluator
# --------------------------------------------------------------------------
class ParseError(Exception):
    pass


class EvalError(Exception):
    pass


_TOKRE = re.compile(r'''\s*(?:
    (?P<num>(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?)
  | (?P<name>@?[A-Za-z_][A-Za-z0-9_$.%#]*)
  | (?P<op>\*\*|<=|>=|<>|[-+*/^(),=<>])
)''', re.X)

FUNCS = {'LOG', 'LN', 'EXP', 'ABS', 'SQR', 'SQRT', 'D', 'DLOG'}
EXCLUDE = {'C', 'AR', 'MA', 'SAR', 'SMA', 'PDL', 'NA'}


def tokenize(s):
    toks, pos, n = [], 0, len(s)
    while pos < n:
        if s[pos].isspace():
            pos += 1
            continue
        m = _TOKRE.match(s, pos)
        if not m or m.end() == pos:
            raise ParseError('unexpected character %r' % s[pos])
        if m.group('num') is not None:
            toks.append(('num', m.group('num')))
        elif m.group('name') is not None:
            toks.append(('name', m.group('name')))
        else:
            op = m.group('op')
            toks.append(('op', '^' if op == '**' else op))
        pos = m.end()
    return toks


class _Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, v):
        if self.take()[1] != v:
            raise ParseError('expected %r' % v)

    def parse(self):
        n = self.expr()
        if self.i != len(self.t):
            raise ParseError('unexpected %r' % (self.peek()[1],))
        return n

    def expr(self):
        n = self.term()
        while self.peek()[1] in ('+', '-'):
            op = self.take()[1]
            n = ('bin', op, n, self.term())
        return n

    def term(self):
        n = self.unary()
        while self.peek()[1] in ('*', '/'):
            op = self.take()[1]
            n = ('bin', op, n, self.unary())
        return n

    def unary(self):
        if self.peek()[1] in ('-', '+'):
            op = self.take()[1]
            u = self.unary()
            return ('neg', u) if op == '-' else u
        return self.power()

    def power(self):
        b = self.atom()
        if self.peek()[1] == '^':
            self.take()
            return ('bin', '^', b, self.unary())
        return b

    def _try_lag(self):
        # NAME ( [+-] int )
        t0, t1, t2, t3 = self.peek(0), self.peek(1), self.peek(2), self.peek(3)
        if t0[1] != '(':
            return None
        if t1[0] == 'num' and t2[1] == ')' and re.fullmatch(r'\d+', t1[1]):
            self.i += 3
            return int(t1[1])
        if t1[1] in ('-', '+') and t2[0] == 'num' and t3[1] == ')' and re.fullmatch(r'\d+', t2[1]):
            self.i += 4
            return -int(t2[1]) if t1[1] == '-' else int(t2[1])
        return None

    def atom(self):
        kind, val = self.take()
        if kind == 'num':
            return ('num', float(val))
        if val == '(':
            n = self.expr()
            self.expect(')')
            return n
        if kind == 'name':
            up = val.upper()
            is_func = up in FUNCS or up.startswith('@')
            if self.peek()[1] == '(':
                if not is_func:
                    lag = self._try_lag()
                    if lag is not None:
                        return ('coef', lag) if up == 'C' else ('var', up, lag)
                self.take()
                args = []
                if self.peek()[1] != ')':
                    args.append(self.expr())
                    while self.peek()[1] == ',':
                        self.take()
                        args.append(self.expr())
                self.expect(')')
                return ('call', up, args)
            if up.startswith('@'):
                return ('call', up, [])
            return ('var', up, 0)
        raise ParseError('unexpected %r' % (val,))


def parse_expr(toks):
    if not toks:
        raise ParseError('empty expression')
    return _Parser(toks).parse()


def ast_vars(node, out=None):
    """Ordered unique variable names in an expression."""
    if out is None:
        out = []
    k = node[0]
    if k == 'var':
        if node[1] not in out and node[1] not in EXCLUDE:
            out.append(node[1])
    elif k == 'neg':
        ast_vars(node[1], out)
    elif k == 'bin':
        ast_vars(node[2], out)
        ast_vars(node[3], out)
    elif k == 'call':
        for a in node[2]:
            ast_vars(a, out)
    return out


def ast_has(node, kinds):
    k = node[0]
    if k in kinds:
        return True
    if k == 'neg':
        return ast_has(node[1], kinds)
    if k == 'bin':
        return ast_has(node[2], kinds) or ast_has(node[3], kinds)
    if k == 'call':
        return any(ast_has(a, kinds) for a in node[2])
    return False


def _const(node):
    if node[0] == 'num':
        return node[1]
    if node[0] == 'neg' and node[1][0] == 'num':
        return -node[1][1]
    raise EvalError('expected a constant argument')


def _log(v):
    return math.log(v) if v > 0 else NAN


def _exp(v):
    try:
        return math.exp(v)
    except OverflowError:
        return NAN


def _pow(x, y):
    if x != x or y != y:
        return NAN
    if x < 0 and not float(y).is_integer():
        return NAN
    try:
        return float(x ** y)
    except (OverflowError, ZeroDivisionError, ValueError):
        return NAN


def _diff(xs, k):
    return [xs[i] - xs[i - k] for i in range(k, len(xs))]


def evaluate(node, get, lo, hi, ppy):
    """Evaluate an expression for period indices lo..hi; returns list of floats (NaN = missing)."""
    k = node[0]
    if k == 'num':
        return [node[1]] * (hi - lo + 1)
    if k == 'var':
        s = get(node[1])
        a, b = lo + node[2], hi + node[2]
        n = len(s)
        if a >= 0 and b < n:
            return list(s[a:b + 1])
        return [s[t] if 0 <= t < n else NAN for t in range(a, b + 1)]
    if k == 'neg':
        return [-v for v in evaluate(node[1], get, lo, hi, ppy)]
    if k == 'bin':
        op = node[1]
        x = evaluate(node[2], get, lo, hi, ppy)
        y = evaluate(node[3], get, lo, hi, ppy)
        if op == '+':
            return [p + q for p, q in zip(x, y)]
        if op == '-':
            return [p - q for p, q in zip(x, y)]
        if op == '*':
            return [p * q for p, q in zip(x, y)]
        if op == '/':
            return [p / q if q != 0 else NAN for p, q in zip(x, y)]
        if op == '^':
            return [_pow(p, q) for p, q in zip(x, y)]
        raise EvalError('operator %s' % op)
    if k == 'coef':
        raise EvalError('equation has unsubstituted coefficient C(%d)' % node[1])
    if k == 'call':
        f = node[1].lstrip('@')
        args = node[2]
        if f in ('LOG', 'LN'):
            return [_log(v) for v in evaluate(args[0], get, lo, hi, ppy)]
        if f == 'EXP':
            return [_exp(v) for v in evaluate(args[0], get, lo, hi, ppy)]
        if f == 'ABS':
            return [abs(v) for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('SQR', 'SQRT'):
            return [math.sqrt(v) if v >= 0 else NAN for v in evaluate(args[0], get, lo, hi, ppy)]
        if f in ('D', 'DLOG'):
            order = int(_const(args[1])) if len(args) > 1 else 1
            seas = int(_const(args[2])) if len(args) > 2 else 0
            xs = evaluate(args[0], get, lo - order - seas, hi, ppy)
            if f == 'DLOG':
                xs = [_log(v) for v in xs]
            for _ in range(order):
                xs = _diff(xs, 1)
            if seas:
                xs = _diff(xs, seas)
            return xs
        if f in ('PCH', 'PC', 'PCHY', 'PCY', 'PCHA', 'PCA'):
            lag = ppy if f in ('PCHY', 'PCY') else 1
            xs = evaluate(args[0], get, lo - lag, hi, ppy)
            out = [xs[i] / xs[i - lag] - 1 if xs[i - lag] != 0 else NAN for i in range(lag, len(xs))]
            if f in ('PCHA', 'PCA'):
                out = [_pow(1 + v, ppy) - 1 for v in out]
            if f in ('PC', 'PCY', 'PCA'):
                out = [100 * v for v in out]
            return out
        if f in ('MOVAV', 'MOVSUM'):
            n = int(_const(args[1]))
            xs = evaluate(args[0], get, lo - n + 1, hi, ppy)
            out = [sum(xs[i - n + 1:i + 1]) for i in range(n - 1, len(xs))]
            return [v / n for v in out] if f == 'MOVAV' else out
        if f == 'TREND':
            return [float(t) for t in range(lo, hi + 1)]
        raise EvalError('unsupported function %s' % node[1])
    raise EvalError('bad node')


def ast_text(node):
    k = node[0]
    if k == 'num':
        v = node[1]
        return ('%d' % v) if v.is_integer() else ('%g' % v)
    if k == 'var':
        return node[1] if node[2] == 0 else '%s(%d)' % (node[1], node[2])
    if k == 'coef':
        return 'C(%d)' % node[1]
    if k == 'neg':
        return '-' + ast_text(node[1])
    if k == 'bin':
        return '%s%s%s' % (ast_text(node[2]), node[1], ast_text(node[3]))
    if k == 'call':
        return '%s(%s)' % (node[1], ','.join(ast_text(a) for a in node[2]))
    return '?'


class Equation:
    __slots__ = ('target', 'lhs', 'rhs', 'lhs_vars', 'rhs_vars', 'drivers', 'text',
                 'kind', 'evaluable', 'source', 'why')

    def __init__(self, target, drivers, text, kind, source, lhs=None, rhs=None,
                 lhs_vars=(), rhs_vars=(), evaluable=False, why=''):
        self.target = target
        self.drivers = list(drivers)
        self.text = text
        self.kind = kind
        self.source = source
        self.lhs, self.rhs = lhs, rhs
        self.lhs_vars, self.rhs_vars = list(lhs_vars), list(rhs_vars)
        self.evaluable = evaluable
        self.why = why


_NAME_SCAN = re.compile(r'(?<![\w$.@%#])([A-Za-z_][A-Za-z0-9_$.%#]*)')


def _scan_names(s):
    out = []
    for m in _NAME_SCAN.finditer(s):
        nm = m.group(1).upper()
        if nm in FUNCS or nm in EXCLUDE or nm in out:
            continue
        out.append(nm)
    return out


def _split_terms(s):
    terms, depth, cur = [], 0, ''
    for ch in s:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch.isspace() and depth == 0:
            if cur:
                terms.append(cur)
            cur = ''
        else:
            cur += ch
    if cur:
        terms.append(cur)
    return terms


_SKIP_START = re.compile(
    r'^(for|next|if|endif|else|while|wend|smpl|include|call|subroutine|endsub|string|scalar|'
    r'vector|matrix|show|freeze|delete|wfopen|wfsave|pageselect|assign|solve|output|close|exit|'
    r'stop|logmode|model|copy|fetch|store|group|graph|table|text|coef|sample|pagecreate|wfcreate|'
    r'open|save|print|setmaxerrs|mode|tic|toc|statusline)\b', re.I)
_SPEC = re.compile(r'^(?:equation\s+)?(?:[A-Za-z_][\w$]*\.)?(ls|tsls|gmm|nls|arch|binary|logit|probit|'
                   r'glm|ardl|qreg|robustls|cointreg)\b(?:\([^)]*\))?\s*', re.I)


def parse_equation_line(raw, source):
    """Parse one logical line of EViews text. Returns Equation, ('addlink', y, af) or None."""
    s = raw.strip()
    if not s or s[0] in "'#;!%{" or s.startswith('//'):
        return None
    s = re.sub(r'\[[^\]]*\]', '', s).strip()
    if not s or s.startswith(':'):
        return None
    m = re.match(r'^@ADD\w*\s*(?:\([^)]*\))?\s+(\S+)\s+(\S+)', s, re.I)
    if m:
        return ('addlink', m.group(1).upper(), m.group(2).upper())
    if _SKIP_START.match(s):
        return None
    kind = 'equation'
    m = re.match(r'^@IDENTITY\s+', s, re.I)
    if m:
        kind, s = 'identity', s[m.end():]
    m = re.match(r'^(series|genr|frml)\s+', s, re.I)
    if m:
        kind, s = 'identity', s[m.end():]
    spec = False
    m = _SPEC.match(s)
    if m:
        spec, s = True, s[m.end():]
        if s.lower().startswith('equation '):
            return None
    if s.startswith('@'):
        return None
    text = s.strip()
    try:
        toks = tokenize(s)
    except ParseError:
        toks = None
    if toks is not None:
        depth, eqi = 0, -1
        for i, (_, v) in enumerate(toks):
            if v == '(':
                depth += 1
            elif v == ')':
                depth -= 1
            elif v == '=' and depth == 0:
                eqi = i
                break
        if eqi > 0:
            try:
                lhs = parse_expr(toks[:eqi])
                rhs = parse_expr(toks[eqi + 1:])
            except ParseError as e:
                lhs = rhs = None
                why = str(e)
            if lhs is not None:
                lv, rv = ast_vars(lhs), ast_vars(rhs)
                if not lv:
                    return None
                tgt = lv[0]
                drivers = rv + [v for v in lv[1:] if v not in rv]
                coef = ast_has(rhs, ('coef',)) or ast_has(lhs, ('coef',))
                return Equation(tgt, drivers, text, kind, source, lhs, rhs, lv, rv,
                                evaluable=not coef,
                                why='has unsubstituted coefficients C(i)' if coef else '')
        elif eqi < 0 and spec:
            terms = _split_terms(s)
            names = []
            for t in terms:
                try:
                    names.append(ast_vars(parse_expr(tokenize(t))))
                except ParseError:
                    names.append(_scan_names(t))
            if not names or not names[0]:
                return None
            tgt = names[0][0]
            drivers = []
            for lst in names[1:] + [names[0][1:]]:
                for v in lst:
                    if v not in drivers:
                        drivers.append(v)
            return Equation(tgt, drivers, text, 'spec', source,
                            why='specification only (no coefficients)')
    # fallback: dependencies only
    if '=' not in s:
        return None
    left, right = s.split('=', 1)
    lv, rv = _scan_names(left), _scan_names(right)
    if not lv:
        return None
    return Equation(lv[0], rv + [v for v in lv[1:] if v not in rv], text, kind, source,
                    why='could not parse expression; dependencies only')


def _logical_lines(text):
    out, buf = [], ''
    for line in text.splitlines():
        st = line.strip()
        if not st:
            if buf:
                out.append(buf)
                buf = ''
            continue
        if buf and (buf.count('(') > buf.count(')') or re.search(r'[-+*/^(,=]\s*$', buf)
                    or re.match(r'^[-+*/^]', st)) and not st.startswith(("'", ':', '@')):
            buf += ' ' + st
        else:
            if buf:
                out.append(buf)
            buf = st
    if buf:
        out.append(buf)
    return out


EQ_EXT = {'.txt', '.prg', '.eq', '.eqs', '.mod', '.model', '.eviews', '.text'}


def load_equations(paths):
    eqs, addlinks, stats, fails = {}, {}, Counter(), []
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, fs in os.walk(p):
                for f in sorted(fs):
                    if os.path.splitext(f)[1].lower() in EQ_EXT:
                        files.append(os.path.join(root, f))
        else:
            files.append(p)
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext in ('.csv', '.tsv', '.xlsx', '.xlsm'):
            grid, _ = read_grid(f)
            lines = [str(c) for r in grid for c in r if c and ('=' in str(c) or _SPEC.match(str(c).strip()))]
        else:
            with open(f, 'r', encoding='utf-8-sig', errors='replace') as fh:
                lines = _logical_lines(fh.read())
        stats['files'] += 1
        for ln in lines:
            stats['lines'] += 1
            r = parse_equation_line(ln, os.path.basename(f))
            if r is None:
                if '=' in ln and len(fails) < 30 and not _SKIP_START.match(ln.strip()):
                    fails.append(ln[:160])
                continue
            if isinstance(r, tuple):
                addlinks[r[1]] = r[2]
                continue
            stats['parsed'] += 1
            prev = eqs.get(r.target)
            if prev is not None:
                stats['duplicate targets'] += 1
                if prev.evaluable and not r.evaluable:
                    continue
            eqs[r.target] = r
    stats['equations'] = len(eqs)
    stats['evaluable'] = sum(1 for e in eqs.values() if e.evaluable)
    return eqs, addlinks, stats, fails


# --------------------------------------------------------------------------
# The comparison engine
# --------------------------------------------------------------------------
def _sum(xs):
    return sum(v for v in xs if v == v)


def _sumabs(xs):
    return sum(abs(v) for v in xs if v == v)


def ast_contemp(node, out=None):
    """Variables used at lag 0 (same-period dependencies)."""
    if out is None:
        out = set()
    k = node[0]
    if k == 'var':
        if node[2] == 0:
            out.add(node[1])
    elif k == 'neg':
        ast_contemp(node[1], out)
    elif k == 'bin':
        ast_contemp(node[2], out)
        ast_contemp(node[3], out)
    elif k == 'call':
        for a in node[2]:
            ast_contemp(a, out)
    return out


def lhs_form(eq):
    l, t = eq.lhs, eq.target
    if l == ('var', t, 0):
        return 'level'
    if l[0] == 'call' and len(l[2]) == 1 and l[2][0] == ('var', t, 0):
        f = l[1].lstrip('@')
        if f == 'DLOG':
            return 'dlog'
        if f in ('LOG', 'LN'):
            return 'log'
        if f == 'D':
            return 'd'
    return 'generic'


def parse_shock(s):
    """'1%' -> ('pct', 1.0); '0.25' or '+0.25' -> ('abs', 0.25); None -> None (use scenario B)."""
    if not s:
        return None
    s = s.strip()
    if s.endswith('%'):
        return ('pct', float(s[:-1]))
    return ('abs', float(s))


class Model:
    def __init__(self, A, B, eqs=None, addlinks=None, start=None, end=None, tol=1e-6,
                 geo_sep='.', adj_threshold=0.25, meta=None, verbose=True, af_suffix='_A'):
        t0 = time.time()
        self.A, self.B = A, B
        self.tol = tol
        self.geo_sep = geo_sep
        self.adj_threshold = adj_threshold
        self.verbose = verbose
        keys = sorted(set(A.keys) | set(B.keys))
        self.keys = keys
        self.labels = [A.keys.get(k) or B.keys.get(k) for k in keys]
        self.ppy = infer_ppy(keys)
        T = self.T = len(keys)
        kidx = {k: i for i, k in enumerate(keys)}
        self.a = self._align(A, kidx, T)
        self.b = self._align(B, kidx, T)
        self.desc = dict(B.desc)
        self.desc.update(A.desc)
        if meta:
            self.desc.update(meta)
        self.eqs = eqs or {}
        self.addlinks = addlinks or {}
        self.has_eqs = bool(self.eqs)
        self.common = sorted(set(self.a) & set(self.b))
        self.only_a = sorted(set(self.a) - set(self.b))
        self.only_b = sorted(set(self.b) - set(self.a))

        self.first_div = {}
        for n in self.common:
            fd = self._first_div(self.a[n], self.b[n])
            if fd is not None:
                self.first_div[n] = fd
        lab2i = {lab: i for i, lab in enumerate(self.labels)}

        def find_label(s):
            if s is None:
                return None
            if s in lab2i:
                return lab2i[s]
            p = parse_period(s)
            if p and p[0] in kidx:
                return kidx[p[0]]
            raise SystemExit('period %r not found; data runs %s .. %s'
                             % (s, self.labels[0], self.labels[-1]))
        ws = find_label(start)
        we = find_label(end)
        if ws is None:
            ws = min(self.first_div.values()) if self.first_div else 0
        if we is None:
            we = T - 1
        if we < ws:
            raise SystemExit('end period is before start period')
        self.ws, self.we = ws, we

        # graph
        self.parents = {t: [d for d in e.drivers if d != t] for t, e in self.eqs.items()}
        self.children = defaultdict(list)
        for t, ps in self.parents.items():
            for p in ps:
                self.children[p].append(t)

        # add-factors: series named like <target>_A, or linked with @ADD
        self.af_suffixes = [s.strip().upper() for s in (af_suffix or '').split(',') if s.strip()]
        allnames = set(self.a) | set(self.b)
        self.af, self.af_series = {}, set()
        for t, e in self.eqs.items():
            for d in e.drivers:
                if self.is_af_name(d):
                    self.af_series.add(d)
                    self.af.setdefault(t, d)
            if t in self.af:
                continue
            for c in self._af_cands(t):
                if c in allnames:
                    self.af[t] = c
                    self.af_series.add(c)
                    break

        self.stats = {}
        for n in self.common:
            st = self._stats(n)
            if st:
                self.stats[n] = st
        self.differs = {n for n, st in self.stats.items() if st['differs']}
        self.differs_any = set(self.first_div)
        self._resid, self._dd, self._lform, self._contemp = {}, {}, {}, {}
        self.load_secs = time.time() - t0

    # ---- helpers
    def _align(self, sc, kidx, T):
        out = {}
        maps = [[kidx[k] for k in blk] for blk in sc.blocks]
        template = array('d', [NAN]) * T
        for name, (bid, arr) in sc.series.items():
            full = array('d', template)
            for i, v in zip(maps[bid], arr):
                full[i] = v
            out[name] = full
        return out

    def is_af_name(self, d):
        for s in self.af_suffixes:
            if d.endswith(s):
                return True
            if self.geo_sep and self.geo_sep in d and d.rsplit(self.geo_sep, 1)[0].endswith(s):
                return True
        return False

    def _af_cands(self, t):
        out = []
        if t in self.addlinks:
            out.append(self.addlinks[t])
        for s in self.af_suffixes:
            out.append(t + s)
            if self.geo_sep and self.geo_sep in t:
                st, g = t.rsplit(self.geo_sep, 1)
                out.append(st + s + self.geo_sep + g)
        return out

    def _first_div(self, a, b):
        tol = self.tol
        for t in range(self.T):
            x, y = a[t], b[t]
            if x == x and y == y and abs(y - x) > tol * max(1.0, abs(x)):
                return t
        return None

    def _scale(self, a, b=None):
        for s in (a, b):
            if s is None:
                continue
            ds = [s[t] - s[t - 1] for t in range(1, len(s)) if s[t] == s[t] and s[t - 1] == s[t - 1]]
            if len(ds) >= 3:
                m = sum(ds) / len(ds)
                sd = math.sqrt(sum((d - m) ** 2 for d in ds) / (len(ds) - 1))
                if sd > 1e-12:
                    return sd
        lv = [abs(v) for v in a if v == v]
        base = (sum(lv) / len(lv)) if lv else 0.0
        return max(base * 0.01, 1e-12)

    def _stats(self, n):
        a, b = self.a[n], self.b[n]
        maxabs, tmax, last, cnt = 0.0, -1, -1, 0
        for t in range(self.ws, self.we + 1):
            x, y = a[t], b[t]
            if x != x or y != y:
                continue
            cnt += 1
            last = t
            d = abs(y - x)
            if d > maxabs:
                maxabs, tmax = d, t
        if cnt == 0:
            return None
        if tmax < 0:
            tmax = last
        dmax = b[tmax] - a[tmax]
        dend = b[last] - a[last]
        pmax = dmax / abs(a[tmax]) * 100 if a[tmax] != 0 else NAN
        pend = dend / abs(a[last]) * 100 if a[last] != 0 else NAN
        differs = maxabs > self.tol * max(1.0, abs(a[tmax])) and n in self.first_div
        return dict(maxabs=maxabs, dmax=dmax, tmax=tmax, dend=dend, pmax=pmax, pend=pend,
                    tend=last, score=min(1e6, maxabs / self._scale(a, b)) if differs else 0.0, differs=differs,
                    first=self.first_div.get(n))

    def geo(self, n):
        if self.geo_sep and self.geo_sep in n:
            return n.rsplit(self.geo_sep, 1)[1]
        return ''

    def stem(self, n):
        if self.geo_sep and self.geo_sep in n:
            return n.rsplit(self.geo_sep, 1)[0]
        return n

    def series(self, n, which):
        s = (self.a if which == 'a' else self.b).get(n)
        if s is None:
            s = (self.b if which == 'a' else self.a).get(n)
        return s

    def kind(self, n):
        """X assumption (no equation), F add-factor series, A equation whose change comes mostly
        from its add-factor/residual, E explained by its equation, N not evaluable, - no equations."""
        if not self.has_eqs:
            return '-'
        if n in self.af_series and n not in self.eqs:
            return 'F'
        eq = self.eqs.get(n)
        if eq is None:
            return 'X'
        if not eq.evaluable:
            return 'N'
        d = self.dyn(n)
        if d is None or 'error' in d:
            return 'N'
        if n in self.differs and d['af_share'] >= self.adj_threshold:
            return 'A'
        return 'E'

    def _getter(self, which):
        prim, sec = (self.a, self.b) if which == 'a' else (self.b, self.a)

        def get(n):
            s = prim.get(n)
            if s is None:
                s = sec.get(n)
            if s is None:
                raise EvalError('series %s is not in the data' % n)
            return s
        return get

    # ---- residuals (= add-factor + equation error) in each scenario
    def resid(self, n, which):
        key = (n, which)
        r = self._resid.get(key)
        if r is None:
            eq = self.eqs[n]
            g = self._getter(which)
            l = evaluate(eq.lhs, g, self.ws, self.we, self.ppy)
            rh = evaluate(eq.rhs, g, self.ws, self.we, self.ppy)
            r = ([x - y for x, y in zip(l, rh)], l)
            self._resid[key] = r
        return r[0]

    def resid_changed(self, n):
        ra = self.resid(n, 'a')
        rb = self.resid(n, 'b')
        la = self._resid[(n, 'a')][1]
        for x, y, l in zip(ra, rb, la):
            if x == x and y == y and abs(y - x) > self.tol * max(1.0, abs(l) if l == l else 1.0):
                return True
        return False

    # ---- model simulation (solve equations period by period)
    def _form(self, n):
        f = self._lform.get(n)
        if f is None:
            f = self._lform[n] = lhs_form(self.eqs[n])
        return f

    def _cont(self, n):
        c = self._contemp.get(n)
        if c is None:
            e = self.eqs[n]
            c = ast_contemp(e.rhs) | (ast_contemp(e.lhs) - {n})
            self._contemp[n] = c
        return c

    def _order(self, nodes):
        ns = set(nodes)
        preds = {n: [p for p in self._cont(n) if p in ns and p != n] for n in nodes}
        succ = defaultdict(list)
        indeg = {n: len(preds[n]) for n in nodes}
        for n, ps in preds.items():
            for p in ps:
                succ[p].append(n)
        q = deque(sorted(n for n in nodes if indeg[n] == 0))
        order = []
        while q:
            u = q.popleft()
            order.append(u)
            for v in succ[u]:
                indeg[v] -= 1
                if indeg[v] == 0:
                    q.append(v)
        done = set(order)
        cyclic = [n for n in nodes if n not in done]
        return order, cyclic

    def _solve(self, n, get, arr, t, res):
        eq = self.eqs[n]
        target = evaluate(eq.rhs, get, t, t, self.ppy)[0] + res
        f = self._form(n)
        if f == 'level':
            y = target
        elif f == 'dlog':
            y = arr[t - 1] * _exp(target) if t > 0 else NAN
        elif f == 'log':
            y = _exp(target)
        elif f == 'd':
            y = arr[t - 1] + target if t > 0 else NAN
        else:
            y = self._secant(eq, get, arr, t, target)
        arr[t] = y
        return y

    def _secant(self, eq, get, arr, t, target):
        def fn(y):
            arr[t] = y
            return evaluate(eq.lhs, get, t, t, self.ppy)[0] - target
        y0 = arr[t]
        if y0 != y0:
            y0 = arr[t - 1] if t > 0 and arr[t - 1] == arr[t - 1] else 1.0
        f0 = fn(y0)
        if f0 != f0:
            return NAN
        if abs(f0) < 1e-13 * (1 + abs(target)):
            return y0
        y1 = y0 * (1 + 1e-4) + 1e-8
        f1 = fn(y1)
        for _ in range(50):
            if f1 != f1 or f1 == f0:
                break
            y2 = y1 - f1 * (y1 - y0) / (f1 - f0)
            y0, f0, y1 = y1, f1, y2
            f1 = fn(y1)
            if abs(f1) < 1e-12 * (1 + abs(target)) or abs(y1 - y0) < 1e-13 * (1 + abs(y1)):
                break
        return y1

    def simulate(self, nodes, override=None, resid_b=()):
        """Re-solve `nodes` over the window. Everything else is fixed at scenario-A values,
        except series in `override` (name -> array). Residuals/add-factors come from A,
        except for equations in `resid_b`, which use scenario B's."""
        override = override or {}
        nodes = [n for n in nodes if n in self.eqs and self.eqs[n].evaluable and n not in override]
        sim = {}
        for n in nodes:
            s = self.series(n, 'a')
            if s is None:
                raise EvalError('series %s is not in the data' % n)
            sim[n] = array('d', s)
        res = {n: self.resid(n, 'b' if n in resid_b else 'a') for n in nodes}
        A, B = self.a, self.b

        def get(name):
            s = sim.get(name)
            if s is None:
                s = override.get(name)
                if s is None:
                    s = A.get(name)
                    if s is None:
                        s = B.get(name)
                        if s is None:
                            raise EvalError('series %s is not in the data' % name)
            return s
        order, cyclic = self._order(nodes)
        ws = self.ws
        for t in range(ws, self.we + 1):
            i = t - ws
            for n in order:
                self._solve(n, get, sim[n], t, res[n][i])
            if cyclic:
                for _ in range(200):
                    mx = 0.0
                    for n in cyclic:
                        old = sim[n][t]
                        new = self._solve(n, get, sim[n], t, res[n][i])
                        if new == new and old == old:
                            mx = max(mx, abs(new - old) / (1.0 + abs(old)))
                    if mx < 1e-11:
                        break
        return sim

    # ---- contributions of each driver to the gap (B - A) of one variable
    def dyn(self, n, periods=False):
        d = self._dd.get(n)
        if d is not None and (not periods or 'error' in d or 'contrib' in d):
            return d
        eq = self.eqs.get(n)
        if eq is None or not eq.evaluable:
            return None
        try:
            d = self._dyn(eq, periods)
        except EvalError as e:
            d = {'error': str(e)}
        except (IndexError, ValueError, ZeroDivisionError, OverflowError, TypeError) as e:
            d = {'error': 'evaluation failed: %s' % e}
        self._dd[n] = d
        return d

    def _dyn(self, eq, periods):
        n = eq.target
        ws, we = self.ws, self.we
        A, Bv = self.series(n, 'a'), self.series(n, 'b')
        if A is None:
            raise EvalError('series %s is not in the data' % n)
        rng = range(ws, we + 1)
        gap = [Bv[t] - A[t] for t in rng]
        base = self.simulate([n])[n]
        fit = max([abs(base[t] - A[t]) / (1.0 + abs(A[t])) for t in rng if base[t] == base[t] and A[t] == A[t]] or [NAN])
        if fit != fit:
            raise EvalError('equation gives missing values over the window')
        drivers = []
        for d in eq.rhs_vars + eq.lhs_vars[1:]:
            if d != n and d in self.differs_any and d not in drivers:
                drivers.append(d)
        contrib = {}
        for d in drivers:
            s = self.simulate([n], override={d: self.series(d, 'b')})[n]
            contrib[d] = [s[t] - base[t] for t in rng]
        af = None
        if self.resid_changed(n):
            s = self.simulate([n], resid_b={n})[n]
            af = [s[t] - base[t] for t in rng]
            big = max([abs(g) for g in gap if g == g] + [0.0])
            lvl = max([abs(A[t]) for t in rng if A[t] == A[t]] + [1.0])
            if max([abs(v) for v in af if v == v] + [0.0]) < max(1e-3 * big, self.tol * lvl):
                af = None  # rounding noise, not a real add-factor change
        inter = []
        for i, g in enumerate(gap):
            c = sum(v[i] for v in contrib.values() if v[i] == v[i])
            if af and af[i] == af[i]:
                c += af[i]
            inter.append(g - c)
        iend = max([i for i, g in enumerate(gap) if g == g] or [-1])
        if iend < 0:
            raise EvalError('no overlapping data in the window')
        ge = gap[iend]

        def sh(v):
            return v / ge if abs(ge) > 1e-12 * max(1.0, abs(A[ws + iend])) else None
        end = {d: c[iend] for d, c in contrib.items()}
        afe = af[iend] if af else 0.0
        absd = sum(_sumabs(c) for c in contrib.values())
        absaf = _sumabs(af) if af else 0.0
        out = dict(target=n, lhs=ast_text(eq.lhs), iend=iend, gap_end=ge, end=end, af_end=afe,
                   inter_end=inter[iend], shares={d: sh(v) for d, v in end.items()},
                   af_share_end=sh(afe) if af else None, formula_share_end=sh(sum(end.values())),
                   af_share=absaf / (absaf + absd) if absaf + absd > 1e-15 else 0.0,
                   af_name=self.af.get(n), af_changed=af is not None, fit_err=fit)
        if periods:
            out.update(contrib=contrib, af=af, inter=inter, gap=gap)
        return out

    def share(self, child, parent):
        d = self.dyn(child)
        if not d or 'error' in d:
            return None
        return d['shares'].get(parent)

    # ---- graph walks
    def descendants(self, x):
        seen, q = {x}, deque([x])
        while q:
            u = q.popleft()
            for v in self.children.get(u, ()):
                if v not in seen:
                    seen.add(v)
                    q.append(v)
        seen.discard(x)
        return seen

    def ancestors(self, y):
        seen, q = {y}, deque([y])
        while q:
            u = q.popleft()
            for v in self.parents.get(u, ()):
                if v not in seen:
                    seen.add(v)
                    q.append(v)
        seen.discard(y)
        return seen

    # ---- impact of X on Y (multi-equation simulation)
    def impact(self, x, y=None, shock=None):
        """Re-run the equations between X and Y with only X changed (to its scenario-B path,
        or by a hypothetical shock). Returns (effects dict name -> list over window, info)."""
        desc = self.descendants(x)
        if y is not None:
            nodes = desc & (self.ancestors(y) | {y})
        else:
            nodes = desc
        ev = [n for n in nodes if n in self.eqs and self.eqs[n].evaluable]
        blocked = sorted(n for n in nodes if n not in self.eqs or not self.eqs[n].evaluable)
        xa = self.series(x, 'a')
        if xa is None:
            raise EvalError('%s is not in the data' % x)
        if shock is None:
            xo = self.series(x, 'b')
        else:
            xo = array('d', xa)
            for t in range(self.ws, self.we + 1):
                xo[t] = xa[t] * (1 + shock[1] / 100.0) if shock[0] == 'pct' else xa[t] + shock[1]
        base = self.simulate(ev)
        sim = self.simulate(ev, override={x: xo})
        rng = range(self.ws, self.we + 1)
        eff = {n: [sim[n][t] - base[n][t] for t in rng] for n in ev}
        eff[x] = [xo[t] - xa[t] for t in rng]
        return eff, {'nodes': len(ev), 'blocked': blocked}

    def sources(self, y):
        """Scenario shocks upstream of Y: changed assumptions / add-factor series / data for
        equations that cannot be evaluated, and equations whose add-factor (residual) changed."""
        anc = self.ancestors(y) | {y}
        out = []
        for s in anc:
            ev = s in self.eqs and self.eqs[s].evaluable
            if not ev and s in self.differs_any and s != y:
                out.append(('data', s))
            elif ev and self.resid_changed(s):
                d = self.dyn(s)
                if d and 'error' not in d and d['af_changed']:
                    out.append(('addfactor', s))
        return out, anc

    def attribute(self, y):
        """Split Y's gap into the effect of each upstream scenario shock."""
        srcs, anc = self.sources(y)
        if y not in self.eqs or not self.eqs[y].evaluable:
            raise EvalError('%s has no evaluable equation' % y)
        nodes = [n for n in anc if n in self.eqs and self.eqs[n].evaluable]
        base = self.simulate(nodes)[y]
        rng = range(self.ws, self.we + 1)
        res = []
        for kind, s in srcs:
            sub = (self.descendants(s) | {s}) & anc if kind == 'addfactor' else self.descendants(s) & anc
            sub = [n for n in sub if n in self.eqs and self.eqs[n].evaluable]
            if kind == 'data':
                sim = self.simulate(sub, override={s: self.series(s, 'b')})
            else:
                sim = self.simulate(sub, resid_b={s})
            yy = sim.get(y, base)
            res.append((kind, s, [yy[t] - base[t] for t in rng]))
        allb = {s for k, s in srcs if k == 'data'}
        sim = self.simulate(nodes, override={s: self.series(s, 'b') for s in allb},
                            resid_b={s for k, s in srcs if k == 'addfactor'})
        joint = [sim[y][t] - base[t] for t in rng]
        A, Bv = self.series(y, 'a'), self.series(y, 'b')
        gap = [Bv[t] - A[t] for t in rng]
        return res, joint, gap

    # ---- ranking helpers
    def ranked(self):
        return sorted(self.differs, key=lambda n: -self.stats[n]['score'])

    def entry_points(self):
        out = [n for n in self.differs if self.kind(n) in ('X', 'A', 'F')]
        return sorted(out, key=lambda n: (self.stats[n]['first'], -self.stats[n]['score']))

    def downstream(self, x, depth=4):
        dist, prev = {x: 0}, {}
        q = deque([x])
        while q:
            u = q.popleft()
            if dist[u] >= depth:
                continue
            for v in self.children.get(u, ()):
                if v not in dist:
                    dist[v] = dist[u] + 1
                    prev[v] = u
                    q.append(v)
        return dist, prev

    def strongest_path(self, x, y, maxdepth=15):
        import heapq
        INF = float('inf')
        best, prev, hops = {x: 0.0}, {}, {x: 0}
        pq = [(0.0, x)]
        while pq:
            c, u = heapq.heappop(pq)
            if c > best.get(u, INF):
                continue
            if u == y:
                break
            if hops[u] >= maxdepth:
                continue
            for v in self.children.get(u, ()):
                if v not in self.differs_any and v != y:
                    continue
                sh = self.share(v, u)
                w = -math.log(min(1.0, abs(sh))) if sh else 7.0
                if c + w < best.get(v, INF):
                    best[v], prev[v], hops[v] = c + w, u, hops[u] + 1
                    heapq.heappush(pq, (c + w, v))
        if y not in best or y == x:
            return None
        p = [y]
        while p[-1] != x:
            p.append(prev[p[-1]])
        return p[::-1]

    def path_product(self, p):
        prod = 1.0
        for a, b in zip(p, p[1:]):
            sh = self.share(b, a)
            if sh is None:
                return None
            prod *= sh
        return prod

    def paths(self, x, y, k=5, maxdepth=12):
        dist = {x: 0}
        q = deque([x])
        while q:
            u = q.popleft()
            if u == y or dist[u] >= maxdepth:
                continue
            for v in self.children.get(u, ()):
                if v not in dist:
                    dist[v] = dist[u] + 1
                    q.append(v)
        if y not in dist:
            return []
        out = []

        def back(node, acc):
            if len(out) >= k:
                return
            if node == x:
                out.append([x] + acc[::-1])
                return
            for p in self.parents.get(node, ()):
                if dist.get(p) == dist[node] - 1:
                    back(p, acc + [node])
        back(y, [])
        return out

    def ranked_paths(self, x, y, k=5):
        ps = self.paths(x, y, k=40)
        sp = self.strongest_path(x, y)
        if sp and sp not in ps:
            ps.append(sp)

        def key(p):
            v = self.path_product(p)
            return -abs(v) if v is not None else 1
        ps.sort(key=key)
        return ps[:k]


# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------
def fmt(v, sig=4):
    if v is None or v != v:
        return 'n/a'
    if v == 0:
        return '0'
    if abs(v) >= 1e4:
        return '{:,.0f}'.format(v)
    return ('%.' + str(sig) + 'g') % v


def fpct(v):
    if v is None or v != v:
        return 'n/a'
    return '%+.2f%%' % v


def fsh(v):
    return '   n/a' if v is None or v != v else '%5.0f%%' % (100 * v)


def _r(v, sig=6):
    if v is None or v != v or v in (float('inf'), float('-inf')):
        return None
    if v == 0:
        return 0
    return float('%.*g' % (sig, v))


KIND_TEXT = {'X': 'assumption (no equation)', 'F': 'add-factor series', 'A': 'equation, mostly add-factor',
             'E': 'explained by its equation', 'N': 'equation not evaluable', '-': ''}


# --------------------------------------------------------------------------
# Terminal commands
# --------------------------------------------------------------------------
def cmd_info(M, args, eqstats=None, fails=None):
    print('Scenario A: %s' % M.A.label)
    for n in M.A.notes:
        print('   ', n)
    print('Scenario B: %s' % M.B.label)
    for n in M.B.notes:
        print('   ', n)
    print('Periods: %s .. %s (%d), %d per year' % (M.labels[0], M.labels[-1], M.T, M.ppy))
    print('Series: %d in both, %d only in A, %d only in B' % (len(M.common), len(M.only_a), len(M.only_b)))
    print('Comparison window: %s .. %s' % (M.labels[M.ws], M.labels[M.we]))
    print('Series that differ: %d' % len(M.differs))
    if eqstats:
        print('Equations: %s' % ', '.join('%s=%d' % kv for kv in sorted(eqstats.items())))
        print('Add-factor series found: %d (suffix %s), e.g. %s' % (
            len(M.af_series), ','.join(M.af_suffixes), ', '.join(sorted(M.af_series)[:5])))
        bad = Counter(e.why for e in M.eqs.values() if not e.evaluable)
        for why, c in bad.most_common():
            print('    not evaluable: %d (%s)' % (c, why))
        errs, fits = Counter(), []
        for n in list(M.differs)[:3000]:
            d = M.dyn(n)
            if d and 'error' in d:
                errs[re.sub(r'series \S+', 'series X', d['error'])] += 1
        for why, c in errs.most_common(8):
            print('    evaluation error: %d (%s)' % (c, why))
        missing = Counter()
        for e in M.eqs.values():
            for d in e.drivers:
                if d not in M.a and d not in M.b and d not in M.eqs:
                    missing[d] += 1
        if missing:
            print('    drivers not found in data: %d, e.g. %s'
                  % (len(missing), ', '.join(n for n, _ in missing.most_common(8))))
    if fails:
        print('Lines with "=" that could not be read (first %d):' % len(fails))
        for f in fails[:10]:
            print('   ', f)
    print('Loaded in %.1fs' % M.load_secs)


def _filter(M, names, args):
    out = names
    if getattr(args, 'geo', None):
        g = args.geo.upper()
        out = [n for n in out if M.geo(n) == g]
    if getattr(args, 'match', None):
        rx = re.compile(args.match, re.I)
        out = [n for n in out if rx.search(n) or rx.search(M.desc.get(n, ''))]
    if getattr(args, 'type', None):
        out = [n for n in out if M.kind(n) == args.type.upper()]
    return out


def _formula_af(M, n):
    d = M.dyn(n)
    if not d or 'error' in d:
        return None, None
    return d['formula_share_end'], d['af_share_end'] if d['af_changed'] else 0.0


def cmd_top(M, args):
    names = _filter(M, M.ranked(), args)
    print('A = %s   B = %s   window %s..%s   %d series differ'
          % (M.A.label, M.B.label, M.labels[M.ws], M.labels[M.we], len(M.differs)))
    print('score = largest gap / typical period-to-period move of the series in A')
    print('formula / add-f. = share of the gap at the window end explained by the drivers / by the add-factor\n')
    w = max([len(n) for n in names[:args.n]] + [6])
    print('%-*s %7s %11s %9s %9s %8s %7s %7s  %s' % (w, 'series', 'score', 'max gap', 'max %', 'end %',
                                                     'starts', 'formula', 'add-f.', 'type'))
    for n in names[:args.n]:
        s = M.stats[n]
        fs, af = _formula_af(M, n)
        print('%-*s %7.1f %11s %9s %9s %8s %7s %7s  %s  %s' % (
            w, n, s['score'], fmt(s['dmax']), fpct(s['pmax']), fpct(s['pend']),
            M.labels[s['first']] if s['first'] is not None else '', fsh(fs) if fs is not None else '',
            fsh(af) if af is not None else '', M.kind(n), M.desc.get(n, '')[:40]))
    if args.csv:
        with open(args.csv, 'w', newline='', encoding='utf-8') as f:
            wr = csv.writer(f)
            wr.writerow(['series', 'description', 'geo', 'type', 'score', 'max_gap', 'max_gap_period',
                         'max_gap_pct', 'end_gap', 'end_gap_pct', 'first_divergence',
                         'share_from_formula_end', 'share_from_addfactor_end', 'addfactor_series'])
            for n in M.ranked():
                s = M.stats[n]
                fs, af = _formula_af(M, n)
                wr.writerow([n, M.desc.get(n, ''), M.geo(n), M.kind(n), _r(s['score']), _r(s['dmax']),
                             M.labels[s['tmax']], _r(s['pmax']), _r(s['dend']), _r(s['pend']),
                             M.labels[s['first']] if s['first'] is not None else '',
                             _r(fs, 4), _r(af, 4), M.af.get(n, '')])
        print('\nAll %d differing series written to %s' % (len(M.differs), args.csv))


def cmd_addfactors(M, args):
    """Formula vs add-factor for every equation whose variable differs."""
    rows = []
    for n in M.ranked():
        if n not in M.eqs or not M.eqs[n].evaluable:
            continue
        d = M.dyn(n)
        if not d or 'error' in d:
            continue
        rows.append((n, d))
    rows.sort(key=lambda r: -abs(r[1]['af_end']) / (abs(r[1]['gap_end']) + 1e-30) if r[1]['af_changed'] else 0)
    changed = [r for r in rows if r[1]['af_changed']]
    print('%d of %d differing equations have a changed add-factor/residual.\n' % (len(changed), len(rows)))
    w = max([len(n) for n, _ in rows[:args.n]] + [6])
    print('%-*s %11s %9s %8s %8s %8s  %s' % (w, 'series', 'gap end', 'gap %', 'formula', 'add-f.', 'other', 'add-factor series'))
    for n, d in (changed if not args.all else rows)[:args.n]:
        s = M.stats[n]
        afn = d['af_name'] or ''
        afc = ''
        st = M.stats.get(afn) if afn else None
        if st and st['differs']:
            afc = ' (changed: largest %s in %s)' % (fmt(st['dmax']), M.labels[st['tmax']])
        elif afn:
            afc = ' (series unchanged; residual moved)'
        ge = d['gap_end']
        oth = d['inter_end'] / ge if abs(ge) > 1e-12 else None
        print('%-*s %11s %9s %8s %8s %8s  %s%s' % (w, n, fmt(ge), fpct(s['pend']), fsh(d['formula_share_end']),
                                                fsh(d['af_share_end']) if d['af_changed'] else '     0%', fsh(oth), afn, afc))


def _norm(M, name):
    n = name.strip().upper()
    if n in M.a or n in M.b or n in M.eqs or n in M.children:
        return n
    cands = [x for x in set(M.a) | set(M.b) if n in x][:10]
    raise SystemExit('%s not found.%s' % (name, (' Did you mean: ' + ', '.join(sorted(cands))) if cands else ''))


def _period_index(M, label):
    if label is None:
        return None
    labs = M.labels[M.ws:M.we + 1]
    if label in labs:
        return labs.index(label)
    p = parse_period(label)
    if p:
        for i, k in enumerate(M.keys[M.ws:M.we + 1]):
            if k == p[0]:
                return i
    raise SystemExit('period %s is not in the window %s..%s' % (label, labs[0], labs[-1]))


def print_contrib(M, n, prefix='', at=None):
    """Print each driver's contribution to the gap of n at one period. Returns contributions."""
    d = M.dyn(n, periods=True)
    eq = M.eqs.get(n)
    if eq is None:
        print(prefix + '(no equation: an assumption set by the scenario)')
        return None
    if not d or 'error' in d:
        print(prefix + '(cannot evaluate: %s)' % (d['error'] if d else eq.why))
        ch = [p for p in M.parents.get(n, []) if p in M.differs]
        if ch:
            print(prefix + 'drivers that differ: ' + ', '.join(ch[:15]))
        return None
    i = d['iend'] if at is None else at
    g = d['gap'][i]
    lab = M.labels[M.ws + i]
    print('%sgap B-A in %s = %s' % (prefix, lab, fmt(g)))
    items = sorted(((k, v[i]) for k, v in d['contrib'].items()), key=lambda kv: -abs(kv[1]))
    rows = [(k, v) for k, v in items]
    if d['af']:
        nm = d['af_name'] or 'residual'
        rows.append(('add-factor %s' % nm if d['af_name'] else 'add-factor / residual change', d['af'][i]))
    rows.append(('(interaction, non-linear)', d['inter'][i]))
    big = max([abs(v) for _, v in rows if v == v] + [abs(g) if g == g else 0, 1e-300])
    for k, v in rows:
        if k.startswith('(') and abs(v) < 1e-3 * big:
            continue
        sh = v / g if abs(g) > 1e-12 * big else None
        print('%s  %s %-12s %s' % (prefix, fsh(sh), fmt(v), k))
    if not d['contrib'] and not d['af']:
        print(prefix + '  none of its drivers differ and its add-factor did not change')
    return d


def cmd_explain(M, args):
    root = _norm(M, args.var)
    at = _period_index(M, args.at)
    seen = {root}

    def head(n, prefix):
        s = M.stats.get(n)
        info = ('gap at end %s (%s)' % (fmt(s['dend']), fpct(s['pend']))) if s and s['differs'] else 'no difference'
        print('%s%s  %s  [%s]' % (prefix, n, info, KIND_TEXT.get(M.kind(n), '')))

    def rec(n, depth, prefix):
        d = print_contrib(M, n, prefix + '    ', at)
        if depth <= 1 or not d:
            return
        i = d['iend'] if at is None else at
        g = d['gap'][i]
        for k, c in sorted(d['contrib'].items(), key=lambda kv: -abs(kv[1][i])):
            if k in seen or (abs(g) > 1e-12 and abs(c[i] / g) < args.min_share):
                continue
            seen.add(k)
            print()
            head(k, prefix + '  > ')
            rec(k, depth - 1, prefix + '    ')

    head(root, '')
    if M.desc.get(root):
        print('   ', M.desc[root])
    if root in M.eqs:
        print('    equation: %s' % M.eqs[root].text[:300])
    rec(root, args.depth, '')
    if args.by_period:
        d = M.dyn(root, periods=True)
        if d and 'error' not in d:
            labs = M.labels[M.ws:M.we + 1]
            cols = sorted(d['contrib'], key=lambda k: -abs(d['end'][k]))[:7]
            hdr = cols + (['add-factor'] if d['af'] else [])
            print('\nBy period, contribution to the gap of %s:' % root)
            print('%-9s %11s ' % ('period', 'gap') + ' '.join('%11s' % h[:11] for h in hdr))
            for i, lab in enumerate(labs):
                vals = [d['contrib'][k][i] for k in cols] + ([d['af'][i]] if d['af'] else [])
                print('%-9s %11s ' % (lab, fmt(d['gap'][i])) + ' '.join('%11s' % fmt(v) for v in vals))
    print('\nHow to read: each line is how much the gap (B-A) of the variable would be if ONLY that input '
          'moved from A to B, with the equation re-solved through time (so effects via its own lags are '
          'credited to the input that caused them).')


def cmd_path(M, args):
    x, y = _norm(M, args.x), _norm(M, args.y)
    ps = M.ranked_paths(x, y, k=args.k)
    if not ps:
        back = M.paths(y, x, k=1)
        print('%s does not affect %s through the loaded equations.' % (x, y))
        if back:
            print('(The reverse is true: %s -> %s)' % (y, ' -> '.join(back[0][1:])))
        if not M.children.get(x):
            print('%s is an end point: no other equation uses it.' % x)
        return
    _print_impact(M, x, y, None)
    print('\nChannels (ordered by how much of the change they carry):')
    for i, p in enumerate(ps, 1):
        print('Path %d:' % i)
        prod = 1.0
        ok = True
        for j, n in enumerate(p):
            s = M.stats.get(n)
            gap = ('%s (%s)' % (fmt(s['dend']), fpct(s['pend']))) if s and s['differs'] else 'no difference'
            if j == 0:
                print('   %s   gap at end %s' % (n, gap))
                continue
            sh = M.share(n, p[j - 1])
            if sh is None:
                ok = False
                shtxt = 'share n/a'
            else:
                prod *= sh
                shtxt = '%.0f%% of its gap comes from %s' % (100 * sh, p[j - 1])
            print('-> %s   gap at end %s   [%s]' % (n, gap, shtxt))
        if ok:
            print('   approx. share of %s\'s gap along this path: %.0f%%' % (y, 100 * prod))
        print()


def _print_impact(M, x, y, shock):
    eff, info = M.impact(x, y, shock)
    rng = range(M.ws, M.we + 1)
    labs = [M.labels[t] for t in rng]
    ey = eff.get(y)
    what = 'its change from A to B' if shock is None else ('a %s%s shock over the window' % (
        ('+' if shock[1] >= 0 else '') + ('%g' % shock[1]), '%' if shock[0] == 'pct' else ' unit'))
    if ey is None:
        print('%s cannot be re-solved (no evaluable equation), so the impact cannot be simulated.' % y)
        return
    A, Bv = M.series(y, 'a'), M.series(y, 'b')
    gap = [Bv[t] - A[t] for t in rng]
    iend = max(i for i, g in enumerate(ey) if g == g) if any(g == g for g in ey) else len(ey) - 1
    print('Impact of %s (%s) on %s, re-solving %d equations in between:' % (x, what, y, info['nodes']))
    pe = ey[iend] / abs(A[M.ws + iend]) * 100 if A[M.ws + iend] else NAN
    line = '  at %s: %s changes by %s (%s)' % (labs[iend], y, fmt(ey[iend]), fpct(pe))
    if shock is None and abs(gap[iend]) > 1e-12:
        line += ' = %.0f%% of its total gap %s' % (100 * ey[iend] / gap[iend], fmt(gap[iend]))
    print(line)
    if info['blocked']:
        print('  note: %d variables in between have no evaluable equation and were held fixed (e.g. %s)'
              % (len(info['blocked']), ', '.join(info['blocked'][:5])))
    return eff, gap, labs


def cmd_impact(M, args):
    x = _norm(M, args.x)
    shock = parse_shock(args.shock)
    if args.y:
        y = _norm(M, args.y)
        if y != x and y not in M.descendants(x):
            print('%s does not affect %s through the loaded equations (no chain of equations links them).' % (x, y))
            return
        r = _print_impact(M, x, y, shock)
        if r and args.by_period:
            eff, gap, labs = r
            print('\n%-9s %12s %12s %8s' % ('period', 'impact', 'total gap', 'share'))
            for i, lab in enumerate(labs):
                sh = eff[y][i] / gap[i] if shock is None and abs(gap[i]) > 1e-12 else None
                print('%-9s %12s %12s %8s' % (lab, fmt(eff[y][i]), fmt(gap[i]), fsh(sh)))
        return
    eff, info = M.impact(x, None, shock)
    rows = []
    for n, e in eff.items():
        if n == x:
            continue
        v = e[-1] if e[-1] == e[-1] else NAN
        a = M.series(n, 'a')[M.we]
        p = v / abs(a) * 100 if a else NAN
        g = M.stats.get(n, {}).get('dend')
        rows.append((n, v, p, (v / g) if (shock is None and g) else None))
    rows.sort(key=lambda r: -abs(r[2]) if r[2] == r[2] else 0)
    print('Variables affected by %s (%d re-solved). Effect at %s:' % (x, info['nodes'], M.labels[M.we]))
    print('%-24s %12s %9s %8s' % ('variable', 'effect', '%', 'of gap'))
    for n, v, p, s in rows[:args.n]:
        print('%-24s %12s %9s %8s' % (n, fmt(v), fpct(p), fsh(s) if s is not None else ''))


def cmd_attribute(M, args):
    y = _norm(M, args.y)
    res, joint, gap = M.attribute(y)
    i = max(j for j, g in enumerate(gap) if g == g)
    g = gap[i]
    print('%s: gap B-A at %s = %s. Split across the scenario shocks upstream of it:' % (y, M.labels[M.ws + i], fmt(g)))
    res.sort(key=lambda r: -abs(r[2][i]) if r[2][i] == r[2][i] else 0)
    shown = 0
    for kind, s, e in res:
        if abs(e[i]) < 1e-9 * max(1.0, abs(g)):
            continue
        lab = s if kind == 'data' else 'add-factor of %s%s' % (s, (' (%s)' % M.af[s]) if s in M.af else '')
        print('  %s %-12s %s  [%s]' % (fsh(e[i] / g if abs(g) > 1e-12 else None), fmt(e[i]), lab,
                                       KIND_TEXT.get(M.kind(s), '') if kind == 'data' else 'equation residual'))
        shown += 1
        if shown >= args.n:
            break
    other = g - sum(e[i] for _, _, e in res if e[i] == e[i])
    print('  %s %-12s (interaction between shocks)' % (fsh(other / g if abs(g) > 1e-12 else None), fmt(other)))
    if abs(joint[i] - g) > 1e-6 * max(1.0, abs(g)):
        print('  note: re-solving with all shocks gives %s vs actual gap %s (model fit difference)' % (fmt(joint[i]), fmt(g)))


# --------------------------------------------------------------------------
# HTML report (static file) and local server mode
# --------------------------------------------------------------------------
def _vals_payload(M, n, es):
    a, b = M.series(n, 'a'), M.series(n, 'b')
    if a is None:
        return None
    return [[_r(v) for v in a[es:M.we + 1]], [_r(v) for v in b[es:M.we + 1]]]


def _dec_payload(M, n):
    d = M.dyn(n, periods=True)
    if not d:
        return None
    if 'error' in d:
        return {'err': d['error']}
    cols = [[k, [_r(v, 5) for v in d['contrib'][k]]]
            for k in sorted(d['contrib'], key=lambda k: -abs(d['end'][k]))]
    return {'gap': [_r(v, 6) for v in d['gap']], 'cols': cols,
            'af': [_r(v, 5) for v in d['af']] if d['af'] else None,
            'afn': d['af_name'], 'int': [_r(v, 5) for v in d['inter']], 'iend': d['iend'],
            'afs': _r(d['af_share'], 3)}


def _af_rows(M):
    rows = []
    for n in M.ranked():
        if n not in M.eqs or not M.eqs[n].evaluable:
            continue
        d = M.dyn(n)
        if not d or 'error' in d:
            continue
        afn = d['af_name']
        afc = None
        st = M.stats.get(afn) if afn else None
        if st and st['differs']:
            afc = [_r(st['dmax']), M.labels[st['tmax']]]
        rows.append([n, M.desc.get(n, ''), _r(d['gap_end']), _r(M.stats[n]['pend'], 4),
                     _r(d['formula_share_end'], 3), _r(d['af_share_end'], 3) if d['af_changed'] else 0,
                     afn or '', afc, _r(d['af_share'], 3)])
    return rows


def build_data(M, args, eqstats, serve=False):
    t0 = time.time()
    ranked = M.ranked()
    labels = M.labels
    es = max(0, M.ws - args.context)
    we = M.we
    per = labels[es:we + 1]

    def log(msg):
        if M.verbose:
            print(msg, file=sys.stderr)

    focus = [_norm(M, f) for f in (args.focus or [])]
    todo = [n for n in ranked if n in M.eqs and M.eqs[n].evaluable][:args.max_decomp]
    for i, n in enumerate(todo):
        M.dyn(n)
        if M.verbose and i and i % 2000 == 0:
            log('  analysed %d / %d equations ...' % (i, len(todo)))
    log('  driver contributions: %d equations in %.1fs' % (len(todo), time.time() - t0))

    nodes = sorted(set(M.a) | set(M.b) | set(M.eqs) | set(M.children))
    nidx = {n: i for i, n in enumerate(nodes)}
    par = [[nidx[p] for p in M.parents.get(n, []) if p in nidx] for n in nodes]

    rows = []
    for n in ranked:
        s = M.stats[n]
        rows.append([n, M.desc.get(n, ''), M.geo(n), _r(s['score'], 4), _r(s['dmax']), _r(s['pmax'], 4),
                     _r(s['dend']), _r(s['pend'], 4), labels[s['first']] if s['first'] is not None else '',
                     M.kind(n)])
    entry = [[n, M.kind(n)] for n in M.entry_points()]

    emb, seen = [], set()

    def add(n):
        if n not in seen and (n in M.a or n in M.b):
            seen.add(n)
            emb.append(n)
    for n in focus:
        add(n)
    for n, _ in entry[:args.embed]:
        add(n)
    for n in ranked[:args.embed]:
        add(n)
    for n in list(emb)[:args.embed]:
        d = M.dyn(n)
        if d and 'error' not in d:
            for k, _ in sorted(d['end'].items(), key=lambda kv: -abs(kv[1]))[:8]:
                add(k)
            if d['af_name']:
                add(d['af_name'])
    vals = {n: _vals_payload(M, n, es) for n in emb}
    dec = {}
    for n in emb:
        p = _dec_payload(M, n)
        if p:
            dec[n] = p
    esh, afsh = {}, {}
    for n, d in M._dd.items():
        if not d or 'error' in d or n not in nidx:
            continue
        esh[str(nidx[n])] = [[nidx[k], _r(v, 3)] for k, v in d['shares'].items() if k in nidx and v is not None]
        if d['af_changed'] and d['af_share_end'] is not None:
            afsh[str(nidx[n])] = _r(d['af_share_end'], 3)
    eqt = {n: M.eqs[n].text[:700] for n in nodes if n in M.eqs and (n in seen or n in M.differs)}
    why = {n: M.eqs[n].why for n in eqt if M.eqs[n].why}

    geo_rows, stem_rows = [], []
    with_sep = sum(1 for n in M.common if M.geo(n))
    if with_sep >= 0.3 * max(1, len(M.common)):
        g_all, g_diff = Counter(), defaultdict(list)
        s_all, s_diff = defaultdict(set), defaultdict(list)
        for n in M.common:
            g_all[M.geo(n)] += 1
            s_all[M.stem(n)].add(M.geo(n))
        for n in M.differs:
            g_diff[M.geo(n)].append(n)
            s_diff[M.stem(n)].append(n)

        def med(xs):
            xs = sorted(xs)
            return xs[len(xs) // 2] if xs else None
        for g, ns in g_diff.items():
            ns.sort(key=lambda n: -M.stats[n]['score'])
            geo_rows.append([g, len(ns), g_all[g], _r(med([M.stats[n]['score'] for n in ns]), 3), ns[:5]])
        geo_rows.sort(key=lambda r: (-r[1], -(r[3] or 0)))
        for st, ns in s_diff.items():
            ns.sort(key=lambda n: -M.stats[n]['score'])
            pm = [abs(M.stats[n]['pend']) for n in ns if M.stats[n]['pend'] == M.stats[n]['pend']]
            stem_rows.append([st, len(ns), len(s_all[st]), _r(med(pm), 3), _r(M.stats[ns[0]]['score'], 3), ns[:5],
                              M.desc.get(ns[0], '')])
        stem_rows.sort(key=lambda r: (-r[1], -(r[4] or 0)))

    data = {
        'meta': {'a': M.A.label, 'b': M.B.label, 'ws': labels[M.ws], 'we': labels[M.we],
                 'first': labels[0], 'last': labels[-1], 'ppy': M.ppy,
                 'n_common': len(M.common), 'n_diff': len(M.differs),
                 'n_only_a': len(M.only_a), 'n_only_b': len(M.only_b),
                 'eqs': dict(eqstats) if eqstats else None, 'has_eqs': M.has_eqs,
                 'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                 'adj_thr': M.adj_threshold, 'embedded': len(emb), 'serve': serve,
                 'af_suffix': ','.join(M.af_suffixes), 'n_af': len(M.af_series)},
        'per': per, 'wsi': M.ws - es,
        'rows': rows, 'entry': entry, 'nodes': nodes, 'par': par,
        'vals': vals, 'dec': dec, 'esh': esh, 'afsh': afsh, 'eqt': eqt, 'why': why,
        'afmap': {k: v for k, v in M.af.items()},
        'afrows': _af_rows(M) if M.has_eqs else [],
        'geo': geo_rows, 'stem': stem_rows,
        'only_a': M.only_a[:500], 'only_b': M.only_b[:500],
        'desc': {n: M.desc[n] for n in nodes if n in M.desc and n not in M.differs},
    }
    log('  report data built in %.1fs (%d series with full detail)' % (time.time() - t0, len(emb)))
    return data, es


def render_html(M, data):
    js = json.dumps(data, separators=(',', ':'), allow_nan=False).replace('</', '<\\/')
    return HTML_TEMPLATE.replace('__TITLE__', _esc('%s vs %s' % (M.A.label, M.B.label))).replace('__DATA__', js)


def build_report(M, args, eqstats):
    data, _ = build_data(M, args, eqstats)
    out = render_html(M, data)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write(out)
    if M.verbose:
        print('Report written: %s (%.1f MB)' % (args.out, os.path.getsize(args.out) / 1e6), file=sys.stderr)


def _impact_payload(M, x, y, shock_s):
    shock = parse_shock(shock_s)
    if y and y != x and y not in M.descendants(x):
        return {'none': True}
    eff, info = M.impact(x, y or None, shock)
    rng = range(M.ws, M.we + 1)
    out = {'per': [M.labels[t] for t in rng], 'nodes': info['nodes'], 'blocked': info['blocked'][:20],
           'shock': shock_s or ''}
    if y:
        A, Bv = M.series(y, 'a'), M.series(y, 'b')
        out['y'] = [_r(v) for v in eff.get(y, [])]
        out['gap'] = [_r(Bv[t] - A[t]) for t in rng]
        out['ya'] = [_r(A[t]) for t in rng]
    top = []
    for n, e in eff.items():
        if n == x or n == y:
            continue
        a = M.series(n, 'a')[M.we]
        v = e[-1]
        if v == v and a:
            top.append((abs(v / a), n, v, v / abs(a) * 100))
    top.sort(reverse=True)
    out['top'] = [[n, _r(v), _r(p, 4)] for _, n, v, p in top[:40]]
    return out


def _attr_payload(M, y):
    res, joint, gap = M.attribute(y)
    i = max(j for j, g in enumerate(gap) if g == g)
    rows = [[s, kind, M.af.get(s) if kind == 'addfactor' else None, [_r(v, 5) for v in e]]
            for kind, s, e in sorted(res, key=lambda r: -abs(r[2][i]) if r[2][i] == r[2][i] else 0)
            if abs(e[i]) > 1e-9 * max(1.0, abs(gap[i]))]
    return {'per': [M.labels[t] for t in range(M.ws, M.we + 1)], 'iend': i, 'gap': [_r(v) for v in gap],
            'joint': [_r(v) for v in joint], 'rows': rows[:60]}


def cmd_serve(M, args, eqstats):
    from http.server import HTTPServer, BaseHTTPRequestHandler
    from urllib.parse import urlparse, parse_qs
    import webbrowser
    data, es = build_data(M, args, eqstats, serve=True)
    page = render_html(M, data).encode('utf-8')

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, body, ctype='application/json'):
            self.send_response(200)
            self.send_header('Content-Type', ctype + '; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            u = urlparse(self.path)
            q = {k: v[0] for k, v in parse_qs(u.query).items()}
            try:
                if u.path in ('/', '/index.html'):
                    return self._send(page, 'text/html')
                if u.path == '/api/detail':
                    n = _norm(M, q['n'])
                    out = {'n': n, 'vals': _vals_payload(M, n, es), 'dec': _dec_payload(M, n),
                           'eqt': M.eqs[n].text[:700] if n in M.eqs else None,
                           'why': M.eqs[n].why if n in M.eqs else None, 'desc': M.desc.get(n, '')}
                elif u.path == '/api/impact':
                    out = _impact_payload(M, _norm(M, q['x']), _norm(M, q['y']) if q.get('y') else None,
                                          q.get('shock'))
                elif u.path == '/api/attr':
                    out = _attr_payload(M, _norm(M, q['y']))
                else:
                    self.send_error(404)
                    return
            except SystemExit as e:
                out = {'error': str(e)}
            except Exception as e:  # keep the server alive
                out = {'error': '%s: %s' % (type(e).__name__, e)}
            self._send(json.dumps(out, allow_nan=False, default=str).encode('utf-8'))

    srv = HTTPServer(('127.0.0.1', args.port), H)
    url = 'http://127.0.0.1:%d/' % args.port
    print('Serving %s  (Ctrl+C to stop)' % url, file=sys.stderr)
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;'))


HTML_TEMPLATE = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Scenario comparison: __TITLE__</title>
<style>
:root{--bg:#fff;--fg:#1d2433;--mut:#667085;--line:#e4e7ec;--card:#f7f8fa;--a:#2563eb;--b:#d9480f;--pos:#0f9d58;--neg:#d92d20;--hl:#fff5d6}
@media (prefers-color-scheme:dark){:root{--bg:#0f1218;--fg:#e6e9ef;--mut:#98a2b3;--line:#2a2f3a;--card:#171b23;--a:#6c9cff;--b:#ff8a4c;--pos:#3ccb7f;--neg:#ff6b5e;--hl:#3a3216}}
*{box-sizing:border-box}
body{margin:0;font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--fg)}
header{padding:16px 20px 8px;border-bottom:1px solid var(--line)}
h1{font-size:19px;margin:0 0 4px}
h2{font-size:16px;margin:18px 0 8px}
.mut{color:var(--mut)}
.small{font-size:12px}
.wrap{padding:12px 20px 40px;max-width:1300px}
nav{display:flex;gap:4px;flex-wrap:wrap;padding:8px 20px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--bg);z-index:5}
nav button{border:1px solid var(--line);background:var(--card);color:var(--fg);padding:6px 12px;border-radius:16px;cursor:pointer;font:inherit}
nav button.on{background:var(--fg);color:var(--bg);border-color:var(--fg)}
.kpis{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:8px 12px;min-width:120px}
.kpi b{display:block;font-size:20px}
.ctl{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0}
input,select{font:inherit;padding:6px 8px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg)}
input[type=text]{min-width:220px}
table{border-collapse:collapse;width:100%}
th,td{padding:5px 8px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:var(--bg);cursor:pointer;white-space:nowrap;font-weight:600;font-size:12px;color:var(--mut)}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr.cl{cursor:pointer}
tr.cl:hover{background:var(--card)}
.tw{overflow-x:auto}
a.v{color:var(--a);cursor:pointer;text-decoration:none;font-weight:500}
a.v:hover{text-decoration:underline}
.pos{color:var(--pos)}.neg{color:var(--neg)}
.badge{display:inline-block;font-size:11px;padding:1px 7px;border-radius:10px;border:1px solid var(--line);background:var(--card);white-space:nowrap}
.bX{border-color:#c4320a;color:#c4320a}.bF{border-color:#7a5af8;color:#7a5af8}.bA{border-color:#b54708;color:#b54708}.bE{color:var(--mut)}.bN{color:var(--mut);border-style:dashed}
#det{position:fixed;top:0;right:0;bottom:0;width:min(760px,100%);background:var(--bg);border-left:1px solid var(--line);box-shadow:-8px 0 24px rgba(0,0,0,.12);overflow-y:auto;transform:translateX(105%);transition:transform .2s;z-index:20;padding:14px 18px 60px}
#det.open{transform:none}
.dh{display:flex;justify-content:space-between;align-items:center;gap:8px}
.btn{border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:8px;padding:5px 10px;cursor:pointer;font:inherit}
pre.eq{white-space:pre-wrap;word-break:break-word;background:var(--card);border:1px solid var(--line);padding:8px;border-radius:8px;font-size:12px;margin:6px 0}
svg text{fill:var(--mut);font-size:11px}
.bars td{padding:3px 6px}
.bar{height:12px;border-radius:2px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:10px 0}
.hint{font-size:12px;color:var(--mut);margin:4px 0 8px}
.path{margin:8px 0;padding:8px 10px;border:1px solid var(--line);border-radius:10px}
.step{display:inline-block;margin:2px 0}
.arrow{color:var(--mut);margin:0 6px}
.hide{display:none}
</style></head><body>
<header>
<h1>Scenario comparison</h1>
<div id="sub" class="mut"></div>
</header>
<nav id="tabs"></nav>
<div class="wrap">
<section id="t-sum"></section>
<section id="t-top" class="hide"></section>
<section id="t-entry" class="hide"></section>
<section id="t-afx" class="hide"></section>
<section id="t-stem" class="hide"></section>
<section id="t-geo" class="hide"></section>
<section id="t-path" class="hide"></section>
</div>
<div id="det"></div>
<script id="data" type="application/json">__DATA__</script>
<script>
"use strict";
const D=JSON.parse(document.getElementById('data').textContent);
const M=D.meta, NI={}; D.nodes.forEach((n,i)=>NI[n]=i);
const CH=D.nodes.map(()=>[]); D.par.forEach((ps,i)=>ps.forEach(p=>CH[p].push(i)));
const ROW={}; D.rows.forEach(r=>ROW[r[0]]=r);
const KT={X:'assumption',F:'add-factor',A:'mostly add-factor',E:'formula',N:'not evaluable','-':''};
const KL={X:'Exogenous: no equation, set directly by the scenario',F:'An add-factor series (name ends in '+D.meta.af_suffix+')',A:'Has an equation, but most of its change comes from its add-factor / residual, not from its drivers',E:'Change is explained by its equation (formula) drivers',N:'Has an equation that could not be evaluated'};
const SERVE=D.meta.serve&&location.protocol.startsWith('http');
const $=(s,e=document)=>e.querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function fmt(v){if(v==null||isNaN(v))return'–';if(v===0)return'0';const a=Math.abs(v);if(a>=1e4)return v.toLocaleString(undefined,{maximumFractionDigits:0});return Number(v.toPrecision(4)).toString()}
function pct(v){if(v==null||isNaN(v))return'–';return(v>0?'+':'')+v.toFixed(2)+'%'}
function cls(v){return v>0?'pos':v<0?'neg':''}
function badge(k){return k&&k!=='-'?`<span class="badge b${k}" title="${esc(KL[k]||'')}">${KT[k]}</span>`:''}
function vlink(n){return `<a class="v" data-v="${esc(n)}">${esc(n)}</a>`}
function desc(n){const r=ROW[n];return r?r[1]:(D.desc[n]||'')}
document.addEventListener('click',e=>{const a=e.target.closest('[data-v]');if(a){e.preventDefault();show(a.dataset.v)}});

// ---------- tabs
const TABS=[['sum','Overview'],['top','Key changes'],['entry','Where B starts'],['afx','Formula vs add-factor'],['stem','By concept'],['geo','By geography'],['path','Impact X → Y']];
$('#tabs').innerHTML=TABS.filter(t=>t[0]==='afx'?D.afrows.length:(t[0]!=='stem'&&t[0]!=='geo')||D[t[0]].length).map(t=>`<button data-t="${t[0]}">${t[1]}</button>`).join('');
function tab(t){document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('on',b.dataset.t===t));TABS.forEach(x=>$('#t-'+x[0]).classList.toggle('hide',x[0]!==t))}
$('#tabs').addEventListener('click',e=>{if(e.target.dataset.t)tab(e.target.dataset.t)});
$('#sub').innerHTML=`<b>A</b> = ${esc(M.a)} &nbsp; vs &nbsp; <b>B</b> = ${esc(M.b)} &nbsp;·&nbsp; window ${M.ws}–${M.we} &nbsp;·&nbsp; generated ${M.generated}`;

// ---------- overview
(function(){
 const kinds={};D.rows.forEach(r=>kinds[r[9]]=(kinds[r[9]]||0)+1);
 let h=`<div class="kpis"><div class="kpi"><span class="mut small">series in both</span><b>${M.n_common.toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">differ between A and B</span><b>${M.n_diff.toLocaleString()}</b></div>`;
 if(M.has_eqs){h+=`<div class="kpi"><span class="mut small">changed assumptions (no equation)</span><b>${(kinds.X||0).toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">add-factor series changed</span><b>${(kinds.F||0).toLocaleString()}</b></div>
 <div class="kpi"><span class="mut small">variables driven mostly by add-factors</span><b>${(kinds.A||0).toLocaleString()}</b></div>`}
 h+=`</div><p class="mut small">Comparison window ${M.ws}–${M.we} (data ${M.first}–${M.last}). ${M.n_only_a?M.n_only_a+' series only in A. ':''}${M.n_only_b?M.n_only_b+' series only in B. ':''}</p>`;
 h+=`<h2>Biggest changes</h2><div class="hint">Score = largest gap between B and A, divided by the series' typical period-to-period move in A, so series in different units can be ranked together.</div>`;
 h+=tableHTML(D.rows.slice(0,15));
 if(D.entry.length){h+=`<h2>Where scenario B departs from A first</h2><div class="hint">Assumptions (no equation), add-factor series that changed, and variables whose gap comes mostly from their add-factor rather than their formula. These are the shocks the scenario is built on — everything else follows through the equations.</div>`;
  h+=tableHTML(D.entry.slice(0,15).map(e=>ROW[e[0]]).filter(Boolean),'first')}
 if(M.eqs){h+=`<p class="mut small">Equations loaded: ${esc(Object.entries(M.eqs).map(([k,v])=>k+' '+v).join(', '))}</p>`}
 $('#t-sum').innerHTML=h;
})();

function tableHTML(rows,sortHint){
 return `<div class="tw"><table><thead><tr><th>Series</th><th>Description</th><th class="n">Score</th><th class="n">Max gap</th><th class="n">Max gap %</th><th class="n">Gap at end %</th><th>Diverges</th><th>Type</th></tr></thead><tbody>${rows.map(r=>`<tr class="cl" data-v="${esc(r[0])}"><td><a class="v">${esc(r[0])}</a></td><td class="small">${esc(r[1]).slice(0,80)}</td><td class="n">${fmt(r[3])}</td><td class="n ${cls(r[4])}">${fmt(r[4])}</td><td class="n ${cls(r[5])}">${pct(r[5])}</td><td class="n ${cls(r[7])}">${pct(r[7])}</td><td>${esc(r[8])}</td><td>${badge(r[9])}</td></tr>`).join('')}</tbody></table></div>`}

// ---------- key changes (filterable)
(function(){
 const geos=[...new Set(D.rows.map(r=>r[2]).filter(Boolean))].sort();
 let h=`<div class="ctl"><input type="text" id="q" placeholder="Search name or description (regex ok)">
 ${geos.length>1?`<select id="g"><option value="">All geographies</option>${geos.map(g=>`<option>${esc(g)}</option>`).join('')}</select>`:''}
 ${M.has_eqs?`<select id="k"><option value="">All types</option><option value="X">Assumptions</option><option value="A">Adjusted equations</option><option value="E">Explained by equation</option><option value="N">Not evaluable</option></select>`:''}
 <select id="dir"><option value="">Up or down</option><option value="u">B above A</option><option value="d">B below A</option></select></div>
 <div id="cnt" class="hint"></div><div class="tw"><table id="tt"><thead><tr>
 <th data-c="0">Series</th><th data-c="1">Description</th><th data-c="3" class="n">Score ▾</th><th data-c="4" class="n">Max gap</th><th data-c="5" class="n">Max gap %</th><th data-c="7" class="n">Gap at end %</th><th data-c="8">Diverges</th><th data-c="9">Type</th></tr></thead><tbody></tbody></table></div>`;
 $('#t-top').innerHTML=h;
 let sc=3,sd=-1;
 function render(){
  const q=$('#q').value.trim();let rx=null;try{rx=q?new RegExp(q,'i'):null}catch(e){rx=new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'i')}
  const g=$('#g')?$('#g').value:'',k=$('#k')?$('#k').value:'',dr=$('#dir').value;
  let rs=D.rows.filter(r=>(!rx||rx.test(r[0])||rx.test(r[1]))&&(!g||r[2]===g)&&(!k||r[9]===k)&&(!dr||(dr==='u'?r[4]>0:r[4]<0)));
  rs.sort((x,y)=>{let a=x[sc],b=y[sc];if(typeof a==='number'||typeof b==='number'){a=a==null?-Infinity:Math.abs(a);b=b==null?-Infinity:Math.abs(b)}return a<b?-sd:a>b?sd:0});
  $('#cnt').textContent=rs.length.toLocaleString()+' series'+(rs.length>400?' (showing first 400)':'')+' · numeric columns sort by size, ignoring sign';
  $('#tt tbody').innerHTML=rs.slice(0,400).map(r=>`<tr class="cl" data-v="${esc(r[0])}"><td><a class="v">${esc(r[0])}</a></td><td class="small">${esc(r[1]).slice(0,80)}</td><td class="n">${fmt(r[3])}</td><td class="n ${cls(r[4])}">${fmt(r[4])}</td><td class="n ${cls(r[5])}">${pct(r[5])}</td><td class="n ${cls(r[7])}">${pct(r[7])}</td><td>${esc(r[8])}</td><td>${badge(r[9])}</td></tr>`).join('')}
 ['q','g','k','dir'].forEach(id=>{const e=$('#'+id);if(e)e.addEventListener('input',render)});
 $('#tt thead').addEventListener('click',e=>{const c=e.target.dataset.c;if(c==null)return;const ci=+c;if(ci===sc)sd=-sd;else{sc=ci;sd=(ci<=2||ci>=8)?1:-1}document.querySelectorAll('#tt th').forEach(t=>t.textContent=t.textContent.replace(/ [▾▴]$/,''));e.target.textContent+=sd<0?' ▾':' ▴';render()});
 render();
})();

// ---------- entry points
(function(){
 let h=`<div class="hint">${M.has_eqs?'Ordered by when each first diverges. <b>Assumptions</b> have no equation in the model, so they are set by the scenario directly. <b>Add-factor</b> series that changed, and variables whose gap comes mostly (≥'+Math.round(M.adj_thr*100)+'%) from their add-factor rather than their formula. Start here to see what the scenario actually changed.':'No equations were loaded, so assumptions cannot be told apart from model outcomes. Listed instead: series ordered by when they first diverge.'}</div>`;
 const rows=M.has_eqs?D.entry.map(e=>ROW[e[0]]).filter(Boolean):D.rows.slice().sort((a,b)=>a[8]<b[8]?-1:a[8]>b[8]?1:b[3]-a[3]).slice(0,300);
 h+=`<div class="hint">${rows.length.toLocaleString()} series</div>`+tableHTML(rows.slice(0,500));
 $('#t-entry').innerHTML=h;
})();

// ---------- formula vs add-factor
(function(){if(!D.afrows.length)return;
 $('#t-afx').innerHTML=`<div class="hint">For every variable with an equation: how much of its gap (B−A) at ${M.we} comes from its <b>formula</b> (its drivers changed) versus its <b>add-factor</b> (${esc(M.af_suffix)} series or residual changed). The equation is the same in both scenarios, so any gap the formula does not explain comes from the add-factor.</div>
 <div class="ctl"><input type="text" id="aq" placeholder="Search"><select id="am"><option value="">All</option><option value="af" selected>Add-factor changed</option><option value="maf">Mostly add-factor (≥50%)</option><option value="f">Formula only</option></select></div><div id="acnt" class="hint"></div>
 <div class="tw"><table><thead><tr><th>Series</th><th>Description</th><th class="n">Gap at end</th><th class="n">Gap %</th><th class="n">From formula</th><th class="n">From add-factor</th><th>Add-factor series (largest B−A change)</th></tr></thead><tbody id="atb"></tbody></table></div>`;
 const sh=v=>v==null?'–':Math.round(v*100)+'%';
 const r=()=>{const q=$('#aq').value.trim().toUpperCase(),m=$('#am').value;
  const rs=D.afrows.filter(x=>(!q||x[0].includes(q)||x[1].toUpperCase().includes(q))&&(!m||(m==='af'?x[5]!==0:m==='maf'?Math.abs(x[5]||0)>=.5:x[5]===0)));
  $('#acnt').textContent=rs.length.toLocaleString()+' variables'+(rs.length>500?' (first 500)':'');
  $('#atb').innerHTML=rs.slice(0,500).map(x=>`<tr class="cl" data-v="${esc(x[0])}"><td><a class="v">${esc(x[0])}</a></td><td class="small">${esc(x[1]).slice(0,60)}</td><td class="n ${cls(x[2])}">${fmt(x[2])}</td><td class="n ${cls(x[3])}">${pct(x[3])}</td><td class="n">${sh(x[4])}</td><td class="n"><b>${sh(x[5])}</b></td><td>${x[6]?vlink(x[6])+(x[7]?` <span class="small ${cls(x[7][0])}">(${fmt(x[7][0])} in ${esc(x[7][1])})</span>`:' <span class="small mut">(unchanged)</span>'):'<span class="mut small">residual (no '+esc(M.af_suffix)+' series)</span>'}</td></tr>`).join('')};
 $('#aq').addEventListener('input',r);$('#am').addEventListener('input',r);r()})();

// ---------- by concept / geography
(function(){
 if(D.stem.length){$('#t-stem').innerHTML=`<div class="hint">Same concept across geographies (name before the geography suffix). Shows how widespread a change is.</div><div class="tw"><table><thead><tr><th>Concept</th><th>Description</th><th class="n">Geos changed</th><th class="n">of</th><th class="n">Median |gap at end %|</th><th class="n">Max score</th><th>Largest</th></tr></thead><tbody>${D.stem.slice(0,1500).map(r=>`<tr><td><b>${esc(r[0])}</b></td><td class="small">${esc(r[6]).slice(0,70)}</td><td class="n">${r[1]}</td><td class="n mut">${r[2]}</td><td class="n">${r[3]==null?'–':r[3].toFixed(2)+'%'}</td><td class="n">${fmt(r[4])}</td><td class="small">${r[5].map(vlink).join(', ')}</td></tr>`).join('')}</tbody></table></div>`}
 if(D.geo.length){$('#t-geo').innerHTML=`<div class="tw"><table><thead><tr><th>Geography</th><th class="n">Series changed</th><th class="n">of</th><th class="n">Median score</th><th>Largest changes</th></tr></thead><tbody>${D.geo.map(r=>`<tr><td><b>${esc(r[0]||'(none)')}</b></td><td class="n">${r[1]}</td><td class="n mut">${r[2]}</td><td class="n">${fmt(r[3])}</td><td class="small">${r[4].map(vlink).join(', ')}</td></tr>`).join('')}</tbody></table></div>`}
})();

// ---------- path tracing
function esh(child,parent){const l=D.esh[NI[child]];if(!l)return null;const p=NI[parent];for(const e of l)if(e[0]===p)return e[1];return null}
function findPaths(x,y,k){const xi=NI[x],yi=NI[y];if(xi==null||yi==null)return[];const dist=new Map([[xi,0]]),q=[xi];
 while(q.length){const u=q.shift();if(u===yi)continue;for(const v of CH[u])if(!dist.has(v)){dist.set(v,dist.get(u)+1);q.push(v)}}
 if(!dist.has(yi))return[];const out=[];(function back(n,acc){if(out.length>=k)return;if(n===xi){out.push([xi,...acc.slice().reverse()]);return}for(const p of D.par[n])if(dist.get(p)===dist.get(n)-1)back(p,[...acc,n])})(yi,[]);return out.map(p=>p.map(i=>D.nodes[i]))}
function strongest(x,y){const xi=NI[x],yi=NI[y];if(xi==null||yi==null||xi===yi)return null;
 const best=new Map([[xi,0]]),prev=new Map(),hops=new Map([[xi,0]]),h=[[0,xi]];
 const push=e=>{h.push(e);let i=h.length-1;while(i>0){const p=(i-1)>>1;if(h[p][0]<=h[i][0])break;[h[p],h[i]]=[h[i],h[p]];i=p}};
 const pop=()=>{const t=h[0],l=h.pop();if(h.length){h[0]=l;let i=0;for(;;){let a=2*i+1,b=a+1,m=i;if(a<h.length&&h[a][0]<h[m][0])m=a;if(b<h.length&&h[b][0]<h[m][0])m=b;if(m===i)break;[h[m],h[i]]=[h[i],h[m]];i=m}}return t};
 while(h.length){const [c,u]=pop();if(c>(best.get(u)??Infinity))continue;if(u===yi)break;if(hops.get(u)>=15)continue;
  for(const v of CH[u]){if(v!==yi&&!ROW[D.nodes[v]])continue;const sh=esh(D.nodes[v],D.nodes[u]);const w=sh?-Math.log(Math.min(1,Math.abs(sh))):7;
   if(c+w<(best.get(v)??Infinity)){best.set(v,c+w);prev.set(v,u);hops.set(v,hops.get(u)+1);push([c+w,v])}}}
 if(!best.has(yi))return null;const p=[yi];while(p[p.length-1]!==xi)p.push(prev.get(p[p.length-1]));return p.reverse().map(i=>D.nodes[i])}
function prodOf(p){let pr=1;for(let j=1;j<p.length;j++){const sh=esh(p[j],p[j-1]);if(sh==null)return null;pr*=sh}return pr}
function rankedPaths(x,y,k){const ps=findPaths(x,y,40),sp=strongest(x,y);if(sp&&!ps.some(p=>p.join()===sp.join()))ps.push(sp);
 ps.sort((a,b)=>{const pa=prodOf(a),pb=prodOf(b);return (pb==null?-1:Math.abs(pb))-(pa==null?-1:Math.abs(pa))});return ps.slice(0,k)}
function gapTxt(n){const r=ROW[n];if(!r)return'<span class="mut">no change</span>';if(r[7]==null)return `<span class="${cls(r[4])} small">max gap ${fmt(r[4])} (${esc(r[8])})</span>`;return `<span class="${cls(r[7])}">${pct(r[7])}</span>`}
function pathHTML(x,y){
 const ps=rankedPaths(x,y,6);
 if(!ps.length){const rev=findPaths(y,x,1);return `<div class="card">${esc(x)} does not feed into ${esc(y)} through the loaded equations.${rev.length?` The reverse holds: ${rev[0].map(vlink).join(' → ')}.`:''}</div>`}
 return ps.map((p,i)=>{let prod=1,ok=true;const st=p.map((n,j)=>{let s='';if(j>0){const sh=esh(n,p[j-1]);if(sh==null){ok=false;s='<span class="mut small">share n/a</span>'}else{prod*=sh;s=`<span class="small">${Math.round(sh*100)}% of its gap from ${esc(p[j-1])}</span>`}}return `<span class="step">${j?'<span class="arrow">→</span>':''}${vlink(n)} ${gapTxt(n)} ${s}</span>`}).join('<br>');
  return `<div class="path"><b>Path ${i+1}</b>${i===0&&ps.length>1?' <span class="small mut">(paths ordered by how much of the change they carry)</span>':''}<br>${st}${ok&&p.length>1?`<div class="hint">≈ ${Math.round(prod*100)}% of ${esc(y)}'s change travels along this channel (product of the link shares, approximate).</div>`:''}</div>`}).join('')}
function upHTML(x,depth){const xi=NI[x];const dist=new Map([[xi,0]]),nxt=new Map(),q=[xi];
 while(q.length){const u=q.shift();if(dist.get(u)>=depth)continue;for(const p of D.par[u])if(!dist.has(p)){dist.set(p,dist.get(u)+1);nxt.set(p,u);q.push(p)}}
 const rs=[...dist.keys()].filter(i=>i!==xi&&ROW[D.nodes[i]]).map(i=>[D.nodes[i],dist.get(i),D.nodes[nxt.get(i)]]);
 [...dist.keys()].forEach(i=>{const a=D.afsh[i];if(a!=null&&Math.abs(a)>=.005)rs.push(['__AF__'+D.nodes[i],dist.get(i)+1,D.nodes[i],a])});
 rs.sort((a,b)=>a[1]-b[1]||Math.abs(b[3]??esh(b[2],b[0])??0)-Math.abs(a[3]??esh(a[2],a[0])??0));
 if(!D.par[xi].length)return `<div class="hint">${esc(x)} has no equation either — it is an assumption set directly by the scenario.</div>`;
 if(!rs.length)return `<div class="hint">None of the variables driving ${esc(x)} (within ${depth} steps) differ between A and B.</div>`;
 return `<div class="tw"><table><thead><tr><th>Driver</th><th class="n">Steps back</th><th>Feeds</th><th class="n">Gap at end %</th><th class="n">Share of that variable's gap</th></tr></thead><tbody>${rs.slice(0,300).map(r=>{if(r[0].startsWith('__AF__')){const t=r[2],af=D.afmap[t];return `<tr><td><i>add-factor</i> ${af?vlink(af):'<span class="small mut">(residual)</span>'}</td><td class="n">${r[1]}</td><td>${vlink(t)}</td><td class="n">${af?gapTxt(af):''}</td><td class="n"><b>${Math.round(r[3]*100)}%</b></td></tr>`}const sh=esh(r[2],r[0]);return `<tr><td>${vlink(r[0])}</td><td class="n">${r[1]}</td><td>${vlink(r[2])}</td><td class="n">${gapTxt(r[0])}</td><td class="n">${sh==null?'–':Math.round(sh*100)+'%'}</td></tr>`}).join('')}</tbody></table></div>`}
function downHTML(x,depth){const xi=NI[x];if(xi==null)return'';
 if(!CH[xi].length)return `<div class="card"><b>${esc(x)} is an end point of the model:</b> no other equation uses it, so a change in it does not flow anywhere. (A lag of itself in its own equation doesn't count.)</div><h2>What drives ${esc(x)} instead</h2>`+upHTML(x,depth);const dist=new Map([[xi,0]]),prev=new Map(),q=[xi];
 while(q.length){const u=q.shift();if(dist.get(u)>=depth)continue;for(const v of CH[u])if(!dist.has(v)){dist.set(v,dist.get(u)+1);prev.set(v,u);q.push(v)}}
 const rs=[...dist.keys()].filter(i=>i!==xi&&ROW[D.nodes[i]]).map(i=>[D.nodes[i],dist.get(i),D.nodes[prev.get(i)]]);
 rs.sort((a,b)=>a[1]-b[1]||ROW[b[0]][3]-ROW[a[0]][3]);
 if(!rs.length)return `<div class="hint">Nothing that ${esc(x)} feeds within ${depth} steps differs between scenarios.</div>`;
 return `<div class="hint">${rs.length} variables within ${depth} steps of ${esc(x)} differ between A and B.</div><div class="tw"><table><thead><tr><th>Variable</th><th class="n">Steps</th><th>Via</th><th class="n">Gap at end %</th><th class="n">Share from via</th></tr></thead><tbody>${rs.slice(0,300).map(r=>{const sh=esh(r[0],r[2]);return `<tr><td>${vlink(r[0])}</td><td class="n">${r[1]}</td><td>${vlink(r[2])}</td><td class="n">${gapTxt(r[0])}</td><td class="n">${sh==null?'–':Math.round(sh*100)+'%'}</td></tr>`}).join('')}</tbody></table></div>`}
(function(){
 $('#t-path').innerHTML=`<div class="hint">Can X affect Y, and by how much? ${SERVE?'The equations between X and Y are re-solved with <b>only X changed</b>, which gives the exact effect.':'The channels below show, at each link, what share of the next variable\'s gap comes from the previous one. <b>For the exact effect</b>, run <code>python scenario_compare.py serve …</code> (or the <code>impact X Y</code> command), which re-solves the equations with only X changed.'}</div>
 <div class="ctl"><input type="text" id="px" placeholder="X (e.g. FOIL)"><input type="text" id="py" placeholder="Y (optional)">${SERVE?'<select id="ps"><option value="">X moves as in scenario B</option><option value="1%">X +1%</option><option value="-1%">X −1%</option><option value="1">X +1 unit</option><option value="0.25">X +0.25 unit</option></select>':''}<select id="pd"><option value="2">2 steps</option><option value="3" selected>3 steps</option><option value="5">5 steps</option></select><button class="btn" id="pgo">Go</button></div><div id="pimp"></div><div id="pout"></div>`;
 const go=()=>{const x=$('#px').value.trim().toUpperCase(),y=$('#py').value.trim().toUpperCase();$('#pimp').innerHTML='';
  if(NI[x]==null){$('#pout').innerHTML=`<div class="card">${esc(x)} not found. ${sugg(x)}</div>`;return}
  if(y&&NI[y]==null){$('#pout').innerHTML=`<div class="card">${esc(y)} not found. ${sugg(y)}</div>`;return}
  $('#pout').innerHTML=y?'<h2>Channels</h2>'+pathHTML(x,y):`<h2>What ${esc(x)} feeds</h2>`+downHTML(x,+$('#pd').value);
  if(SERVE)runImpact(x,y,$('#ps').value)};
 $('#pgo').onclick=go;['px','py'].forEach(i=>$('#'+i).addEventListener('keydown',e=>{if(e.key==='Enter')go()}));
})();
function runImpact(x,y,shock){$('#pimp').innerHTML='<div class="hint">Re-solving the equations…</div>';
 fetch(`/api/impact?x=${encodeURIComponent(x)}&y=${encodeURIComponent(y)}&shock=${encodeURIComponent(shock)}`).then(r=>r.json()).then(o=>{
  if(o.error){$('#pimp').innerHTML=`<div class="card">${esc(o.error)}</div>`;return}
  if(o.none){$('#pimp').innerHTML=`<div class="card"><b>No.</b> ${esc(x)} cannot affect ${esc(y)}: no chain of equations links them.</div>`;return}
  const what=shock?`a ${esc(shock)}${shock.endsWith('%')?'':' unit'} change in ${esc(x)}`:`${esc(x)}'s change from A to B`;
  let h='';
  if(y&&o.y){const i=o.y.length-1,v=o.y[i],g=o.gap[i],a=o.ya[i];
   h+=`<div class="card"><b>Effect of ${what} on ${esc(y)}</b> at ${esc(o.per[i])}: <b class="${cls(v)}">${fmt(v)}</b> (${pct(a?100*v/Math.abs(a):null)})${!shock&&g?` = <b>${Math.round(100*v/g)}%</b> of ${esc(y)}'s total gap (${fmt(g)})`:''}. ${o.nodes} equations re-solved.${o.blocked.length?` <span class="mut small">${o.blocked.length} variables in between have no evaluable equation and were held fixed.</span>`:''}</div>`;
   h+=lineChart2(o.y,shock?null:o.gap,o.per,'Effect of '+x,'Total gap B−A')}
  else if(y&&!o.y)h+=`<div class="card">${esc(y)} has no evaluable equation, so the effect cannot be simulated.</div>`;
  if(o.top&&o.top.length)h+=`<h2>${y?'Variables in between':'Most affected variables'} (effect at ${esc(o.per[o.per.length-1])})</h2><div class="tw"><table><thead><tr><th>Variable</th><th class="n">Effect</th><th class="n">%</th></tr></thead><tbody>${o.top.map(r=>`<tr><td>${vlink(r[0])}</td><td class="n ${cls(r[1])}">${fmt(r[1])}</td><td class="n ${cls(r[2])}">${pct(r[2])}</td></tr>`).join('')}</tbody></table></div>`;
  $('#pimp').innerHTML=h})}
function lineChart2(a,b,per,la,lb){const W=640,H=180,L=56,R=8,T=10,B=24,n=per.length;const vs=a.concat(b||[]).concat([0]).filter(v=>v!=null);let lo=Math.min(...vs),hi=Math.max(...vs);if(lo===hi)hi=lo+1;
 const X=i=>L+(W-L-R)*(n>1?i/(n-1):.5),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));const pl=arr=>arr.map((v,i)=>v==null?'':(i&&arr[i-1]!=null?'L':'M')+X(i).toFixed(1)+','+Y(v).toFixed(1)).join('');
 return `<svg viewBox="0 0 ${W} ${H}" width="100%"><line x1="${L}" x2="${W-R}" y1="${Y(0)}" y2="${Y(0)}" stroke="var(--line)"/><text x="${L-6}" y="${Y(hi)+8}" text-anchor="end">${fmt(hi)}</text><text x="${L-6}" y="${Y(lo)}" text-anchor="end">${fmt(lo)}</text><text x="${L}" y="${H-6}">${esc(per[0])}</text><text x="${W-R}" y="${H-6}" text-anchor="end">${esc(per[n-1])}</text>
 ${b?`<path d="${pl(b)}" fill="none" stroke="var(--mut)" stroke-width="2" stroke-dasharray="4,3"/>`:''}<path d="${pl(a)}" fill="none" stroke="var(--b)" stroke-width="2.5"/></svg><div class="small"><span style="color:var(--b)">━ ${esc(la)}</span>${b?` &nbsp; <span class="mut">╍ ${esc(lb)}</span>`:''}</div>`}
function sugg(x){if(!x)return'';const m=D.nodes.filter(n=>n.includes(x)).slice(0,12);return m.length?'Did you mean: '+m.map(vlink).join(', '):''}

// ---------- charts
function lineChart(a,b,per,wsi){
 const W=640,H=200,L=56,R=8,T=10,B=24,n=per.length;const vs=a.concat(b).filter(v=>v!=null);if(!vs.length)return'';
 let lo=Math.min(...vs),hi=Math.max(...vs);if(lo===hi){lo-=1;hi+=1}const pad=(hi-lo)*.06;lo-=pad;hi+=pad;
 const X=i=>L+(W-L-R)*(n>1?i/(n-1):.5),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));
 const pl=arr=>{let d='',pen=false;arr.forEach((v,i)=>{if(v==null){pen=false;return}d+=(pen?'L':'M')+X(i).toFixed(1)+','+Y(v).toFixed(1);pen=true});return d};
 return `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img"><line x1="${L}" x2="${W-R}" y1="${H-B}" y2="${H-B}" stroke="var(--line)"/>
 ${wsi>0?`<line x1="${X(wsi)}" x2="${X(wsi)}" y1="${T}" y2="${H-B}" stroke="var(--line)" stroke-dasharray="3,3"/>`:''}
 <text x="${L-6}" y="${Y(hi-pad)+4}" text-anchor="end">${fmt(hi-pad)}</text><text x="${L-6}" y="${Y(lo+pad)+4}" text-anchor="end">${fmt(lo+pad)}</text>
 <text x="${L}" y="${H-6}">${esc(per[0])}</text>${wsi>0?`<text x="${X(wsi)}" y="${H-6}" text-anchor="middle">${esc(per[wsi])}</text>`:''}<text x="${W-R}" y="${H-6}" text-anchor="end">${esc(per[n-1])}</text>
 <path d="${pl(a)}" fill="none" stroke="var(--a)" stroke-width="2"/><path d="${pl(b)}" fill="none" stroke="var(--b)" stroke-width="2" stroke-dasharray="5,3"/></svg>
 <div class="small"><span style="color:var(--a)">━ A: ${esc(M.a)}</span> &nbsp; <span style="color:var(--b)">╍ B: ${esc(M.b)}</span></div>`}
function barChart(vals,per,label){
 const W=640,H=120,L=56,R=8,T=8,B=20,n=per.length;const vs=vals.filter(v=>v!=null);if(!vs.length)return'';
 let lo=Math.min(0,...vs),hi=Math.max(0,...vs);if(lo===hi)hi=1;const X=i=>L+(W-L-R)*i/n,bw=Math.max(1,(W-L-R)/n-1),Y=v=>T+(H-T-B)*(1-(v-lo)/(hi-lo));
 return `<svg viewBox="0 0 ${W} ${H}" width="100%"><line x1="${L}" x2="${W-R}" y1="${Y(0)}" y2="${Y(0)}" stroke="var(--line)"/>
 ${vals.map((v,i)=>v==null?'':`<rect x="${X(i)}" width="${bw}" y="${Math.min(Y(v),Y(0))}" height="${Math.max(.5,Math.abs(Y(v)-Y(0)))}" fill="${v>=0?'var(--pos)':'var(--neg)'}" opacity=".8"><title>${esc(per[i])}: ${fmt(v)}</title></rect>`).join('')}
 <text x="${L-6}" y="${Y(hi)+8}" text-anchor="end">${fmt(hi)}</text><text x="${L-6}" y="${Y(lo)}" text-anchor="end">${fmt(lo)}</text>
 <text x="${L}" y="${H-4}">${esc(per[0])}</text><text x="${W-R}" y="${H-4}" text-anchor="end">${esc(per[n-1])}</text></svg><div class="small mut">${label}</div>`}

// ---------- detail panel
const hist=[];
function show(n,noPush){
 n=String(n).toUpperCase();
 if(SERVE&&!(n in D.vals)){fetch('/api/detail?n='+encodeURIComponent(n)).then(r=>r.json()).then(o=>{D.vals[n]=o.error?null:o.vals;if(o.dec)D.dec[n]=o.dec;if(o.eqt)D.eqt[n]=o.eqt;if(o.why)D.why[n]=o.why;if(o.desc&&!ROW[n])D.desc[n]=o.desc;show(n,noPush)});return}
 if(!noPush){hist.push(n)}
 const r=ROW[n],v=D.vals[n],d=D.dec[n],eq=D.eqt[n],kind=r?r[9]:(M.has_eqs?(eq?'':'X'):'-');
 let h=`<div class="dh"><div>${hist.length>1?'<button class="btn" id="bk">← Back</button> ':''}</div><button class="btn" id="cx">Close ✕</button></div>
 <h2 style="margin-top:8px">${esc(n)} ${badge(r?r[9]:'')}</h2><div class="mut">${esc(desc(n))}</div>`;
 if(r)h+=`<div class="kpis"><div class="kpi"><span class="mut small">gap at end (B−A)</span><b class="${cls(r[6])}">${fmt(r[6])}</b><span class="small ${cls(r[7])}">${pct(r[7])}</span></div><div class="kpi"><span class="mut small">largest gap</span><b class="${cls(r[4])}">${fmt(r[4])}</b><span class="small ${cls(r[5])}">${pct(r[5])}</span></div><div class="kpi"><span class="mut small">first differs</span><b>${esc(r[8])}</b></div><div class="kpi"><span class="mut small">score</span><b>${fmt(r[3])}</b></div></div>`;
 else h+=`<div class="card">No difference between A and B in the window${(n in D.vals)?'':' (or series not in the data)'}.</div>`;
 if(v){h+=lineChart(v[0],v[1],D.per,D.wsi);const df=v[0].map((x,i)=>x==null||v[1][i]==null?null:v[1][i]-x);h+=barChart(df,D.per,'Gap B − A by period')}
 else if(r)h+=`<div class="hint">Time series not embedded for this variable (raise --embed to include more).</div>`;
 if(eq){h+=`<h2>Equation</h2><pre class="eq">${esc(eq)}</pre>`;if(D.why[n])h+=`<div class="hint">${esc(D.why[n])}</div>`;if(D.afmap[n])h+=`<div class="hint">Add-factor series: ${vlink(D.afmap[n])} ${gapTxt(D.afmap[n])}</div>`}
 else if(M.has_eqs)h+=`<div class="card">No equation for ${esc(n)} — it is an <b>assumption</b> (exogenous). Its path is set directly by the scenario.</div>`;
 if(d&&d.err)h+=`<div class="card">Could not decompose: ${esc(d.err)}</div>`;
 else if(d)h+=decHTML(n,d);
 if(SERVE&&eq&&r)h+=`<p><button class="btn" id="atb2">Split ${esc(n)}'s gap across the scenario shocks</button></p><div id="atout"></div>`;
 else if(r&&eq&&!D.dec[n]){const l=D.esh[NI[n]];if(l&&l.length){h+=`<h2>Why it changed</h2><div class="hint">Share of the change in its left-hand side coming from each driver:</div>`+l.slice().sort((a,b)=>Math.abs(b[1])-Math.abs(a[1])).map(e=>`${vlink(D.nodes[e[0]])} ${Math.round(e[1]*100)}%`).join(' · ')+(D.afsh[NI[n]]!=null?` · <i>add-factor</i> ${D.afmap[n]?vlink(D.afmap[n]):''} <b>${Math.round(D.afsh[NI[n]]*100)}%</b>`:'')}}
 const ps=(D.par[NI[n]]||[]).map(i=>D.nodes[i]),cs=(CH[NI[n]]||[]).map(i=>D.nodes[i]);
 if(ps.length)h+=`<h2>Drivers in its equation</h2><div>${ps.map(p=>`${vlink(p)} ${gapTxt(p)}`).join(' &nbsp;·&nbsp; ')}</div>`;
 if(!cs.length&&M.has_eqs)h+=`<h2>Feeds into</h2><div class="hint">Nothing — no other equation uses ${esc(n)}. It is an end point of the model.</div>`;
 if(cs.length){const cd=cs.filter(c=>ROW[c]),cn=cs.length-cd.length;h+=`<h2>Feeds into</h2><div class="hint">${cs.length} variable(s) use ${esc(n)} in their equations${cn?`; ${cn} of them do not differ`:''}.</div><div>${cd.sort((a,b)=>ROW[b][3]-ROW[a][3]).slice(0,60).map(c=>{const sh=esh(c,n);return `${vlink(c)} ${gapTxt(c)}${sh!=null?` <span class="small mut">(${Math.round(sh*100)}% from ${esc(n)})</span>`:''}`}).join(' &nbsp;·&nbsp; ')}</div>`}
 const det=$('#det');det.innerHTML=h;det.classList.add('open');det.scrollTop=0;
 $('#cx').onclick=()=>{det.classList.remove('open');hist.length=0};
 const bk=$('#bk');if(bk)bk.onclick=()=>{hist.pop();show(hist[hist.length-1],true)};
 const pb=$('#pb');if(pb)pb.onclick=()=>{$('#ptab').classList.toggle('hide')};
 const dp=$('#dper');if(dp)dp.onchange=()=>{$('#dbars').innerHTML=decBars(n,D.dec[n],+dp.value)};
 const at=$('#atb2');if(at)at.onclick=()=>{$('#atout').innerHTML='<div class="hint">Re-solving…</div>';fetch('/api/attr?y='+encodeURIComponent(n)).then(r=>r.json()).then(o=>{if(o.error){$('#atout').innerHTML=`<div class="card">${esc(o.error)}</div>`;return}
  const i=o.iend,g=o.gap[i],mx=Math.max(1e-30,...o.rows.map(r=>Math.abs(r[3][i]||0)));let s=0;o.rows.forEach(r=>s+=r[3][i]||0);
  $('#atout').innerHTML=`<div class="hint">Gap at ${esc(o.per[i])}: <b>${fmt(g)}</b>. Each line re-solves the model with only that shock switched from A to B.</div><table class="bars">${o.rows.map(r=>`<tr><td style="width:42%">${r[1]==='addfactor'?'add-factor of '+vlink(r[0])+(r[2]?' ('+esc(r[2])+')':''):vlink(r[0])}</td><td style="width:34%">${bar(r[3][i],mx)}</td><td class="n">${fmt(r[3][i])}</td><td class="n mut">${g?Math.round(100*r[3][i]/g)+'%':''}</td></tr>`).join('')}<tr><td><i>Interaction between shocks</i></td><td></td><td class="n">${fmt(g-s)}</td><td class="n mut">${g?Math.round(100*(g-s)/g)+'%':''}</td></tr></table>`})};
}
document.addEventListener('keydown',e=>{if(e.key==='Escape'){$('#det').classList.remove('open');hist.length=0}});
function bar(v,mx){v=v||0;const w=Math.min(100,100*Math.abs(v)/mx);return `<div style="display:flex"><div style="width:50%;display:flex;justify-content:flex-end">${v<0?`<div class="bar" style="width:${w}%;background:var(--neg)"></div>`:''}</div><div style="width:50%">${v>0?`<div class="bar" style="width:${w}%;background:var(--pos)"></div>`:''}</div></div>`}
function decBars(n,d,i){
 const g=d.gap[i];const rows=d.cols.map(c=>[c[0],c[1][i]]);
 if(d.af)rows.push(['__AF__',d.af[i]]);
 const mx=Math.max(1e-30,Math.abs(g||0),...rows.map(r=>Math.abs(r[1]||0)));
 if(d.int&&Math.abs(d.int[i]||0)>1e-3*mx)rows.push(['__INT__',d.int[i]]);
 const sh=v=>g&&Math.abs(g)>1e-12*mx?Math.round(100*(v||0)/g)+'%':'';
 let h=`<table class="bars"><tr><td style="width:40%"><b>Gap B−A</b></td><td style="width:34%">${bar(g,mx)}</td><td class="n"><b>${fmt(g)}</b></td><td class="n mut">100%</td></tr>`;
 h+=rows.map(r=>`<tr><td>${r[0]==='__AF__'?`<i>Add-factor</i> ${d.afn?vlink(d.afn):'<span class="small mut">(residual)</span>'}`:r[0]==='__INT__'?'<i>Interaction (non-linear)</i>':vlink(r[0])+' '+badge(ROW[r[0]]?ROW[r[0]][9]:'')}</td><td>${bar(r[1],mx)}</td><td class="n">${fmt(r[1])}</td><td class="n mut">${sh(r[1])}</td></tr>`).join('');
 if(!d.cols.length&&!d.af)h+=`<tr><td colspan="4" class="mut">None of its drivers differ and its add-factor did not change${g?'':' — that is why it did not move'}.</td></tr>`;
 return h+'</table>'}
function decHTML(n,d){
 const i=d.iend;const opts=d.gap.map((g,k)=>`<option value="${k}" ${k===i?'selected':''}>${esc(D.per[k+D.wsi])}</option>`).join('');
 const fs=d.cols.reduce((s,c)=>s+(c[1][i]||0),0),g=d.gap[i];
 let h=`<h2>Where the gap comes from</h2><div class="hint">Each line is what the gap of ${esc(n)} would be if <i>only</i> that input moved from A to B, with the equation re-solved through time (effects that build up through its own lags are credited to the input that caused them). The formula is the same in both scenarios, so the rest comes from the <b>add-factor</b>.${g?` At ${esc(D.per[i+D.wsi])}: <b>${Math.round(100*fs/g)}%</b> from the formula${d.af?`, <b>${Math.round(100*d.af[i]/g)}%</b> from the add-factor`:''}.`:''}</div>
 <div class="ctl"><label class="small mut">Period <select id="dper">${opts}</select></label></div><div id="dbars">${decBars(n,d,i)}</div>`;
 if(d.cols.length||d.af){const ks=d.cols.slice(0,8);h+=`<p><button class="btn" id="pb">Show every period</button></p><div id="ptab" class="hide tw"><table><thead><tr><th>Period</th><th class="n">Gap</th>${ks.map(k=>`<th class="n">${esc(k[0])}</th>`).join('')}${d.af?'<th class="n">Add-factor</th>':''}</tr></thead><tbody>${d.gap.map((v,k)=>`<tr><td>${esc(D.per[k+D.wsi])}</td><td class="n">${fmt(v)}</td>${ks.map(c=>`<td class="n ${cls(c[1][k])}">${fmt(c[1][k])}</td>`).join('')}${d.af?`<td class="n ${cls(d.af[k])}">${fmt(d.af[k])}</td>`:''}</tr>`).join('')}</tbody></table></div>`}
 return h}
tab('sum');
</script></body></html>
'''


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def build_model(args):
    t0 = time.time()
    say = (lambda m: print(m, file=sys.stderr)) if not getattr(args, 'quiet', False) else (lambda m: None)
    say('Loading scenario A ...')
    A = load_scenario(args.a, args.name_a or _label(args.a), args.sheet)
    say('Loading scenario B ...')
    B = load_scenario(args.b, args.name_b or _label(args.b), args.sheet)
    eqs, addl, eqstats, fails = {}, {}, None, []
    if args.eq:
        say('Reading equations ...')
        eqs, addl, eqstats, fails = load_equations(args.eq)
    meta = load_meta(args.meta) if args.meta else None
    M = Model(A, B, eqs, addl, start=args.start, end=args.end, tol=args.tol, geo_sep=args.geo_sep,
              adj_threshold=args.adj_threshold, meta=meta, verbose=not getattr(args, 'quiet', False),
              af_suffix=args.af_suffix)
    say('Loaded %d + %d series, %d periods, %d equations in %.1fs'
        % (len(A.series), len(B.series), M.T, len(eqs), time.time() - t0))
    return M, eqstats, fails


def _label(paths):
    return os.path.splitext(os.path.basename(paths[0]))[0]


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest='cmd')

    def common(sp):
        sp.add_argument('--a', action='append', required=True, metavar='FILE',
                        help='scenario A data (csv/tsv/xlsx). Repeat to combine several files.')
        sp.add_argument('--b', action='append', required=True, metavar='FILE', help='scenario B data')
        sp.add_argument('--eq', action='append', metavar='PATH',
                        help='equation text file(s) or a folder of them (EViews syntax). Repeatable.')
        sp.add_argument('--name-a', help='label for A (default: file name)')
        sp.add_argument('--name-b', help='label for B')
        sp.add_argument('--sheet', help='xlsx sheet name or number (default: first)')
        sp.add_argument('--meta', help='optional CSV with series name + description columns')
        sp.add_argument('--start', help='first period to compare (default: where scenarios first diverge)')
        sp.add_argument('--end', help='last period to compare (default: end of data)')
        sp.add_argument('--tol', type=float, default=1e-6, help='relative tolerance for "differs" (default 1e-6)')
        sp.add_argument('--geo-sep', default='.', help='separator before geography code in names (default ".")')
        sp.add_argument('--adj-threshold', type=float, default=0.25,
                        help='flag a variable as "mostly add-factor" above this share of its change (0.25)')
        sp.add_argument('--af-suffix', default='_A',
                        help='suffix of add-factor series, e.g. GDP.US_A (default _A; comma-separate several)')
        sp.add_argument('--quiet', action='store_true')

    def report_opts(sp):
        sp.add_argument('--embed', type=int, default=1500,
                        help='how many top series get charts + full detail in the report (default 1500)')
        sp.add_argument('--max-decomp', type=int, default=50000,
                        help='max equations analysed for driver contributions (default 50000)')
        sp.add_argument('--context', type=int, default=8, help='periods of history shown before the window (8)')
        sp.add_argument('--focus', action='append', help='variable to always include in full detail (repeatable)')

    sp = sub.add_parser('report', help='build the interactive HTML report (a single file)')
    common(sp)
    report_opts(sp)
    sp.add_argument('-o', '--out', default='scenario_compare.html')

    sp = sub.add_parser('serve', help='same report, served locally, with exact impact X->Y and shock attribution')
    common(sp)
    report_opts(sp)
    sp.add_argument('--port', type=int, default=8765)
    sp.add_argument('--no-browser', action='store_true')

    sp = sub.add_parser('top', help='print the biggest changes')
    common(sp)
    sp.add_argument('-n', type=int, default=30)
    sp.add_argument('--geo', help='only this geography code, e.g. IUSA')
    sp.add_argument('--match', help='regex on name or description')
    sp.add_argument('--type', help='X=assumption, F=add-factor series, A=mostly add-factor, E=formula, N=not evaluable')
    sp.add_argument('--csv', help='write stats for ALL differing series to this CSV')

    sp = sub.add_parser('explain', help='contribution of each input to the gap of a variable')
    sp.add_argument('var')
    common(sp)
    sp.add_argument('--at', help='period to explain (default: end of window)')
    sp.add_argument('--depth', type=int, default=1, help='levels of inputs to follow (default 1)')
    sp.add_argument('--min-share', type=float, default=0.1, help='only follow inputs above this share (0.1)')
    sp.add_argument('--by-period', action='store_true', help='also print contributions period by period')

    sp = sub.add_parser('addfactors', help='formula vs add-factor share of the gap, for every equation')
    common(sp)
    sp.add_argument('-n', type=int, default=40)
    sp.add_argument('--all', action='store_true', help='include equations whose add-factor did not change')

    sp = sub.add_parser('impact', help='can X affect Y, and by how much? (re-solves the equations)')
    sp.add_argument('x')
    sp.add_argument('y', nargs='?')
    common(sp)
    sp.add_argument('--shock', help='hypothetical change in X instead of its A->B change: "1%%" or "0.25" (units)')
    sp.add_argument('--by-period', action='store_true')
    sp.add_argument('-n', type=int, default=30, help='rows when Y is omitted')

    sp = sub.add_parser('attribute', help='split the gap of Y across all scenario shocks upstream of it')
    sp.add_argument('y')
    common(sp)
    sp.add_argument('-n', type=int, default=25)

    sp = sub.add_parser('path', help='the channels through which X reaches Y')
    sp.add_argument('x')
    sp.add_argument('y')
    common(sp)
    sp.add_argument('-k', type=int, default=5, help='max number of paths')

    sp = sub.add_parser('info', help='show what was loaded and parsed')
    common(sp)

    args = p.parse_args(argv)
    if not args.cmd:
        p.print_help()
        return 1
    M, eqstats, fails = build_model(args)
    if args.cmd == 'info':
        cmd_info(M, args, eqstats, fails)
    elif args.cmd == 'top':
        cmd_top(M, args)
    elif args.cmd == 'explain':
        cmd_explain(M, args)
    elif args.cmd == 'path':
        cmd_path(M, args)
    elif args.cmd == 'report':
        build_report(M, args, eqstats)
    elif args.cmd == 'serve':
        cmd_serve(M, args, eqstats)
    elif args.cmd == 'addfactors':
        cmd_addfactors(M, args)
    elif args.cmd == 'impact':
        cmd_impact(M, args)
    elif args.cmd == 'attribute':
        cmd_attribute(M, args)
    return 0


if __name__ == '__main__':
    sys.exit(main())
