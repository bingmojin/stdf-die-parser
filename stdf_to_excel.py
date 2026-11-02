"""Parse per-die ring-osc / IV tests from an STDF file into an Excel sheet, one row per (device, die).
Layout follows 56019A0_Vega2_..._LIB_PM_DATA.xlsx (group header row, column-name row, data rows).

Rules:
  1. DieN_F_ring_<x>              -> decimal value in MHz (raw Hz / 1e6), column TEST_LIB_PM_RINGS_F_ring_<x>
  2. DieN_U32_<x> + DieN_L32_<x>  -> '%04X%04X' (U32 then L32), column keeps <x> only (leading ring_ dropped)
  3. Same as 2 for IV names, e.g. cpp48_nmos_ulvtll_idlin_avg
Usage: python3 stdf_to_excel.py out.xlsx a.std [b.std ...]   (or run stdf_to_excel_gui.py)
"""
import re, sys, collections
from pystdf.IO import Parser
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

RX = re.compile(r'^die(\d+)_(F_ring_.+|[UL]32_.+)$', re.I)

# test-group suffix in TEST_TXT ("<> group") -> (sheet group title, column prefix)
GROUPS = collections.OrderedDict([
    ('RingOSC_Read_allDie', ('RINGS Tests', 'TEST_LIB_PM_RINGS_read_reg_ring_')),
    ('IV_AllDies',          ('DEV_IV Tests', 'TEST_LIB_PM_DEV_IV_read_reg_')),
    ('IV_MetalRC_AllDies',  ('DEV_IV METALRC Tests', 'TEST_LIB_PM_DEV_IV_read_reg_')),
    ('IV_Leak_AllDies',     ('DEV_IV LEAKAGE Tests', 'TEST_LIB_PM_DEV_IV_read_reg_')),
])
F_GROUP = ('RINGS F Tests', 'TEST_LIB_PM_RINGS_F_ring_')
META = ['Lot_Number', 'Part_Number', 'Chip_Name', 'Die_Location_X', 'Die_Location_Y', 'LIB_PM_Number', 'Temperature']


def fields(rec, v):
    return dict(zip([f[0] for f in rec.fieldMap], v))


class Sink:
    def __init__(self):
        self.mir = {}
        self.cur = {}
        self.rows = []                              # (meta dict, {col: value})
        self.cols = collections.OrderedDict()       # col -> group title

    def after_send(self, ds, rec_data):
        rec, v = rec_data
        t = rec.__class__.__name__
        if t == 'Mir':
            self.mir = fields(rec, v)
        elif t == 'Pir':
            self.cur = {}
        elif t == 'Ptr':
            d = fields(rec, v)
            name, _, grp = d['TEST_TXT'].partition('<>')
            m = RX.match(name.strip())
            if m and d['RESULT'] is not None:
                self.cur[(int(m.group(1)), m.group(2), grp.strip())] = d['RESULT']
        elif t == 'Prr':
            self._flush(fields(rec, v))

    def _flush(self, prr):
        dies = collections.OrderedDict()            # die -> {col: value}
        pairs = collections.defaultdict(dict)       # (die, grp, x) -> {U, L}
        for (die, rest, grp), val in self.cur.items():
            out = dies.setdefault(die, {})
            if rest.upper().startswith('F_RING_'):
                self._col(out, F_GROUP, rest[7:], val / 1e6)   # Hz -> MHz
            elif grp in GROUPS:
                pairs[(die, grp, re.sub(r'^ring_', '', rest[4:], flags=re.I))][rest[0].upper()] = val
        for (die, grp, x), lu in pairs.items():
            if 'U' in lu and 'L' in lu:
                self._col(dies[die], GROUPS[grp], x, '%04X%04X' % (int(lu['U']), int(lu['L'])))
        x = prr['X_COORD'] if prr['X_COORD'] != -32768 else None
        y = prr['Y_COORD'] if prr['Y_COORD'] != -32768 else None
        for die, vals in sorted(dies.items()):
            meta = dict(Lot_Number=self.mir.get('LOT_ID'), Part_Number=prr['PART_ID'], Chip_Name=self.mir.get('PART_TYP'),
                        Die_Location_X=x, Die_Location_Y=y, LIB_PM_Number=die, Temperature=self.mir.get('TST_TEMP'))
            self.rows.append((meta, vals))

    def _col(self, out, grp, x, val):
        col = grp[1] + x
        out[col] = val
        self.cols.setdefault(col, grp[0])

    def __getattr__(self, n):  # ignore unused Parser hooks
        if n.startswith(('before_', 'after_')):
            return lambda *a: None
        raise AttributeError(n)


def parse_files(srcs, log=print):
    """Parse one or more STDF files into a single Sink (rows appended in file order)."""
    s = Sink()
    for src in srcs:
        log(f'Parsing {src} ...')
        before = len(s.rows)
        p = Parser(inp=open(src, 'rb'))
        p.addSink(s)
        p.parse()
        log(f'  {len(s.rows) - before} rows')
    return s


def write_xlsx(s, dst):
    order = [g[0] for g in GROUPS.values()] + [F_GROUP[0]]
    cols = sorted(s.cols, key=lambda c: order.index(s.cols[c]))   # stable: keeps first-seen order inside a group
    wb = Workbook(); ws = wb.active; ws.title = 'LIB_PM_DATA'
    ws.append([None] * len(META) + [None] * len(cols))
    ws.append(META + cols)
    ws.cell(1, 1, 'DUT Info'); ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(META))
    i = 0
    while i < len(cols):
        j = i
        while j + 1 < len(cols) and s.cols[cols[j + 1]] == s.cols[cols[i]]:
            j += 1
        c0 = len(META) + i + 1
        ws.cell(1, c0, s.cols[cols[i]]); ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=len(META) + j + 1)
        i = j + 1
    for c in ws[1] + ws[2]:
        c.font = Font(bold=True); c.alignment = Alignment(horizontal='center')
    for meta, vals in s.rows:
        ws.append([meta[k] for k in META] + [vals.get(c) for c in cols])
    ws.freeze_panes = ws.cell(3, len(META) + 1)
    wb.save(dst)
    return len(s.rows), len(cols)


if __name__ == '__main__':
    n, m = write_xlsx(parse_files(sys.argv[2:]), sys.argv[1])
    print(f'{n} rows, {m} test columns -> {sys.argv[1]}')
