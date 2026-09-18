# -*- coding: utf-8 -*-
"""Example scene data: rebuild the TARPON pipeline flowchart (2531x1316,
PLOS Computational Biology, CC BY 4.0) as a native Visio .vsdx via COM.
Pass 1 writes out/replica.vsdx + out/text_runs.json; pass 2 (patch_text.py)
applies the rich text.  1 px = 0.1 mm.

Source figure: Deimler N, Ho DV, Paul N, Gill Z, Baumann P (2026)
"TARPON - A Telomere Analysis and Research Pipeline Optimized for Nanopore
sequencing", PLOS Computational Biology. Fig 1(a). CC BY 4.0.
"""
import os, json
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(BASE, 'out'), exist_ok=True)
OUTV = os.path.join(BASE, 'out', 'replica.vsdx')
W, H = 2531, 1316
MM = 0.1
IN = 254.0

def ix(px): return px / IN
def iy(py): return (H - py) / IN
def hx(h): return f'RGB({int(h[0:2],16)},{int(h[2:4],16)},{int(h[4:6],16)})'

C = dict(green='084D41', magenta='D81A61', blue='3A82C4', yellow='FEC110',
         dark='231F20', white='FFFFFF')

app = win32com.client.Dispatch('Visio.InvisibleApp')
app.AlertResponse = 2
doc = app.Documents.Add('')
page = doc.Pages(1)
page.PageSheet.Cells('PageWidth').FormulaU = f'{W*MM:g} mm'
page.PageSheet.Cells('PageHeight').FormulaU = f'{H*MM:g} mm'

_runs = {}
_tno = [0]

def rect(x, y, w, h, fill, r=0, name=''):
    s = page.DrawRectangle(ix(x), iy(y + h), ix(x + w), iy(y))
    s.Cells('FillForegnd').FormulaU = hx(fill)
    s.Cells('LinePattern').FormulaU = '0'
    if r: s.Cells('Rounding').FormulaU = f'{r*MM:g} mm'
    if name: s.Name = name
    return s

def poly(pts, fill=None, line=None, lw=3, close=True):
    p = list(pts) + ([pts[0]] if close else [])
    arr = []
    for x, y in p: arr += [x / IN, (H - y) / IN]
    s = page.DrawPolyline(arr, 0)
    if fill:
        s.Cells('FillForegnd').FormulaU = hx(fill)
    else:
        s.Cells('FillPattern').FormulaU = '0'
    if line:
        s.Cells('LineColor').FormulaU = hx(line)
        s.Cells('LineWeight').FormulaU = f'{lw*MM:g} mm'
    else:
        s.Cells('LinePattern').FormulaU = '0'
    return s

def invis(x, y, w, h):
    s = page.DrawRectangle(ix(x), iy(y + h), ix(x + w), iy(y))
    s.Cells('LinePattern').FormulaU = '0'
    s.Cells('FillPattern').FormulaU = '0'
    return s

def put_text(shape, lines, align='c', tw=None, th=None, dy=0):
    shape.Text = '\n'.join(''.join(t for t, _ in ln) for ln in lines)
    shape.Cells('Para.HorzAlign').FormulaU = '1' if align == 'c' else '0'
    for c in ('LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin'):
        shape.Cells(c).FormulaU = '0 mm'
    if dy:
        shape.Cells('TopMargin' if dy > 0 else 'BottomMargin').FormulaU = f'{abs(dy)*2*MM:g} mm'
    _tno[0] += 1
    name = f'TEXT{_tno[0]:03d}'
    shape.Name = name
    _runs[name] = lines
    if tw: shape.Cells('TxtWidth').FormulaU = f'{tw*MM:g} mm'
    if th: shape.Cells('TxtHeight').FormulaU = f'{th*MM:g} mm'
    return shape

def S(t, sz=60, **o):
    return (t, dict(sz=sz) | o)

def G(t, sz=60, **o):
    return (t, dict(sz=sz, f='A') | o)

def ftext(x, y, w, h, lines, align='c'):
    return put_text(invis(x, y, w, h), lines, align=align, tw=w, th=h)

# ================================================================ green step boxes
GREEN = [
    (916, 319, 'Putative Telomere', 'Identification'),
    (1796, 319, 'Strand Determination', 'and Complementation'),
    (10, 649, 'Subtelomere-to-Telomere', 'Boundary Identification'),
    (916, 640, 'Subtelomeric', 'Filtering'),
    (1796, 640, 'Probe Identification', 'and Demultiplexing'),
]
DYS = {1: 0, 2: -2, 3: 5, 4: 0, 5: -1}
SZS = {1: 60, 2: 61, 3: 61, 4: 61, 5: 61}
for i, (x, y, l1, l2) in enumerate(GREEN):
    n = i + 1
    sz = SZS[n]
    s = rect(x, y, 724, 204, C['green'], r=66, name=f'STEP{n}')
    put_text(s, [[G(l1, sz, col='FFFFFF')], [G(l2, sz, col='FFFFFF')]],
             tw=724, th=204, dy=DYS[n])

# ================================================================ diamonds
d1 = poly([(439, 258), (764, 421), (439, 584), (114, 421)], fill=C['magenta'])
d1.Name = 'DIAM1'
put_text(d1, [[G('Basecalled Data', 61, col='FFFFFF')]], tw=500, th=326, dy=-2)
d2 = poly([(372, 979), (697, 1142), (372, 1305), (47, 1142)], fill=C['blue'])
d2.Name = 'DIAM2'
put_text(d2, [[G('Analysis and', 61, col='231F20')], [G('Output', 61, col='231F20')]],
         tw=500, th=326, dy=13)

# ================================================================ optional box
y1 = rect(1525, 10, 354, 172, C['yellow'], r=66, name='OPTIONAL')
put_text(y1, [[G('Basecalling', 61, col='231F20')], [G('(Optional)', 61, col='231F20')]],
         tw=354, th=172)

# ================================================================ solid wedges
poly([(791, 366), (884, 421), (791, 477)], fill=C['dark'])
poly([(1671, 365), (1764, 421), (1671, 477)], fill=C['dark'])
poly([(2002, 550), (2314, 550), (2158, 613)], fill=C['dark'])
poly([(885, 695), (792, 751), (885, 807)], fill=C['dark'])
poly([(1765, 686), (1672, 741), (1765, 797)], fill=C['dark'])
poly([(216, 884), (528, 884), (372, 947)], fill=C['dark'])

# ================================================================ arrows
# curved arrow: circle fitted through (1478,95) (1340,148) (1278,280)
import math
UCX, UCY, UR = 1474.4, 291.7, 196.7
a1 = math.atan2(95-UCY, 1478-UCX)
a3 = math.atan2(280-UCY, 1278-UCX)
sweep = a3 - a1
while sweep > 0: sweep -= 2*math.pi
N = 48
pts = []
for k in range(N+1):
    ang = a1 + sweep*k/N
    pts.append((UCX + UR*math.cos(ang), UCY + UR*math.sin(ang)))
arr = []
for x, y in pts: arr += [x/IN, (H-y)/IN]
arc = page.DrawPolyline(arr, 0)
arc.Cells('LineColor').FormulaU = hx(C['dark'])
arc.Cells('LineWeight').FormulaU = f'{6.5*MM:g} mm'
arc.Cells('LineCap').FormulaU = '1'
arc.Name = 'CURVE'
# stealth head, outline fitted from the source bitmap: tip / barbL / notch / barbR
poly([(1278, 281), (1259, 233), (1285, 247), (1306, 238)], fill=C['dark'])

# straight up arrow below the optional box
s = page.DrawLine(ix(1705.5), iy(337), ix(1705.5), iy(210))
s.Cells('LineColor').FormulaU = hx(C['dark'])
s.Cells('LineWeight').FormulaU = f'{5*MM:g} mm'
poly([(1705.5, 198), (1680.5, 243), (1705.5, 232), (1730.5, 243)], fill=C['dark'])

# ================================================================ panel label
ftext(51, 26, 90, 66, [[G('a', 100, b=1, col='231F20')]], align='l')

# ================================================================ save
doc.SaveAs(OUTV)
print('saved', OUTV, 'texts:', _tno[0])
json.dump(_runs, open(os.path.join(BASE, 'out', 'text_runs.json'), 'w'),
          ensure_ascii=False, indent=1)
doc.Close()
app.Quit()
print('done pass1')