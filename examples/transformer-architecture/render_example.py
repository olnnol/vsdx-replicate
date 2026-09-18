# -*- coding: utf-8 -*-
"""Example scene data: rebuild the Transformer architecture diagram
(1426x1500, Wikimedia Commons, CC BY 4.0, by dvgodoy) as a native Visio
.vsdx via COM automation.
Pass 1 writes out/replica.vsdx + out/text_runs.json; pass 2 (patch_text.py)
applies the rich text.  1 px = 0.1 mm.

Source: https://commons.wikimedia.org/wiki/File:Transformer,_full_architecture.png
        dvgodoy, "dl-visuals" - CC BY 4.0
"""
import os, json, math
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(BASE, 'out'), exist_ok=True)
OUTV = os.path.join(BASE, 'out', 'replica.vsdx')
W, H = 1426, 1500
MM = 0.1
IN = 254.0

def ix(px): return px / IN
def iy(py): return (H - py) / IN
def hx(h): return f'RGB({int(h[0:2],16)},{int(h[2:4],16)},{int(h[4:6],16)})'

C = dict(
    black='000000', white='FFFFFF',
    gfill='D5E8D4', gline='82B366',
    yfill='FFF2CC', yline='D6B656',
    bfill='DAE8FD', bline='6C8EBF',
    pfill='F8CECC', pline='B85450',
    oline='D79B00',
)

app = win32com.client.Dispatch('Visio.InvisibleApp')
app.AlertResponse = 2
doc = app.Documents.Add('')
page = doc.Pages(1)
page.PageSheet.Cells('PageWidth').FormulaU = f'{W*MM:g} mm'
page.PageSheet.Cells('PageHeight').FormulaU = f'{H*MM:g} mm'

_runs = {}
_tno = [0]

def rect(x, y, w, h, fill, line, lw=2.0, r=10, name='', dashed=False):
    s = page.DrawRectangle(ix(x), iy(y + h), ix(x + w), iy(y))
    if fill: s.Cells('FillForegnd').FormulaU = hx(fill)
    else: s.Cells('FillPattern').FormulaU = '0'
    if line:
        s.Cells('LineColor').FormulaU = hx(line)
        s.Cells('LineWeight').FormulaU = f'{lw*MM:g} mm'
        if dashed: s.Cells('LinePattern').FormulaU = '2'
    else:
        s.Cells('LinePattern').FormulaU = '0'
    if r: s.Cells('Rounding').FormulaU = f'{r*MM:g} mm'
    if name: s.Name = name
    return s

def invis(x, y, w, h):
    s = page.DrawRectangle(ix(x), iy(y + h), ix(x + w), iy(y))
    s.Cells('LinePattern').FormulaU = '0'
    s.Cells('FillPattern').FormulaU = '0'
    return s

def seg(x1, y1, x2, y2, color=C['black'], lw=2.5, dashed=False):
    s = page.DrawLine(ix(x1), iy(y1), ix(x2), iy(y2))
    s.Cells('LineColor').FormulaU = hx(color)
    s.Cells('LineWeight').FormulaU = f'{lw*MM:g} mm'
    s.Cells('LineCap').FormulaU = '0'
    if dashed: s.Cells('LinePattern').FormulaU = '2'
    return s

def poly(pts, fill=None, line=None, lw=2.0):
    p = list(pts) + ([pts[0]] if fill else [])   # close filled polygons
    arr = []
    for x, y in p: arr += [x / IN, (H - y) / IN]
    s = page.DrawPolyline(arr, 0)
    if fill: s.Cells('FillForegnd').FormulaU = hx(fill)
    else: s.Cells('FillPattern').FormulaU = '0'
    if line:
        s.Cells('LineColor').FormulaU = hx(line)
        s.Cells('LineWeight').FormulaU = f'{lw*MM:g} mm'
    else:
        s.Cells('LinePattern').FormulaU = '0'
    return s

def put_text(shape, lines, align='c', tw=None, th=None, dy=0, angle=0):
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
    if angle: shape.Cells('TxtAngle').FormulaU = f'{angle:g} deg'
    return shape

LS = 0        # letter-spacing disabled until unit is calibrated

def B(t, sz=20, **o):
    return (t, dict(sz=sz, f='Verdana', b=1, col='000000', ls=LS) | o)

def ftext(x, y, w, h, lines, align='c'):
    return put_text(invis(x, y, w, h), lines, align=align, tw=w, th=h)

def box(x, y, w, h, fill, line, label, sz=20.5, name=''):
    s = rect(x-1.5, y-1.5, w+3, h+3, fill, line, lw=3.0, r=10, name=name)
    put_text(s, [[B(t, sz)] for t in label.split('\n')], tw=w, th=h)
    return s

def circle(cx, cy, r, plus=True, lw=3.0, nm=''):
    s = page.DrawOval(ix(cx - r), iy(cy - r), ix(cx + r), iy(cy + r))
    s.Cells('LineColor').FormulaU = hx(C['black'])
    s.Cells('LineWeight').FormulaU = f'{lw*MM:g} mm'
    s.Cells('FillPattern').FormulaU = '0'
    if nm: s.Name = nm
    if plus:
        seg(cx - r*0.31, cy, cx + r*0.31, cy, lw=lw)
        seg(cx, cy - r*0.31, cx, cy + r*0.31, lw=lw)
    return s

def arrowhead(x, y, dx, dy, size=13, halfw=6.5):
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    px_, py_ = -uy, ux
    b1 = (x - size*ux + halfw*px_, y - size*uy + halfw*py_)
    b2 = (x - size*ux - halfw*px_, y - size*uy - halfw*py_)
    return poly([(x, y), b1, b2], fill=C['black'])

def arrow(x1, y1, x2, y2, lw=2.5):
    dx, dy = x2 - x1, y2 - y1
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    seg(x1, y1, x2 - ux*10, y2 - uy*10, lw=lw)
    arrowhead(x2, y2, dx, dy)

def curve(pts, lw=2.5, r=14):
    """polyline with rounded corners (quadratic-bezier corner arcs)"""
    if len(pts) == 2:
        seg(pts[0][0], pts[0][1], pts[1][0], pts[1][1], lw=lw)
        return
    out = [pts[0]]
    for i in range(1, len(pts) - 1):
        x0, y0 = pts[i-1]; x1, y1 = pts[i]; x2, y2 = pts[i + 1]
        d1 = math.hypot(x1 - x0, y1 - y0); d2 = math.hypot(x2 - x1, y2 - y1)
        rr = min(r, d1/2, d2/2)
        a = (x1 - (x1-x0)/d1*rr, y1 - (y1-y0)/d1*rr)
        b = (x1 + (x2-x1)/d2*rr, y1 + (y2-y1)/d2*rr)
        out.append(a)
        for k in range(1, 8):
            t = k/8
            out.append(((1-t)**2*a[0] + 2*(1-t)*t*x1 + t**2*b[0],
                        (1-t)**2*a[1] + 2*(1-t)*t*y1 + t**2*b[1]))
        out.append(b)
    out.append(pts[-1])
    arr = []
    for x, y in out: arr += [x / IN, (H - y) / IN]
    s = page.DrawPolyline(arr, 0)
    s.Cells('LineColor').FormulaU = hx(C['black'])
    s.Cells('LineWeight').FormulaU = f'{lw*MM:g} mm'
    s.Cells('LineCap').FormulaU = '0'
    s.Cells('FillPattern').FormulaU = '0'
    return s

def path_arrow(pts, lw=2.5, r=14):
    """rounded-corner path ending with an arrowhead at the last point"""
    x0, y0 = pts[-2]; x1, y1 = pts[-1]
    dx, dy = x1-x0, y1-y0
    L = math.hypot(dx, dy); ux, uy = dx/L, dy/L
    pts2 = list(pts[:-1]) + [(x1 - ux*10, y1 - uy*10)]
    curve(pts2, lw=lw, r=r)
    arrowhead(x1, y1, dx, dy)

ENCX, DECX = 509.5, 890.5

# ================================================================ boxes
box(407, 521, 203, 41, C['gfill'], C['gline'], 'Norm', name='E_NORM1')
box(407, 689, 203, 79, C['yfill'], C['yline'], 'Feed-Forward\nNetwork', name='E_FFN')
box(407, 799, 203, 45, C['gfill'], C['gline'], 'Norm', name='E_NORM2')
box(407, 968, 203, 79, C['bfill'], C['bline'], 'Multi-Headed\nSelf-Attention', name='E_ATTN')
box(407, 1119, 203, 41, C['gfill'], C['gline'], 'Norm', name='E_NORM3')
box(407, 1329, 203, 69, C['white'], C['black'], 'Embeddings/\nProjections', name='E_EMBED')
box(789, 81, 203, 37, C['white'], C['black'], 'Linear', name='D_LINEAR')
box(789, 149, 203, 45, C['gfill'], C['gline'], 'Norm', name='D_NORM1')
box(789, 318, 203, 79, C['yfill'], C['yline'], 'Feed-Forward\nNetwork', name='D_FFN')
box(789, 431, 203, 38, C['gfill'], C['gline'], 'Norm', name='D_NORM2')
box(789, 596, 203, 80, C['bfill'], C['bline'], 'Multi-Headed\nCross-Attention', name='D_ATTN')
box(789, 768, 203, 45, C['gfill'], C['gline'], 'Norm', name='D_NORM3')
box(789, 937, 203, 100, C['pfill'], C['pline'], 'Masked\nMulti-Headed\nSelf-Attention', name='D_MASK')
box(789, 1117, 203, 42, C['gfill'], C['gline'], 'Norm', name='D_NORM4')
box(789, 1329, 203, 69, C['white'], C['black'], 'Embeddings/\nProjections', name='D_EMBED')

# ================================================================ containers + dashed rects
rect(374.5, 594.5, 289, 619, None, C['oline'], lw=3.0, r=40, name='ENC_BOX')
rect(751.5, 223.5, 299, 990, None, C['gline'], lw=3.0, r=40, name='DEC_BOX')
for (x, y, w, h, nm) in ((384.5, 607.5, 268, 274, 'E_DASH1'), (384.5, 884.5, 268, 317, 'E_DASH2'),
                         (766.5, 234.5, 268, 276, 'D_DASH1'), (766.5, 513.5, 268, 337, 'D_DASH2'),
                         (766.5, 853.5, 268, 348, 'D_DASH3')):
    rect(x, y, w, h, None, C['black'], lw=3.0, r=40, name=nm, dashed=True)

# ================================================================ circles
for cy in (636, 914, 1264):
    circle(ENCX, cy, 21, lw=3.0, nm=f'E_ADD{cy}')
for cy in (264, 543, 883, 1264):
    circle(DECX, cy, 21, lw=3.0, nm=f'D_ADD{cy}')

def pos_icon(cx, cy, r=21):
    circle(cx, cy, r, plus=False, lw=3.0)
    pts = []
    for k in range(33):
        t = k/32
        pts.append((cx - r*0.95 + t*r*1.90, cy - math.sin(t*2*math.pi)*r*0.45))
    arr = []
    for x, y in pts: arr += [x / IN, (H - y) / IN]
    s = page.DrawPolyline(arr, 0)
    s.Cells('LineColor').FormulaU = hx(C['black'])
    s.Cells('LineWeight').FormulaU = f'{3.0*MM:g} mm'
    s.Cells('FillPattern').FormulaU = '0'
    return s

pos_icon(415, 1264)
pos_icon(981, 1264)

def vkq(ax, bar_y, box_bottom, left, right, q_from=None):
    """V K Q distributor: bar at bar_y, three arrows up into box_bottom"""
    if q_from is None:
        curve([(ax, bar_y + 19), (ax, bar_y)], lw=2.5) if False else None
        path_arrow([(ax, bar_y), (left, bar_y), (left, box_bottom)], lw=2.5)
        path_arrow([(ax, bar_y), (right, bar_y), (right, box_bottom)], lw=2.5)
        arrow(ax, bar_y, ax, box_bottom)
        # labels
        ftext(left-30, box_bottom+16, 30, 22, [[B('V', 19)]])
        ftext(ax-15, box_bottom+16, 30, 22, [[B('K', 19)]])
        ftext(right-8, box_bottom+16, 30, 22, [[B('Q', 19)]])
    else:
        path_arrow([(ax, bar_y), (left, bar_y), (left, box_bottom)], lw=2.5)
        arrow(ax, bar_y, ax, box_bottom)
        path_arrow(q_from + [(right, box_bottom)], lw=2.5)
        ftext(left-30, box_bottom+16, 30, 22, [[B('V', 19)]])
        ftext(ax-15, box_bottom+16, 30, 22, [[B('K', 19)]])
        ftext(right-8, box_bottom+16, 30, 22, [[B('Q', 19)]])

# ================================================================ encoder flow
arrow(ENCX, 1329, ENCX, 1287)                    # embeddings -> add
arrow(437, 1264, 487.5, 1264)                    # pos icon -> add
arrow(ENCX, 1245, ENCX, 1162)                    # add -> norm3
seg(ENCX, 1119, ENCX, 1100)                      # norm3 -> distributor bar
vkq(ENCX, 1100, 1049, ENCX-52.5, ENCX+50.5)      # V K Q into self-attention
arrow(ENCX, 968, ENCX, 936)                      # attn -> add
arrow(ENCX, 894, ENCX, 846)                      # add -> norm2
arrow(ENCX, 799, ENCX, 770)                      # norm2 -> ffn
arrow(ENCX, 689, ENCX, 658)                      # ffn -> add
arrow(ENCX, 616, ENCX, 564)                      # add -> norm1
# encoder output -> arrow into decoder dashed rect, then feed line to V/K
path_arrow([(ENCX, 521), (ENCX, 491), (704, 491), (704, 739), (747, 739)], lw=2.5)
seg(747, 739, 886, 739, lw=2.5)
# residuals (encoder)
path_arrow([(ENCX, 872), (632, 872), (632, 636), (531.5, 636)], lw=2.5)
path_arrow([(ENCX, 1192), (632, 1192), (632, 914), (531.5, 914)], lw=2.5)

# ================================================================ decoder flow
ftext(806, 14, 170, 30, [[B('Predictions', 21)]])          # top label
arrow(DECX, 81, DECX, 46)                                  # linear -> predictions
arrow(DECX, 149, DECX, 120)                                # norm1 -> linear
arrow(DECX, 244, DECX, 196)                                # add -> norm1
arrow(DECX, 318, DECX, 286)                                # ffn -> add
arrow(DECX, 431, DECX, 399)                                # norm2 -> ffn
arrow(DECX, 523, DECX, 471)                                # add -> norm2   (add543 top at 523)
arrow(DECX, 596, DECX, 565)                                # cross-attn -> add543
# cross-attention V K Q (V/K from encoder output line at y=739, Q from norm3 top)
path_arrow([(838, 739), (838, 678)], lw=2.5)
path_arrow([(886, 739), (886, 678)], lw=2.5)
path_arrow([(DECX, 768), (DECX, 752), (942, 752), (942, 678)], lw=2.5, r=12)
ftext(808, 690, 30, 22, [[B('V', 19)]])
ftext(874, 690, 30, 22, [[B('K', 19)]])
ftext(927, 690, 30, 22, [[B('Q', 19)]])
arrow(DECX, 863, DECX, 815)                                # add883 -> norm3
arrow(DECX, 937, DECX, 905)                                # mask-attn -> add883
seg(DECX, 1117, DECX, 1095)                                # norm4 -> bar
vkq(DECX, 1095, 1039, DECX-49.5, DECX+49.5)                # V K Q into masked attn
arrow(DECX, 1245, DECX, 1161)                              # add1265 -> norm4
arrow(DECX, 1329, DECX, 1287)                              # embeddings -> add
arrow(964, 1264, 913.5, 1264)                              # pos icon -> add
# residuals (decoder)
path_arrow([(DECX, 500), (1014, 500), (1014, 264), (912.5, 264)], lw=2.5)
path_arrow([(DECX, 842), (1014, 842), (1014, 543), (912.5, 543)], lw=2.5)
path_arrow([(DECX, 1192), (1014, 1192), (1014, 883), (912.5, 883)], lw=2.5)

# ================================================================ bottom labels
ftext(360, 1436, 300, 32, [[B('Source Sequence', 21)]])
ftext(740, 1436, 300, 60, [[B('Shifted', 21)], [B('Target Sequence', 21)]])
arrow(ENCX, 1426, ENCX, 1400)
arrow(DECX, 1426, DECX, 1400)
ftext(240, 1242, 170, 54, [[B('Positional', 21)], [B('Encoding', 21)]])
ftext(1000, 1242, 170, 54, [[B('Positional', 21)], [B('Encoding', 21)]])
ftext(230, 860, 160, 54, [[B('Nx', 19)], [B(chr(34)+'Layers'+chr(34), 19)]])
ftext(1030, 860, 160, 54, [[B('Nx', 19)], [B(chr(34)+'Layers'+chr(34), 19)]])

# ================================================================ save
doc.SaveAs(OUTV)
print('saved', OUTV, 'texts:', _tno[0])
json.dump(_runs, open(os.path.join(BASE, 'out', 'text_runs.json'), 'w'),
          ensure_ascii=False, indent=1)
doc.Close()
app.Quit()
print('done pass1')