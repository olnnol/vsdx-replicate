# -*- coding: utf-8 -*-
"""Rebuild the paper figure as a native Visio .vsdx via COM automation.
Pass 1: geometry + plain text via COM, rich-run specs dumped to text_runs.json.
Pass 2 (patch_text.py) rewrites Character sections in the vsdx XML.
Coordinates: source pixels (2800x1718), 1 px = 0.1 mm.
"""
import os, json
import win32com.client
import pythoncom

BASE = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(BASE, 'out', 'img')
OUTV = os.path.join(BASE, 'out', 'output.vsdx')
W, H = 2800, 1718
MM = 0.1

def fx(px):
    return f'{px * MM:g} mm'
def fy(py):
    return f'{(H - py) * MM:g} mm'

IN = 254.0
def ix(px): return px / IN
def iy(py): return (H - py) / IN

def hx(h):
    return f'RGB({int(h[0:2],16)},{int(h[2:4],16)},{int(h[4:6],16)})'

C = dict(
    dec_f='FDEDD4', dec_l='B2682E', dec_b='FBD5AE', dec_bl='AF6226',
    enc_f='DCEFF6', enc_l='2A6190', enc_b='C0DBF0', enc_bl='265A8A',
    att_f='EDDCEE', att_l='6F2379',
    cls_f='E6E6E6', cls_l='8D8D8D', cls_b='C8C8C8', cls_bl='585656',
    grp_f='E4F2E3', grp_l='206530',
    dark='515151', wire='000000', dash='5F5F5F', graytxt='7F7F7F',
)

app = win32com.client.Dispatch('Visio.InvisibleApp')
app.AlertResponse = 2
doc = app.Documents.Add('')
page = doc.Pages(1)
page.PageSheet.Cells('PageWidth').FormulaU = fx(W)
page.PageSheet.Cells('PageHeight').FormulaU = fx(H)

_runs = {}
_tno = [0]
_dims = {}

def rect(x, y, w, h, fill, line, lw=4, r=22, name=''):
    s = page.DrawRectangle(ix(x), iy(y + h), ix(x + w), iy(y))
    s.Cells('FillForegnd').FormulaU = hx(fill)
    s.Cells('LineColor').FormulaU = hx(line)
    s.Cells('LineWeight').FormulaU = f'{lw * MM:g} mm'
    s.Cells('Rounding').FormulaU = f'{r * MM:g} mm'
    if name:
        s.Name = name
        _dims[name] = (w, h)
    return s

def invisible_rect(x, y, w, h, name=''):
    s = page.DrawRectangle(ix(x), iy(y + h), ix(x + w), iy(y))
    s.Cells('LinePattern').FormulaU = '0'
    s.Cells('FillPattern').FormulaU = '0'
    if name:
        s.Name = name
    return s

def wire(pts, color=C['wire'], lw=6, dashed=False, arrow=False, asize=3):
    for i in range(len(pts) - 1):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        s = page.DrawLine(ix(x1), iy(y1), ix(x2), iy(y2))
        s.Cells('LineColor').FormulaU = hx(color)
        s.Cells('LineWeight').FormulaU = f'{lw * MM:g} mm'
        if dashed:
            s.Cells('LinePattern').FormulaU = '2'
            s.Cells('LineCap').FormulaU = '1'
        if arrow and i == len(pts) - 2:
            s.Cells('EndArrow').FormulaU = '4'
            s.Cells('EndArrowSize').FormulaU = str(asize)

def hop(xc, yc, r=13):
    s = page.DrawArcByThreePoints(ix(xc - r), iy(yc), ix(xc + r), iy(yc),
                                  ix(xc), iy(yc - r))
    s.Cells('LineColor').FormulaU = hx(C['wire'])
    s.Cells('LineWeight').FormulaU = f'{6 * MM:g} mm'

def circle(cx, cy, r, ring, fill='FFFFFF', lw=5):
    s = page.DrawOval(ix(cx - r), iy(cy - r), ix(cx + r), iy(cy + r))
    s.Cells('LineColor').FormulaU = hx(ring)
    s.Cells('LineWeight').FormulaU = f'{lw * MM:g} mm'
    if fill:
        s.Cells('FillForegnd').FormulaU = hx(fill)
    return s

def seg(x1, y1, x2, y2, color=C['wire'], lw=6):
    s = page.DrawLine(ix(x1), iy(y1), ix(x2), iy(y2))
    s.Cells('LineColor').FormulaU = hx(color)
    s.Cells('LineWeight').FormulaU = f'{lw * MM:g} mm'
    return s

def image(name, x, y, w, h):
    s = page.Import(os.path.join(IMG, name + '.png'))
    s.Cells('Width').FormulaU = fx(w)
    s.Cells('Height').FormulaU = fx(h)
    s.Cells('PinX').FormulaU = fx(x + w / 2)
    s.Cells('PinY').FormulaU = fy(y + h / 2)
    return s

def put_text(shape, lines, align='c', angle=0, tw=None, th=None):
    shape.Text = '\n'.join(''.join(t for t, _ in ln) for ln in lines)
    try:
        shape.CellsSRC(5, 0, 0).FormulaU = '1' if align == 'c' else '0'
    except Exception:
        pass
    for c in ('LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin'):
        shape.Cells(c).FormulaU = '0 mm'
    _tno[0] += 1
    name = f'TEXT{_tno[0]:02d}'
    shape.Name = name
    _runs[name] = lines
    if angle:
        shape.Cells('TxtAngle').FormulaU = f'{angle:g} deg'
    bw, bh = (tw, th) if tw else _dims.get(shape.NameU if hasattr(shape, 'NameU') else name, (100, 40))
    if angle:
        bw, bh = bh, bw
    shape.Cells('TxtWidth').FormulaU = fx(bw)
    shape.Cells('TxtHeight').FormulaU = fx(bh)
    return shape

def S(t, sz=40, **o):
    return (t, o | dict(sz=sz))

# ================================================================ containers
rect(1293, 222, 869, 396, C['dec_f'], C['dec_l'], lw=4, r=28, name='DECODER')
rect(489, 706, 796, 411, C['enc_f'], C['enc_l'], lw=4, r=28, name='ENCODER')
rect(1569, 778, 425, 340, C['att_f'], C['att_l'], lw=4, r=28, name='ATTN')
rect(2079, 705, 482, 412, C['cls_f'], C['cls_l'], lw=4, r=28, name='CLSHEAD')
rect(1270, 1259, 727, 331, C['grp_f'], C['grp_l'], lw=4, r=28, name='GRAPH')

o1 = rect(1334, 354, 226, 163, C['dec_b'], C['dec_bl'], lw=4, r=22, name='UP1')
o2 = rect(1641, 354, 225, 163, C['dec_b'], C['dec_bl'], lw=4, r=22, name='UP2')
o3 = rect(1948, 354, 176, 163, C['dec_b'], C['dec_bl'], lw=4, r=22, name='FINAL')
b1 = rect(523, 793, 246, 296, C['enc_b'], C['enc_bl'], lw=4, r=22, name='STEM')
b2 = rect(827, 793, 263, 296, C['enc_b'], C['enc_bl'], lw=4, r=22, name='MID')
b3 = rect(1154, 793, 108, 296, C['enc_b'], C['enc_bl'], lw=4, r=22, name='EMB')
g1 = rect(2106, 793, 104, 297, C['cls_b'], C['cls_bl'], lw=4, r=16, name='EXIT')
g2 = rect(2257, 793, 60, 297, C['cls_b'], C['cls_bl'], lw=4, r=16, name='GAP')
g3 = rect(2364, 793, 60, 297, C['cls_b'], C['cls_bl'], lw=4, r=16, name='FC')
d1 = rect(2607, 385, 154, 100, C['dark'], C['dark'], lw=1, r=14, name='LREC')
d2 = rect(2607, 878, 154, 128, C['dark'], C['dark'], lw=1, r=14, name='LCLS')
rect(2295, 1418, 465, 177, 'FFFFFF', '000000', lw=3, r=20, name='LEGEND')

put_text(o1, [[S('上采样', 47, b=1)], [S('模块', 47, b=1), S(' 1', 43, f='T', b=1)]], tw=226, th=163)
put_text(o2, [[S('上采样', 47, b=1)], [S('模块', 47, b=1), S(' 2', 43, f='T', b=1)]], tw=225, th=163)
put_text(o3, [[S('最终', 47, b=1)], [S('重建', 47, b=1)]], tw=176, th=163)
put_text(b1, [[S('入口流', 44, b=1)],
              [S('Conv1-2', 42, f='T', b=1), S(' 和', 44, b=1)],
              [S('模块', 44, b=1), S(' 1-3', 42, f='T', b=1)]], tw=246, th=296)
put_text(b2, [[S('中间流', 44, b=1)],
              [S('模块', 44, b=1), S(' 4', 42, f='T', b=1)],
              [S('(728-ch)', 40, f='T', i=1)]], tw=263, th=296)
put_text(b3, [[S('嵌入', 42, b=1)],
              [S('(', 36, f='T'), S('F', 38, f='T', i=1),
               S('enc', 34, f='T', i=1, sub=1), S(')', 36, f='T')]], angle=90, tw=108, th=296)
put_text(g1, [[S('出口流', 42, b=1)],
              [S('(', 30, b=1), S('模块 9-12', 32, b=1), S(')', 30, b=1)]], angle=90, tw=104, th=297)
put_text(g2, [[S('全局平均池化', 36, b=1)]], angle=90, tw=60, th=297)
put_text(g3, [[S('全连接', 38, b=1)]], angle=90, tw=60, th=297)
put_text(d1, [[S('L', 52, f='T', i=1, col='FFFFFF'),
               S('rec', 44, f='T', i=1, sub=1, col='FFFFFF')]], tw=154, th=100)
put_text(d2, [[S('L', 52, f='T', i=1, col='FFFFFF'),
               S('cls', 44, f='T', i=1, sub=1, col='FFFFFF')],
              [S('(BCE)', 38, f='T', col='FFFFFF')]], tw=154, th=128)

# images
image('left', 0, 694, 410, 434)
image('recon', 2204, 0, 282, 360)
image('attention', 1839, 872, 111, 107)
image('graph1', 1323, 1355, 158, 151)
image('graph2', 1552, 1355, 159, 155)
image('graph3', 1780, 1354, 160, 152)

# ================================================================ solid wires
wire([(409, 935), (486, 935)], arrow=True)
wire([(410, 646), (1197, 646)])
hop(1211, 646)
wire([(1225, 646), (1479, 646)])
hop(1492, 646)
wire([(1506, 646), (1662, 646), (1662, 766)], arrow=True)
wire([(410, 646), (410, 1428)])
wire([(410, 1428), (1256, 1428)], arrow=True)
wire([(1210, 432), (1326, 432)], arrow=True)
wire([(1210, 432), (1210, 790)])
wire([(1264, 933), (1388, 933)], arrow=True)
wire([(1455, 933), (1479, 933)])
hop(1492, 933)
wire([(1506, 933), (1562, 933)], arrow=True)
wire([(1492, 519), (1492, 1252)], arrow=True)
wire([(1562, 432), (1633, 432)], arrow=True)
wire([(1868, 432), (1940, 432)], arrow=True)
wire([(2126, 432), (2206, 432)], arrow=True)
wire([(2337, 492), (2337, 651), (1890, 651), (1890, 766)], arrow=True)
wire([(1774, 519), (1774, 766)], arrow=True)
wire([(1996, 933), (2070, 933)], arrow=True)
wire([(2212, 941), (2249, 941)], arrow=True, asize=2)
wire([(2319, 941), (2356, 941)], arrow=True, asize=2)
wire([(771, 941), (819, 941)], arrow=True, asize=2)
wire([(1092, 941), (1146, 941)], arrow=True, asize=2)
wire([(1780, 1259), (1780, 1124)], arrow=True)
wire([(1422, 1259), (1422, 972)], arrow=True)
wire([(1335, 933), (1335, 1252)], arrow=True)
wire([(1717, 925), (1831, 925)], arrow=True)
wire([(2312, 1449), (2372, 1449)], arrow=True, lw=4, asize=2)

# dashed wires
wire([(1752, 354), (1752, 300), (968, 300), (968, 398)], color=C['dash'], lw=4,
     dashed=True, arrow=True)
wire([(1492, 300), (1492, 348)], color=C['dash'], lw=4, dashed=True, arrow=True)
wire([(1208, 433), (1004, 433)], color=C['dash'], lw=4, dashed=True, arrow=True)
wire([(1207, 688), (998, 470)], color=C['dash'], lw=4, dashed=True, arrow=True)
wire([(1205, 1091), (1205, 1236), (1092, 1236)], color=C['dash'], lw=4, dashed=True, arrow=True)
wire([(2489, 433), (2598, 433)], color=C['dash'], lw=4, dashed=True, arrow=True)
wire([(2565, 941), (2598, 941)], color=C['dash'], lw=4, dashed=True, arrow=True)
wire([(2315, 1497), (2372, 1497)], color=C['dash'], lw=4, dashed=True, arrow=True, asize=2)

# circles + glyphs
circle(966, 435, 30, '000000')
seg(947, 437, 985, 437, lw=4)
arc = page.DrawArcByThreePoints(ix(947), iy(439), ix(966), iy(434), ix(956), iy(430))
arc.Cells('LineColor').FormulaU = hx('000000')
arc.Cells('LineWeight').FormulaU = f'{5 * MM:g} mm'
arc = page.DrawArcByThreePoints(ix(966), iy(434), ix(985), iy(439), ix(976), iy(442))
arc.Cells('LineColor').FormulaU = hx('000000')
arc.Cells('LineWeight').FormulaU = f'{5 * MM:g} mm'

circle(1423, 937, 30, '000000')
seg(1403, 937, 1443, 937, lw=5)
seg(1423, 917, 1423, 957, lw=5)

circle(1685, 925, 30, '000000', fill=C['att_f'])
seg(1667, 925, 1703, 925, lw=5)

circle(2360, 1547, 22, '000000', lw=3)
seg(2344, 1547, 2376, 1547, lw=3)
seg(2360, 1531, 2360, 1563, lw=3)

# ================================================================ free texts
def ftext(x, y, w, h, lines, align='c', angle=0):
    s = invisible_rect(x, y, w, h, name=f'FTXT{_tno[0]+1:02d}')
    return put_text(s, lines, align=align, angle=angle, tw=w, th=h)

ftext(650, 372, 226, 60, [[S('对比损失', 50, b=1)]])
ftext(650, 430, 226, 76, [[S('(', 44, f='T'), S('L', 46, f='T', i=1, b=1),
                           S('contra', 40, f='T', i=1, sub=1), S(')', 44, f='T')]])
ftext(605, 584, 260, 68, [[S('原始输入', 48, b=1), S(' (', 44, f='T'), S('I', 46, f='T', i=1),
                           S(')', 44, f='T')]])
ftext(880, 714, 415, 62, [[S('Xception', 44, f='T', b=1), S(' ', 44, f='T', b=1), S('编码器', 54, b=1)]])
ftext(1568, 225, 320, 64, [[S('重建解码器', 60, b=1)]])
ftext(1628, 792, 306, 64, [[S('引导注意力', 56, b=1)]])
ftext(2234, 716, 172, 58, [[S('分类头', 56, b=1)]])
ftext(1484, 1277, 316, 62, [[S('图推理模块', 56, b=1)]])
ftext(1448, 1518, 356, 64, [[S('多尺度图投影', 52, b=1)]])
ftext(1315, 768, 200, 170, [[S('Guided', 40, f='T')], [S('Attention', 40, f='T')],
                            [S('Fusion', 40, f='T')]])
ftext(2197, 384, 300, 56, [[S('Reconstructed', 46, f='T')]])
ftext(2197, 440, 300, 56, [[S('Image', 46, f='T'), S(' (', 44, f='T'), S('I*', 46, f='T', i=1),
                            S(')', 44, f='T')]])
ftext(1810, 998, 170, 104, [[S('注意力', 46, b=1)], [S('图', 46, b=1)]])
ftext(1613, 982, 130, 50, [[S('|', 46, f='T'), S('I', 46, f='T', i=1), S(' − ', 46, f='T'),
                            S('I*', 46, f='T', i=1), S('|', 46, f='T')]])
ftext(88, 1140, 300, 116, [[S('AIS', 44, f='T'), S(' 轨迹图像', 44, b=1)],
                            [S('Image', 44, f='T'), S(' (', 42, f='T'), S('I', 44, f='T', i=1),
                             S(')', 42, f='T')]])
ftext(110, 1245, 262, 52, [[S('201 × 402 × 1', 42, f='T', i=1, col=C['graytxt'])]])
ftext(820, 1188, 260, 58, [[S('对比损失', 48, b=1)]])
ftext(800, 1244, 300, 58, [[S('Loss', 44, f='T'), S(' (', 42, f='T'), S('L', 44, f='T', i=1, b=1),
                            S('contra', 40, f='T', i=1, sub=1), S(')', 42, f='T')]])
ftext(695, 1437, 265, 62, [[S('原始输入', 48, b=1), S(' (', 44, f='T'), S('I', 46, f='T', i=1),
                            S(')', 44, f='T')]])
ftext(1768, 1128, 200, 116, [[S('Fused', 40, f='T')], [S('Features', 40, f='T')]], align='l')
t = invisible_rect(1493, 514, 160, 96, name='FD2')
put_text(t, [[S('F', 44, f='T', i=1, b=1), S('d2', 38, f='T', i=1, sub=1)],
             [S('(256-ch)', 40, f='T', i=1, col=C['graytxt'])]], align='l', tw=160, th=96)
t = invisible_rect(1771, 510, 165, 96, name='FD4')
put_text(t, [[S('F', 44, f='T', i=1, b=1), S('d4', 38, f='T', i=1, sub=1)],
             [S('(128-ch)', 40, f='T', i=1, col=C['graytxt'])]], align='l', tw=165, th=96)
t = invisible_rect(2480, 790, 52, 320, name='YSCORE')
put_text(t, [[S('异常分数', 42, b=1), S(' (', 38, f='T', b=1), S('y', 40, f='T', i=1, b=1),
              S(')', 38, f='T', b=1)]], angle=90, tw=52, th=320)
ftext(2398, 1420, 150, 58, [[S('数据流', 44, b=1)]])
ftext(2390, 1468, 200, 58, [[S('监督/损失', 42, b=1)]])
ftext(2396, 1518, 234, 58, [[S('逐元素相加', 44, b=1)]])

# ================================================================ save
doc.SaveAs(OUTV)
print('saved', OUTV)
json.dump(_runs, open(os.path.join(BASE, 'out', 'text_runs.json'), 'w'),
          ensure_ascii=False, indent=1)
doc.Close()
app.Quit()
print('done pass1')
