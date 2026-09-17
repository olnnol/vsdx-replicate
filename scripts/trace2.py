# -*- coding: utf-8 -*-
"""Stage 2b: clean wire tracing with thumbnail exclusion; condensed JSON output."""
from PIL import Image
import numpy as np, cv2, json, os

# TODO: point at your source image
SRC = r'./source.png'
OUT = os.path.dirname(os.path.abspath(__file__))
a = np.array(Image.open(SRC).convert('RGB'))
H, W = a.shape[:2]
g = cv2.cvtColor(a, cv2.COLOR_RGB2GRAY)

# attention map bbox inside purple container
purple = np.array([237, 220, 238])
pm = (np.abs(a.astype(int) - purple).max(axis=2) <= 8)
pm[770:1120, 1560:2000] = pm[770:1120, 1560:2000]
dark_in_purple = ((g < 110).astype(np.uint8))[778:1110, 1575:1990]
n_, l_, st_, _ = cv2.connectedComponentsWithStats(dark_in_purple, 8)
att = None
for i in range(1, n_):
    x, y, w, h, ar = st_[i]
    if ar > 3000:
        att = (int(x + 1575), int(y + 778), int(w), int(h))
print('attention map bbox:', att)

EXCLUDE = [(0, 694, 412, 1130), (2202, 0, 2488, 362)]
if att:
    x, y, w, h = att
    EXCLUDE.append((x - 2, y - 2, x + w + 3, y + h + 3))
excl = np.zeros((H, W), np.uint8)
for x0, y0, x1, y1 in EXCLUDE:
    excl[y0:y1, x0:x1] = 1

black = ((g < 90) & (excl == 0)).astype(np.uint8)

def runs(m, min_len):
    idx = np.where(m)[0]
    if len(idx) == 0:
        return []
    out, s, p = [], idx[0], idx[0]
    for v in idx[1:]:
        if v <= p + 1:
            p = v
        else:
            if p - s + 1 >= min_len:
                out.append((int(s), int(p)))
            s = p = v
    if p - s + 1 >= min_len:
        out.append((int(s), int(p)))
    return out

def extract(mask, klen, axis):
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (klen, 1) if axis == 0 else (1, klen))
    w = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k)
    lines = []
    if axis == 0:
        for y in range(H):
            for s, e in runs(w[y] > 0, klen):
                lines.append((y, s, e))
    else:
        for x in range(W):
            for s, e in runs(w[:, x] > 0, klen):
                lines.append((x, s, e))
    return w, lines

def merge(lines, tol=5, gap=15):
    lines.sort()
    out, used = [], [False] * len(lines)
    for i, (p, s, e) in enumerate(lines):
        if used[i]:
            continue
        P, A, B = p, s, e
        used[i] = True
        changed = True
        while changed:
            changed = False
            for j, (p2, s2, e2) in enumerate(lines):
                if used[j]:
                    continue
                if abs(p2 - P) <= tol and s2 <= B + gap and e2 >= A - gap:
                    A, B = min(A, s2), max(B, e2)
                    used[j] = True
                    changed = True
        out.append((int(P), int(A), int(B)))
    return out

hwire, hl = extract(black, 61, 0)
vwire, vl = extract(black, 61, 1)
hl = merge(hl)
vl = merge(vl)

# hops: gaps inside merged lines (where a crossing vertical/horizontal interrupts)
def find_gaps(line, mask, axis):
    p, A, B = line
    if axis == 0:
        seg = mask[p - 2:p + 3, max(A, 0):B + 1].max(axis=0)
    else:
        seg = mask[max(A, 0):B + 1, p - 2:p + 3].max(axis=1)
    idx = np.where(seg == 0)[0]
    gaps = []
    if len(idx):
        s = p_ = idx[0]
        for v in idx[1:]:
            if v <= p_ + 1:
                p_ = v
            else:
                if 4 <= v - s <= 45:
                    gaps.append((int(s + A), int(v + A)))
                s = p_ = v
    return gaps

res = dict(hlines=[], vlines=[])
for ln in hl:
    y, s, e = ln
    if e - s < 70:
        continue
    res['hlines'].append(dict(y=y, x1=s, x2=e, gaps=find_gaps(ln, hwire, 0)))
for ln in vl:
    x, s, e = ln
    if e - s < 70:
        continue
    res['vlines'].append(dict(x=x, y1=s, y2=e, gaps=find_gaps(ln, vwire, 1)))

print('== H wires')
for d in res['hlines']:
    print('  y=%4d x %4d->%4d gaps=%s' % (d['y'], d['x1'], d['x2'], d['gaps']))
print('== V wires')
for d in res['vlines']:
    print('  x=%4d y %4d->%4d gaps=%s' % (d['x'], d['y1'], d['y2'], d['gaps']))

# dashed lines
graym = ((np.abs(a[:, :, 0].astype(int) - a[:, :, 1]) < 14) &
         (np.abs(a[:, :, 1].astype(int) - a[:, :, 2]) < 14) &
         (g > 110) & (g < 215) & (excl == 0)).astype(np.uint8)
hd = cv2.morphologyEx(graym, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (41, 1)))
hd = cv2.morphologyEx(hd, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (81, 1)))
vd = cv2.morphologyEx(graym, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 41)))
vd = cv2.morphologyEx(vd, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 81)))
dhl = []
for y in range(H):
    for s, e in runs(hd[y] > 0, 81):
        dhl.append((y, s, e))
dvl = []
for x in range(W):
    for s, e in runs(vd[:, x] > 0, 81):
        dvl.append((x, s, e))
dhl = merge(dhl)
dvl = merge(dvl)
print('== dashed H')
for y, s, e in sorted(dhl):
    if e - s >= 70:
        print('  y=%4d x %4d->%4d' % (y, s, e))
print('== dashed V')
for x, s, e in sorted(dvl):
    if e - s >= 70:
        print('  x=%4d y %4d->%4d' % (x, s, e))

# arrowheads & symbols: black blobs left after wire removal, in work area
left = cv2.subtract(black, cv2.bitwise_or(hwire, vwire))
left = cv2.morphologyEx(left, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3)))
n, lab, stats, cents = cv2.connectedComponentsWithStats(left, 8)
heads = []
for i in range(1, n):
    x, y, w, h, area = stats[i]
    if 120 <= area <= 4000 and 8 <= w <= 70 and 8 <= h <= 70:
        heads.append(dict(cx=int(cents[i][0]), cy=int(cents[i][1]), x=int(x), y=int(y),
                          w=int(w), h=int(h), area=int(area)))
print('== black blobs (arrowheads/symbols/text-frag) count=%d' % len(heads))
for c in sorted(heads, key=lambda t: (t['cy'] // 50, t['cx'])):
    print('  (%(cx)4d,%(cy)4d) bbox=(%(x)4d,%(y)4d,%(w)3dx%(h)3d) a=%(area)d' % c)

# text blocks: group leftover components (bigger merge) excluding heads
res2 = dict(hlines=res['hlines'], vlines=res['vlines'],
            dh=[list(t) for t in dhl if t[2] - t[1] >= 70],
            dv=[list(t) for t in dvl if t[2] - t[1] >= 70],
            heads=heads, attention=att)
json.dump(res2, open(os.path.join(OUT, 'wires.json'), 'w'), indent=1)
print('saved wires.json')
