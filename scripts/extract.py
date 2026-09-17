# -*- coding: utf-8 -*-
"""Stage 1 (v3): merge near-duplicate exact colors (noise), then per-cluster components."""
from PIL import Image
import numpy as np, cv2, json, os

# TODO: point at your source image
SRC = r'./source.png'
OUT = os.path.dirname(os.path.abspath(__file__))

im = Image.open(SRC).convert('RGB')
a = np.array(im)
H, W = a.shape[:2]
flat = a.reshape(-1, 3)
colors, counts = np.unique(flat, axis=0, return_counts=True)
order = np.argsort(-counts)

keep = [(colors[i].astype(int), int(counts[i])) for i in order if counts[i] >= 2000]

# union-find merge of colors within max-channel distance <= 4
parent = list(range(len(keep)))
def find(i):
    while parent[i] != i:
        parent[i] = parent[parent[i]]
        i = parent[i]
    return i
for i in range(len(keep)):
    ci = keep[i][0]
    for j in range(i + 1, len(keep)):
        cj = keep[j][0]
        if np.abs(ci - cj).max() <= 4:
            parent[find(i)] = find(j)

clusters = {}
for i, (c, n) in enumerate(keep):
    r = find(i)
    if r not in clusters:
        clusters[r] = dict(count=0, pix=c, members=0)
    clusters[r]['count'] += n
    clusters[r]['members'] += 1

merged = sorted(clusters.values(), key=lambda d: -d['count'])
print(f'{len(keep)} exact colors >=2000px -> {len(merged)} clusters\n')

def components(color, tol=6, min_area=800):
    c = np.array(color, dtype=np.int16)
    d = np.abs(a.astype(np.int16) - c).max(axis=2)
    m = (d <= tol).astype(np.uint8)
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    n, labels, stats, _ = cv2.connectedComponentsWithStats(m, 8)
    out = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if area >= min_area and w >= 20 and h >= 20:
            out.append(dict(x=int(x), y=int(y), w=int(w), h=int(h),
                            area=int(area), fill=round(float(area) / (w * h), 2)))
    out.sort(key=lambda r: -r['area'])
    return out

result = {}
for cl in merged:
    if cl['count'] < 6000:
        continue
    hexc = '#%02X%02X%02X' % tuple(cl['pix'])
    boxes = components(cl['pix'])
    result[hexc] = dict(rgb=[int(v) for v in cl['pix']], count=cl['count'],
                        members=cl['members'], boxes=boxes)
    print(f'{hexc}  n={cl["count"]:7d} ({cl["members"]} exact shades)')
    for b in boxes[:10]:
        print('    x=%(x)4d y=%(y)4d w=%(w)4d h=%(h)4d fill=%(fill)s' % b)

json.dump(result, open(os.path.join(OUT, 'palette.json'), 'w'), indent=1)
print('\nsaved palette.json')
