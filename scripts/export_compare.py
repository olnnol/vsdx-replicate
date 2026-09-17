# -*- coding: utf-8 -*-
"""Export the patched vsdx to PDF, rasterize, build comparison images.

Usage: python export_compare.py [--vsdx path] [--src source_image]
       Defaults: <script_dir>/out/output.vsdx and ./source.png.
       Rasterize width follows the source image width (fallback 2800).
"""
import os, sys
import win32com.client
import pythoncom
import fitz
import numpy as np
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
VSDX = os.path.join(BASE, 'out', 'output.vsdx')
# TODO: point at your source image
SRC = r'./source.png'

args = sys.argv[1:]
for i, a in enumerate(args):
    if a == '--vsdx' and i + 1 < len(args):
        VSDX = args[i + 1]
    if a == '--src' and i + 1 < len(args):
        SRC = args[i + 1]
OUTDIR = os.path.dirname(os.path.abspath(VSDX)) or BASE
PDF = os.path.join(OUTDIR, 'render.pdf')

pythoncom.CoInitialize()
app = win32com.client.Dispatch('Visio.InvisibleApp')
app.AlertResponse = 2
doc = app.Documents.Open(os.path.abspath(VSDX))
doc.ExportAsFixedFormat(1, PDF, 1, 0)
doc.Close()
app.Quit()

d = fitz.open(PDF)
p = d[0]
o = Image.open(SRC).convert('RGB')
z = o.width / p.rect.width
pix = p.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
pix.save(os.path.join(OUTDIR, 'render_pdf.png'))
print('rasterized', pix.width, pix.height)

r = Image.open(os.path.join(OUTDIR, 'render_pdf.png')).convert('RGB').resize(o.size)
a, b = np.array(o).astype(int), np.array(r).astype(int)
diff = np.abs(a - b).max(axis=2)
print('mean abs diff:', round(float(diff.mean()), 2), ' px>60:', round(float((diff > 60).mean()) * 100, 2), '%')
ov = np.zeros((*diff.shape, 3), np.uint8)
ov[..., 0] = 255 - np.array(o.convert('L'))
ov[..., 1] = 255 - np.array(r.convert('L'))
Image.fromarray(ov).save(os.path.join(OUTDIR, 'overlay.png'))
sbs = Image.new('RGB', (o.width, o.height * 2 + 8), 'white')
sbs.paste(o, (0, 0))
sbs.paste(r, (0, o.height + 8))
sbs.save(os.path.join(OUTDIR, 'side_by_side.png'))
print('comparison images saved ->', OUTDIR)
