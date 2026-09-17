# -*- coding: utf-8 -*-
"""Export the patched vsdx to PDF, rasterize at 2800 px, build comparison images."""
import os, sys
import win32com.client
import pythoncom
import fitz
import numpy as np
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
VSDX = os.path.join(BASE, 'out', 'output.vsdx')
PDF = os.path.join(BASE, 'out', 'render.pdf')
# TODO: point at your source image
SRC = r'./source.png'

pythoncom.CoInitialize()
app = win32com.client.Dispatch('Visio.InvisibleApp')
app.AlertResponse = 2
doc = app.Documents.Open(VSDX)
doc.ExportAsFixedFormat(1, PDF, 1, 0)
doc.Close()
app.Quit()

d = fitz.open(PDF)
p = d[0]
z = 2800 / p.rect.width
pix = p.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
pix.save(os.path.join(BASE, 'out', 'render_pdf.png'))
print('rasterized', pix.width, pix.height)

o = Image.open(SRC).convert('RGB')
r = Image.open(os.path.join(BASE, 'out', 'render_pdf.png')).convert('RGB').resize(o.size)
a, b = np.array(o).astype(int), np.array(r).astype(int)
diff = np.abs(a - b).max(axis=2)
print('mean abs diff:', round(float(diff.mean()), 2), ' px>60:', round(float((diff > 60).mean()) * 100, 2), '%')
ov = np.zeros((*diff.shape, 3), np.uint8)
ov[..., 0] = 255 - np.array(o.convert('L'))
ov[..., 1] = 255 - np.array(r.convert('L'))
Image.fromarray(ov).save(os.path.join(BASE, 'out', 'overlay.png'))
sbs = Image.new('RGB', (1400, 1718), 'white')
sbs.paste(o.resize((1400, 859)), (0, 0))
sbs.paste(r.resize((1400, 859)), (0, 859))
sbs.save(os.path.join(BASE, 'out', 'side_by_side.png'))
print('comparison images saved')
