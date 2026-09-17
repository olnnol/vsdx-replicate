# -*- coding: utf-8 -*-
"""Pass 2: rewrite Character sections + Text elements inside the saved vsdx
with exact per-run formatting (font, size, style, pos, color).

Usage: python patch_text.py [vsdx_path] [text_runs.json]
       (defaults: <script_dir>/out/output.vsdx + <script_dir>/out/text_runs.json;
        relative paths resolve against the current working directory.)"""
import json, os, re, shutil, sys, zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
VSDX = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, 'out', 'output.vsdx')
RUNS_PATH = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, 'out', 'text_runs.json')
RUNS = json.load(open(RUNS_PATH, encoding='utf-8'))

FONTS = {'S': 'SimSun', 'T': 'Times New Roman', 'A': 'Arial', 'G': 'Century Gothic'}

def font_name(key):
    """Map short font keys to font names; unknown keys pass through literally
    (e.g. f='Arial' uses Arial directly)."""
    return FONTS.get(key, key)

def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def fmt_key(o):
    return (o.get('f', 'S'), o.get('sz', 40), int(bool(o.get('b'))),
            int(bool(o.get('i'))), int(bool(o.get('sub'))), o.get('col'))

def build_shape_xml(lines):
    # flatten: list of (text_with_newlines, key)
    flat, cur, curk = [], '', None
    def flush():
        nonlocal cur, curk
        if cur:
            flat.append((cur, curk))
        cur, curk = '', None
    for li, ln in enumerate(lines):
        for t, o in ln:
            k = fmt_key(o)
            if k != curk:
                flush()
                curk = k
            cur += t
        cur += '\n'
    flush()
    if flat and flat[-1][0].endswith('\n'):
        t, k = flat[-1]
        flat[-1] = (t.rstrip('\n'), k)
        if not flat[-1][0]:
            flat.pop()

    keys = []
    for _, k in flat:
        if k not in keys:
            keys.append(k)
    rows = []
    for i, (f, sz, b, it, sub, col) in enumerate(keys):
        style = b | (it << 1)
        color = f"#{col}" if col else None
        cells = [
            ('Font', font_name(f), None),
            ('Color', color, None),
            ('Style', str(style), None),
            ('Case', '0', None),
            ('Pos', '2' if sub else '0', None),
            ('FontScale', '1', None),
            ('Size', f'{sz / 254:.16g}', None),
            ('DblUnderline', '0', None),
            ('Overline', '0', None),
            ('Strikethru', '0', None),
            ('DoubleStrikethrough', '0', None),
            ('Letterspace', '0', None),
            ('ColorTrans', '0', None),
            ('AsianFont', 'Themed', 'THEMEVAL()'),
            ('ComplexScriptFont', 'Themed', 'THEMEVAL()'),
        ]
        row = ''.join(
            f"<Cell N='{n}' V='{v}'" + (f" F='{fo}'/>" if fo else '/>')
            for n, v, fo in cells if v is not None)
        rows.append(f"<Row IX='{i}'>{row}</Row>")
    section = "<Section N='Character'>" + ''.join(rows) + '</Section>'

    parts = []
    for t, k in flat:
        i = keys.index(k)
        parts.append(f"<cp IX='{i}'/>{esc(t)}")
    text = '<Text><cp IX=\'0\'/><pp IX=\'0\'/><tp IX=\'0\'/>' + ''.join(parts) + '</Text>'
    return section, text

def patch_shape(xml, name):
    m = re.search(r"<Shape ID='(\d+)'[^>]*?Name='%s'" % re.escape(name), xml)
    if not m:
        print('  !! shape not found:', name)
        return xml
    start = xml.rfind('<Shape', 0, m.end())
    end = xml.find('</Shape>', m.end())
    block = xml[start:end]
    if '<Shape' in block[len('<Shape'):]:
        print('  !! nested shape unexpected for', name)
        return xml
    lines = RUNS[name]
    section, text = build_shape_xml(lines)
    if re.search(r"<Section N='Character'>", block):
        block2 = re.sub(r"<Section N='Character'>.*?</Section>", section, block, flags=re.S)
    else:
        block2 = block.replace('<Text>', section + '<Text>', 1)
    block2 = re.sub(r'<Text>.*?</Text>', text, block2, flags=re.S)
    return xml[:start] + block2 + xml[end:]

def main():
    src = zipfile.ZipFile(VSDX)
    entries = {n: src.read(n) for n in src.namelist()}
    src.close()
    xml = entries['visio/pages/page1.xml'].decode('utf-8')
    for name in RUNS:
        xml = patch_shape(xml, name)
    entries['visio/pages/page1.xml'] = xml.encode('utf-8')
    tmp = VSDX + '.tmp'
    out = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for n, data in entries.items():
        out.writestr(n, data)
    out.close()
    shutil.move(tmp, VSDX)
    print('patched', VSDX)

if __name__ == '__main__':
    main()
