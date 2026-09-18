# -*- coding: utf-8 -*-
"""Scrub author metadata from .vsdx files (Visio writes the OS account's
email into docProps/core.xml dc:creator / cp:lastModifiedBy on save).

Usage: python scrub_vsdx.py FILE [FILE...]
       (run before publishing a replica.vsdx anywhere public)
"""
import os, re, shutil, sys, zipfile

FIELDS = ('dc:creator', 'cp:lastModifiedBy', 'dc:rights')

def scrub(path):
    tmp = path + '.scrub'
    changed = []
    with zipfile.ZipFile(path) as zin:
        entries = {n: zin.read(n) for n in zin.namelist()}
    core = entries.get('docProps/core.xml')
    if core is not None:
        xml = core.decode('utf-8')
        for f in FIELDS:
            m = re.search(rf'<{f}>([^<]+)</{f}>', xml)
            if m and m.group(1).strip():
                changed.append(f'{f}={m.group(1)}')
                xml = xml.replace(f'<{f}>{m.group(1)}</{f}>', f'<{f}></{f}>')
        entries['docProps/core.xml'] = xml.encode('utf-8')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, data in entries.items():
            zout.writestr(n, data)
    shutil.move(tmp, path)
    return changed

if __name__ == '__main__':
    for f in sys.argv[1:]:
        ch = scrub(f)
        print(f'{f}: ' + ('removed ' + '; '.join(ch) if ch else 'clean'))
