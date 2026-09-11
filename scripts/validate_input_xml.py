from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RID = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def textval(c, shared):
    v = c.find(f'{{{NS}}}v')
    if v is None or v.text is None:
        return None
    return shared[int(v.text)] if c.attrib.get('t') == 's' else v.text


def inspect(path):
    with ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(f'{{{NS}}}si'):
                shared.append(
                    ''.join((t.text or '') for t in si.iter(f'{{{NS}}}t'))
                )
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rel = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rels = {x.attrib['Id']: x.attrib['Target'] for x in rel}
        out = []
        for sh in wb.find(f'{{{NS}}}sheets'):
            target = rels[sh.attrib[f'{{{RID}}}id']]
            if not target.startswith('xl/'):
                target = 'xl/' + target
            root = ET.fromstring(z.read(target))
            rows = root.findall(f'.//{{{NS}}}sheetData/{{{NS}}}row')
            vals = []
            for row in rows:
                for c in row.findall(f'{{{NS}}}c'):
                    v = textval(c, shared)
                    if v is not None and v != '':
                        vals.append(v)
            out.append(
                {
                    'sheet': sh.attrib['name'],
                    'rows': len(rows),
                    'nonempty_cells': len(vals),
                }
            )
        return out

root = Path(__file__).resolve().parents[1] / 'input/附件'
for p in sorted(root.glob('附件*.xlsx')):
    print(p.name, inspect(p))
