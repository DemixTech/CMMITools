"""
injectTemplate.py - Generic template injection utility for CAS portal templates.

Usage: python injectTemplate.py <template_path> <output_path> <json_data>

json_data: JSON array of objects with keys matching column config.
           Each column config entry: [col_letter, key, type]
           type: "ss" = shared string (dropdown), "inline" = free text

Called from TypeScript handlers via execSync.
Prints "ok:<rowcount>" on success.
"""

import sys, zipfile, json, html, re
from io import BytesIO

template_path = sys.argv[1]
output_path   = sys.argv[2]
config        = json.loads(sys.argv[3])   # column config: [[col, key, type], ...]
data_rows     = json.loads(sys.argv[4])   # array of row objects

with zipfile.ZipFile(template_path, 'r') as z:
    sheet1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
    ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')

# Parse existing shared strings
si_texts = re.findall(r'<x:t[^>]*>([^<]*)</x:t>', ss_xml)
str_index = {s: i for i, s in enumerate(si_texts)}
new_strings = list(si_texts)

def get_idx(s):
    s = s.strip()
    if s not in str_index:
        str_index[s] = len(new_strings)
        new_strings.append(s)
    return str_index[s]

def make_row(rn, d):
    cells = ''
    num_cols = len(config)
    spans = f"1:{num_cols}"
    for col_letter, key, cell_type in config:
        val = (d.get(key) or '').strip()
        if not val:
            cells += f'<x:c r="{col_letter}{rn}" s="2" />'
        elif cell_type == 'ss':
            cells += f'<x:c r="{col_letter}{rn}" s="2" t="s"><x:v>{get_idx(val)}</x:v></x:c>'
        else:
            cells += f'<x:c r="{col_letter}{rn}" s="2" t="inlineStr"><x:is><x:t>{html.escape(val)}</x:t></x:is></x:c>'
    return f'<x:row r="{rn}" spans="{spans}" ht="15" customHeight="1">{cells}</x:row>'

# Preserve header row, replace all data rows
header_match = re.search(r'<x:row r="1"[^>]*>.*?</x:row>', sheet1, re.DOTALL)
header_row = header_match.group(0) if header_match else ''
data_rows_xml = ''.join(make_row(i + 2, row) for i, row in enumerate(data_rows))
new_sheetdata = f'<x:sheetData>{header_row}{data_rows_xml}</x:sheetData>'

sd_start = sheet1.find('<x:sheetData>')
sd_end   = sheet1.find('</x:sheetData>') + len('</x:sheetData>')
new_sheet1 = sheet1[:sd_start] + new_sheetdata + sheet1[sd_end:]

# Rebuild shared strings XML
ns_url = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
new_ss = (f'<?xml version="1.0" encoding="utf-8"?>'
          f'<x:sst xmlns:x="{ns_url}" count="{len(new_strings)}" uniqueCount="{len(new_strings)}">')
for s in new_strings:
    new_ss += f'<x:si><x:t>{html.escape(s)}</x:t></x:si>'
new_ss += '</x:sst>'

# Write output ZIP
buf = BytesIO()
with zipfile.ZipFile(template_path, 'r') as zin:
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'xl/worksheets/sheet1.xml':
                zout.writestr(item, new_sheet1.encode('utf-8'))
            elif item.filename == 'xl/sharedStrings.xml':
                zout.writestr(item, new_ss.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item.filename))

with open(output_path, 'wb') as f:
    f.write(buf.getvalue())

print(f"ok:{len(data_rows)}")
