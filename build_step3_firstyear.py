"""Step 3: fill column G (first.year.proj) rows 2:424 in every compiled sheet.

Resolves each compiled sheet back to its source .xlsm, extracts first year of
projection using priority fallbacks, and writes the value into G2:G424 of the
compiled sheet.
"""
import os, re, warnings, csv
warnings.filterwarnings('ignore')
from openpyxl import load_workbook

BASE = '/home/gabos/projects/external-debt-extraction/after-2018'
COMPILED_DIR = os.path.join(BASE, '_compiled')
SKIP_DIRS = {'Metadata', 'Completed_data_all_backup', '_compiled', '_compiled_step1_backup'}
FILL_ROWS = range(2, 425)  # rows 2..424, column G

def extract_first_year(path):
    """Return (year, source_label) or (None, None)."""
    try:
        wb = load_workbook(path, data_only=True, read_only=True, keep_vba=False)
    except Exception as e:
        return None, f'ERROR_OPEN:{e}'
    sheets = {s.lower(): s for s in wb.sheetnames}
    try:
        # Priority 1: Input 1 - Basics!C18
        if 'input 1 - basics' in sheets:
            ws = wb[sheets['input 1 - basics']]
            v = ws['C18'].value
            if isinstance(v, (int, float)) and 1990 <= v <= 2050:
                return int(v), 'Input 1 - Basics!C18'
        # Priority 2: Output - Submit!C9
        if 'output - submit' in sheets:
            ws = wb[sheets['output - submit']]
            v = ws['C9'].value
            if isinstance(v, (int, float)) and 1990 <= v <= 2050:
                return int(v), 'Output - Submit!C9'
        # Priority 3: PV Stress!D3
        if 'pv stress' in sheets:
            ws = wb[sheets['pv stress']]
            v = ws['D3'].value
            if isinstance(v, (int, float)) and 1990 <= v <= 2050:
                return int(v), 'PV Stress!D3'
        return None, 'MISSING'
    finally:
        wb.close()

def base_id_from_filename(fname):
    """AFG_EBS.18.101, Sup. 1.xlsm -> AFG_EBS.18.101"""
    name = os.path.splitext(fname)[0]
    name = re.split(r'\s*,\s*', name)[0]
    name = re.split(r'[_\s\.]Sup', name, maxsplit=1, flags=re.IGNORECASE)[0]
    return name.strip()

# Build mapping: country -> {file_id -> filepath}
country_map = {}
for d in sorted(os.listdir(BASE)):
    if d in SKIP_DIRS or not os.path.isdir(os.path.join(BASE, d)):
        continue
    cdir = os.path.join(BASE, d)
    m = {}
    for f in os.listdir(cdir):
        if not f.lower().endswith(('.xlsm', '.xlsx')) or f.startswith('~$'):
            continue
        m[base_id_from_filename(f)] = os.path.join(cdir, f)
    country_map[d] = m

report = []
total_ok = 0
total_missing = 0

for fname in sorted(os.listdir(COMPILED_DIR)):
    if not fname.endswith('.xlsx') or fname.startswith('_'):
        continue
    country = os.path.splitext(fname)[0]
    cpath = os.path.join(COMPILED_DIR, fname)
    wb = load_workbook(cpath)
    source_files = country_map.get(country, {})
    changed = False
    for sname in wb.sheetnames:
        if sname == 'EMPTY':
            continue
        # sheet name = <file_id>_<Ext_Debt_Data|Inp_Out_Debt>  (maybe _v2)
        # strip the sheet suffix
        m = re.match(r'^(.+?)_(Ext_Debt_Data|Inp_Out_Debt)(?:_v\d+)?$', sname)
        if not m:
            report.append((country, sname, None, 'SHEETNAME_PARSE_FAIL', ''))
            total_missing += 1
            continue
        file_id = m.group(1)
        src = source_files.get(file_id)
        if not src:
            report.append((country, sname, None, 'NO_SOURCE_MATCH', file_id))
            total_missing += 1
            continue
        year, label = extract_first_year(src)
        if year is None:
            report.append((country, sname, None, label, os.path.basename(src)))
            total_missing += 1
            continue
        ws = wb[sname]
        for r in FILL_ROWS:
            ws.cell(row=r, column=7, value=year)
        changed = True
        report.append((country, sname, year, label, os.path.basename(src)))
        total_ok += 1
    if changed:
        wb.save(cpath)
        print(f'{country}: saved', flush=True)
    else:
        print(f'{country}: no changes', flush=True)

with open(os.path.join(COMPILED_DIR, '_first_year_report.csv'), 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['country', 'sheet', 'first_year_proj', 'source_label', 'source_file'])
    w.writerows(report)

print(f'\nOK: {total_ok}   MISSING: {total_missing}')
print(f'Report: {os.path.join(COMPILED_DIR, "_first_year_report.csv")}')
