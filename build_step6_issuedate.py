"""Step 6: populate issue.date column (col F) using pairs year/month, or internal cell when available."""
import os, re, warnings, datetime
from collections import defaultdict
warnings.filterwarnings('ignore')
from openpyxl import load_workbook

BASE = '/home/gabos/projects/external-debt-extraction/after-2018'
COMPILED_DIR = os.path.join(BASE, '_compiled')
PAIRS_PATH = '/home/gabos/projects/external-debt-extraction/dsa_parent_pairs.xlsx'
SKIP = {'Metadata','Completed_data_all_backup','_compiled','_compiled_step1_backup'}

NAME_MAP = {
    'Congo DR': 'Congo, Dem. Rep',
    'Congo Republic of': 'Congo, Rep',
    "Cote d'Ivoire": "Côte d'Ivoire",
    'Gambia': 'Gambia, The',
    'Micronesia': 'Micronesia, Federated States of',
    'St. Vincent': 'St. Vincent and the Grenadines',
    'Yemen': 'Yemen, Rep',
}
TYPO_FIX = {'Afghanisatn': 'Afghanistan', "Cote d'lvoire": "Cote d'Ivoire", 'Tazania': 'Tanzania'}

# Load pairs: (country_pairs, year, filename) -> (year, month)
wb = load_workbook(PAIRS_PATH, read_only=True)
pairs = defaultdict(lambda: defaultdict(list))  # country -> year -> [(month, filename)]
for row in wb.active.iter_rows(min_row=2, values_only=True):
    fn, country, _, year, month = row[0], row[1], row[2], row[3], row[4]
    if fn and country and year and month:
        try: pairs[country][int(year)].append((int(month), fn))
        except: pass
wb.close()
for c in pairs:
    for y in pairs[c]: pairs[c][y].sort()

SHEET_RE = re.compile(r'^([A-Z]{3})_(EBS|SM)\.(\d{2})\.(\d+)')

def base_id(fname):
    name = os.path.splitext(fname)[0]
    name = re.split(r'\s*,\s*', name)[0]
    name = re.split(r'[_\s\.]Sup', name, maxsplit=1, flags=re.IGNORECASE)[0]
    return name.strip()

# Map: country folder -> {file_id -> source filepath}
country_src = {}
for d in sorted(os.listdir(BASE)):
    if d in SKIP or not os.path.isdir(os.path.join(BASE, d)): continue
    m = {}
    for f in os.listdir(os.path.join(BASE, d)):
        if not f.lower().endswith(('.xlsm','.xlsx')) or f.startswith('~$'): continue
        m[base_id(f)] = os.path.join(BASE, d, f)
    country_src[d] = m

def internal_date(path):
    """Return datetime from Output-Submit!C21 or Data-Input!C19 if present."""
    try:
        wb = load_workbook(path, data_only=True, read_only=True, keep_vba=False)
    except: return None
    sh = {s.lower():s for s in wb.sheetnames}
    for sname, cell in [('output - submit','C21'),('data-input','C19')]:
        if sname in sh:
            try:
                v = wb[sh[sname]][cell].value
                if isinstance(v, datetime.datetime):
                    wb.close(); return v.date()
            except: pass
    wb.close(); return None

counts = defaultdict(int)
for cf in sorted(os.listdir(COMPILED_DIR)):
    if not cf.endswith('.xlsx') or cf.startswith('_') or cf.startswith('~$'): continue
    folder = os.path.splitext(cf)[0]
    cname = TYPO_FIX.get(folder.rsplit('_',1)[0].strip(), folder.rsplit('_',1)[0].strip())
    cpairs = NAME_MAP.get(cname, cname)
    src_map = country_src.get(folder, {})
    path = os.path.join(COMPILED_DIR, cf)
    wb = load_workbook(path)

    # Group sheets by year
    year_groups = defaultdict(list)  # year -> [(doc, sheet, file_id)]
    for s in wb.sheetnames:
        if s == 'EMPTY': continue
        m = SHEET_RE.match(s)
        if not m: continue
        yy = int(m.group(3))
        year = 2000 + yy if yy < 50 else 1900 + yy
        # recover file_id from sheet name (strip _Ext_Debt_Data / _Inp_Out_Debt / _v\d)
        fid = re.sub(r'_(Ext_Debt_Data|Inp_Out_Debt)(?:_v\d+)?$','',s)
        year_groups[year].append((int(m.group(4)), s, fid))

    changed = False
    for year, lst in year_groups.items():
        lst.sort()  # ascending by doc
        plist = pairs.get(cpairs, {}).get(year, [])
        for i, (doc, sname, fid) in enumerate(lst):
            ws = wb[sname]
            date_val = None
            # priority: internal source date
            # resolve source path (try exact match then prefix match)
            src_path = src_map.get(fid)
            if not src_path:
                for k,v in src_map.items():
                    if k.startswith(fid): src_path = v; break
            if src_path:
                d = internal_date(src_path)
                if d:
                    date_val = d
                    counts['internal'] += 1
            if date_val is None and i < len(plist):
                mo, _ = plist[i]
                date_val = datetime.date(year, mo, 1)
                counts['pairs'] += 1
            if date_val is None:
                counts['blank'] += 1
                continue
            for r in range(2, 425):
                ws.cell(row=r, column=6, value=date_val)
            changed = True
    if changed:
        wb.save(path)
        print(f'{cf}: saved', flush=True)

print('\nIssue.date source breakdown:')
for k,v in counts.items(): print(f'  {k}: {v}')
