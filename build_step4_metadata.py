"""Step 4: fill dsa.vintage (col B), country.name (col C), country.code (col D)."""
import os, re, warnings
warnings.filterwarnings('ignore')
from openpyxl import load_workbook

BASE = '/home/gabos/projects/external-debt-extraction/after-2018'
COMPILED_DIR = os.path.join(BASE, '_compiled')
SKIP = {'Metadata','Completed_data_all_backup','_compiled','_compiled_step1_backup'}
FILL_ROWS = range(2, 425)

def base_id_from_filename(fname):
    name = os.path.splitext(fname)[0]
    name = re.split(r'\s*,\s*', name)[0]
    name = re.split(r'[_\s\.]Sup', name, maxsplit=1, flags=re.IGNORECASE)[0]
    return name.strip()

def vintage_from_filename(fname):
    """BEN_EBS.18.102, Sup. 1.xlsm -> EBS.18.102, Sup. 1"""
    stem = os.path.splitext(fname)[0]
    m = re.match(r'^[A-Za-z]{3}\s*_\s*(.+)$', stem)
    return m.group(1).strip() if m else stem.strip()

COUNTRY_NAME_FIX = {
    'Afghanisatn': 'Afghanistan',
    "Cote d'lvoire": "Cote d'Ivoire",
    'Tazania': 'Tanzania',
}

def split_country_folder(folder):
    """'Congo DR_COD' -> ('Congo DR', 'COD')"""
    if '_' in folder:
        name, code = folder.rsplit('_', 1)
        name = name.strip()
        name = COUNTRY_NAME_FIX.get(name, name)
        return name, code.strip()
    return folder, ''

# country folder -> {file_id -> filename}
country_files = {}
for d in sorted(os.listdir(BASE)):
    if d in SKIP or not os.path.isdir(os.path.join(BASE, d)): continue
    m = {}
    for f in os.listdir(os.path.join(BASE, d)):
        if not f.lower().endswith(('.xlsm','.xlsx')) or f.startswith('~$'): continue
        m[base_id_from_filename(f)] = f
    country_files[d] = m

total_ok = total_miss = 0
for cf in sorted(os.listdir(COMPILED_DIR)):
    if not cf.endswith('.xlsx') or cf.startswith('_') or cf.startswith('~$'): continue
    folder = os.path.splitext(cf)[0]
    cname, ccode = split_country_folder(folder)
    cpath = os.path.join(COMPILED_DIR, cf)
    wb = load_workbook(cpath)
    file_map = country_files.get(folder, {})
    changed = False
    for sname in wb.sheetnames:
        if sname == 'EMPTY': continue
        m = re.match(r'^(.+?)_(Ext_Debt_Data|Inp_Out_Debt)(?:_v\d+)?$', sname)
        if not m:
            total_miss += 1; continue
        file_id = m.group(1)
        # exact match first, then prefix match (for truncated names like YEM_EBS.16.53_2016)
        src = file_map.get(file_id)
        if not src:
            cands = [fn for fid, fn in file_map.items() if fid.startswith(file_id)]
            src = cands[0] if cands else None
        if not src:
            total_miss += 1; continue
        vintage = vintage_from_filename(src)
        source_kind = m.group(2)  # Ext_Debt_Data or Inp_Out_Debt
        methodo = '2018gn' if source_kind == 'Ext_Debt_Data' else '2013gn'
        ws = wb[sname]
        for r in FILL_ROWS:
            ws.cell(row=r, column=2, value=vintage)   # B dsa.vintage
            ws.cell(row=r, column=3, value=cname)     # C country.name
            ws.cell(row=r, column=4, value=ccode)     # D country.code
            ws.cell(row=r, column=6, value=methodo)   # F methodo.gn
        changed = True
        total_ok += 1
    if changed:
        wb.save(cpath)
        print(f'{folder}: {cname} / {ccode}', flush=True)

print(f'\nOK: {total_ok}   MISS: {total_miss}')
