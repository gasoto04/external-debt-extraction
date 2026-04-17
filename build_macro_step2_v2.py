"""Macro Step 2 (revised): trim source data above 'DATA' marker, paste template
cols A-H only, place trimmed data at I1. No description/trs1..3 headers."""
import os, warnings
warnings.filterwarnings('ignore')
from openpyxl import load_workbook

TPL_PATH = '/home/gabos/projects/external-debt-extraction/DSA_Assumptions_ExternalDebtData_DomesticDebtData.xlsx'
COMPILED_DIR = '/home/gabos/projects/external-debt-extraction/after-2018/_compiled'
BACKUP_DIR = '/home/gabos/projects/external-debt-extraction/after-2018/_compiled_macro_step1_backup'

print('Loading TEMPLATE A:H...')
tpl_wb = load_workbook(TPL_PATH, data_only=True)
tpl_ws = tpl_wb['TEMPLATE']
TPL_ROWS, TPL_COLS = 424, 8  # A-H only this time
template = [[tpl_ws.cell(row=r, column=c).value for c in range(1, TPL_COLS + 1)]
            for r in range(1, TPL_ROWS + 1)]
tpl_wb.close()

target_suffixes = ('_Macro_Debt_Data', '_Data_Input')
total = 0; warned = []
for fname in sorted(os.listdir(COMPILED_DIR)):
    if not fname.endswith('.xlsx') or fname.startswith('_') or fname.startswith('~$'):
        continue
    cur_path = os.path.join(COMPILED_DIR, fname)
    bak_path = os.path.join(BACKUP_DIR, fname)
    if not os.path.exists(bak_path):
        warned.append(f'{fname}: no backup')
        continue
    cur_wb = load_workbook(cur_path)
    bak_wb = load_workbook(bak_path, data_only=True)
    n = 0
    for sname in list(cur_wb.sheetnames):
        if not sname.endswith(target_suffixes): continue
        if sname not in bak_wb.sheetnames:
            warned.append(f'{fname}/{sname}: missing in backup')
            continue
        bws = bak_wb[sname]
        # find DATA marker row in col B (search first 40 rows)
        data_row = None
        for r in range(1, min(bws.max_row, 40) + 1):
            for c in range(1, 4):
                v = bws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip().upper() == 'DATA':
                    data_row = r; break
            if data_row: break
        if not data_row:
            warned.append(f'{fname}/{sname}: no DATA marker found')
            continue
        # snapshot trimmed source data (rows after the DATA marker row)
        trimmed = []
        for r in range(data_row + 1, bws.max_row + 1):
            trimmed.append([bws.cell(row=r, column=c).value for c in range(1, bws.max_column + 1)])
        # rebuild sheet
        idx = cur_wb.sheetnames.index(sname)
        del cur_wb[sname]
        new_ws = cur_wb.create_sheet(sname, idx)
        # paste template A:H (cols 1..8)
        for r in range(TPL_ROWS):
            for c in range(TPL_COLS):
                v = template[r][c]
                if v is not None:
                    new_ws.cell(row=r + 1, column=c + 1, value=v)
        # paste trimmed source data at I1 (col 9)
        for ri, row in enumerate(trimmed):
            for ci, val in enumerate(row):
                if val is None: continue
                new_ws.cell(row=1 + ri, column=9 + ci, value=val)
        n += 1
    bak_wb.close()
    cur_wb.save(cur_path)
    print(f'{fname}: {n} macro sheets rebuilt', flush=True)
    total += n

print(f'\nTotal rebuilt: {total}')
if warned:
    print('Warnings:')
    for w in warned: print(' ', w)
