import openpyxl
from openpyxl.utils import get_column_letter
fp = "Cij analysis/Vanadium_Cij_Brian_Claude.xlsx"
wf = openpyxl.load_workbook(fp, data_only=False)
wv = openpyxl.load_workbook(fp, data_only=True)
print("SHEETS:", wf.sheetnames)
sn = "No Vs 111"
sf, sv = wf[sn], wv[sn]
print(f"\n=== {sn}  max_col={sf.max_column} ({get_column_letter(sf.max_column)}) max_row={sf.max_row}")
# dump header row1 and row2 formula + row2 value for cols 1..60
for c in range(1, min(sf.max_column,70)+1):
    L = get_column_letter(c)
    h1 = sf.cell(1,c).value
    f2 = sf.cell(2,c).value
    v2 = sv.cell(2,c).value
    if h1 is None and f2 is None and v2 is None: continue
    print(f"{L:>3} h1={str(h1)[:22]:<22} f2={str(f2)[:38]:<38} v2={str(v2)[:16]}")
