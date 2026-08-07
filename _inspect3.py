import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
fp = "Cij analysis/Vanadium_Cij_Brian_Claude.xlsx"
wf = openpyxl.load_workbook(fp, data_only=False)
wv = openpyxl.load_workbook(fp, data_only=True)
sn="No Vp 100"; sf,sv=wf[sn],wv[sn]
print(f"=== {sn} max_col={get_column_letter(sf.max_column)}")
for c in range(1, column_index_from_string("AI")+1):
    L=get_column_letter(c)
    h1=sf.cell(1,c).value; f2=sf.cell(2,c).value; v2=sv.cell(2,c).value
    if h1 is None and f2 is None and v2 is None: continue
    print(f"{L:>3} h1={str(h1)[:24]:<24} f2={str(f2)[:34]:<34} v2={str(v2)[:14]}")
print("\n-- sigma block FI..FO --")
for c in range(column_index_from_string("FH"), column_index_from_string("FO")+1):
    L=get_column_letter(c)
    h1=sf.cell(1,c).value; f2=sf.cell(2,c).value; v2=sv.cell(2,c).value
    print(f"{L:>3} h1={str(h1)[:20]:<20} f2={str(f2)[:40]:<40} v2={str(v2)[:12]}")
