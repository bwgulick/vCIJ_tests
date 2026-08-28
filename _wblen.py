import pandas as pd, numpy as np
path = r"C:\Users\bgulick\Desktop\V_single_figures\VCIJplotdata.xlsx"
d = pd.read_excel(path, sheet_name="HPCAT", header=0)
print("=== full HPCAT V columns ===")
v = d[['P V','L V','2TP V','2TS V']].apply(pd.to_numeric, errors='coerce').dropna(how='all')
print(v.to_string())

# test V/V0 = (L/L0)^3 against the All-sheet V/V0
allv = pd.read_excel(path, sheet_name="All", header=0)[['P EXP','V/V0']].apply(pd.to_numeric, errors='coerce').dropna(how='all')
L = v['L V'].to_numpy(); P = v['P V'].to_numpy()
L0 = L[0]
print(f"\nL0 (first L V) = {L0}")
print("\nP V   L V   (L/L0)^3")
for p,l in zip(P,L):
    print(f"{p:6.2f} {l:8.2f}  {(l/L0)**3:.5f}")
print("\nAll-sheet V/V0 for comparison (P EXP, V/V0):")
print(allv.to_string())
