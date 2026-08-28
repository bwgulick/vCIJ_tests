import pandas as pd, numpy as np
path = r"C:\Users\bgulick\Desktop\V_single_figures\VCIJplotdata.xlsx"
raw = pd.read_excel(path, sheet_name="uncertainty calcs", header=None)

def col(i):
    return pd.to_numeric(raw.iloc[1:, i], errors='coerce')

P   = col(0)    # A V100 Pressure
M   = col(12)   # Cook's Length
Q   = col(16)   # L uncert
R   = col(17)   # Volume
S   = col(18)   # V ncrt
T   = col(19)   # V/v0
U   = col(20)   # v/v0 ncrt

df = pd.DataFrame({'P':P,'Len':M,'Luncert':Q,'Vol':R,'Vncrt':S,'VV0':T,'VV0ncrt':U})
df = df.dropna(subset=['VV0'])
print("rows with V/V0:", len(df))
print(df.to_string())

print("\n--- checks (first valid row) ---")
r = df.iloc[0]
print("L^3           =", r['Len']**3, " vs Vol =", r['Vol'])
print("3 L^2 sigmaL  =", 3*r['Len']**2*r['Luncert'], " vs Vncrt =", r['Vncrt'])
print("Vncrt/Vol     =", r['Vncrt']/r['Vol'], " vs VV0ncrt =", r['VV0ncrt'])
print("3*(sigL/L)*VV0=", 3*(r['Luncert']/r['Len'])*r['VV0'], " vs VV0ncrt =", r['VV0ncrt'])

print("\nL uncert unique:", sorted(df['Luncert'].dropna().unique()))

# recompute VV0ncrt two ways across all rows and compare
recomp1 = df['Vncrt']/ (df['Vol'].iloc[0])          # sigmaV / V0
recomp2 = 3*(df['Luncert']/df['Len'])*df['VV0']     # 3 (sigL/L) (V/V0)
print("\nmax|Vncrt/V0 - VV0ncrt| =", (recomp1-df['VV0ncrt']).abs().max())
print("max|3(sigL/L)VV0 - VV0ncrt| =", (recomp2-df['VV0ncrt']).abs().max())

# cross check against All sheet
allv = pd.read_excel(path, sheet_name="All", header=0)[['P EXP','V/V0']].apply(pd.to_numeric, errors='coerce').dropna()
print("\nAll-sheet rows:", len(allv), " unc-sheet V/V0 rows:", len(df))
merged = pd.DataFrame({'P_all':allv['P EXP'].values[:len(df)],'VV0_all':allv['V/V0'].values[:len(df)],
                       'P_unc':df['P'].values,'VV0_unc':df['VV0'].values})
print(merged.to_string())
print("\nmax|VV0_all - VV0_unc| =", (merged['VV0_all']-merged['VV0_unc']).abs().max())
