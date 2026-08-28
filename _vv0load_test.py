import pandas as pd, numpy as np
path = r"C:\Users\bgulick\Desktop\V_single_figures\VCIJplotdata.xlsx"

UNC_SHEET = "uncertainty calcs"

def load_vv0_ncrt(path):
    """Return {round(P,4): sigma_vv0} from the 'uncertainty calcs' sheet.

    Uses the 'V/v0' + 'v/v0 ncrt' columns, restricted to the compression block
    (0.5 <= V/v0 <= 1.001) so the unrelated lower block is excluded, and pairs
    each with its pressure from the first 'V100 Pressure' column.
    """
    d = pd.read_excel(path, sheet_name=UNC_SHEET, header=0)
    # locate columns by header (first match wins for duplicated names)
    def find(name):
        for c in d.columns:
            if str(c).strip().lower() == name:
                return c
        return None
    pcol = find("v100 pressure")
    vcol = find("v/v0")
    ecol = find("v/v0 ncrt")
    P = pd.to_numeric(d[pcol], errors="coerce")
    V = pd.to_numeric(d[vcol], errors="coerce")
    E = pd.to_numeric(d[ecol], errors="coerce")
    m = P.notna() & V.between(0.5, 1.001) & E.notna()
    return {round(p, 4): e for p, e in zip(P[m], E[m])}

nc = load_vv0_ncrt(path)
print("n uncertainty points:", len(nc))

allv = pd.read_excel(path, sheet_name="All", header=0)[['P EXP','V/V0','ncrt P']].apply(pd.to_numeric, errors='coerce').dropna(subset=['V/V0'])
print("\n P EXP     V/V0     ncrt P   sig(V/V0)  matched?")
for _,r in allv.iterrows():
    key = round(r['P EXP'],4)
    e = nc.get(key)
    print(f"{r['P EXP']:8.4f} {r['V/V0']:.5f}  {r['ncrt P']:.3f}   "
          f"{('%.5f'%e) if e is not None else '  MISS ':>8}   {'OK' if e is not None else 'NO'}")
