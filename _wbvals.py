import pandas as pd
path = r"C:\Users\bgulick\Desktop\V_single_figures\VCIJplotdata.xlsx"
df = pd.read_excel(path, sheet_name="All", header=0)
sub = df[['P EXP','V/V0','ncrt P']].apply(pd.to_numeric, errors='coerce')
print("=== All: P EXP, V/V0, ncrt P ===")
print(sub.dropna(how='all').to_string())

for sh in ['HPCAT','CK']:
    print(f"\n\n=== sheet {sh}: columns ===")
    d = pd.read_excel(path, sheet_name=sh, header=0)
    for c in d.columns:
        print(repr(c))
    print(f"--- {sh} head ---")
    print(d.head(8).to_string())
