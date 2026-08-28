import pandas as pd, numpy as np
from openpyxl.utils import get_column_letter
path = r"C:\Users\bgulick\Desktop\V_single_figures\VCIJplotdata.xlsx"
raw = pd.read_excel(path, sheet_name="uncertainty calcs", header=None)
print("shape:", raw.shape)

# Print every column that has a text label in ANY of the first 3 rows
print("\n=== all labeled columns (text in rows 0-2) ===")
for c in range(raw.shape[1]):
    labels = []
    for r in range(0, 3):
        v = raw.iat[r, c] if r < raw.shape[0] else None
        if isinstance(v, str) and v.strip():
            labels.append(f"r{r}:{v.strip()[:40]}")
    if labels:
        # sample value from row 1 or 3
        sample = raw.iat[1, c] if raw.shape[0] > 1 else None
        print(f"col {c} ({get_column_letter(c+1)}): {' | '.join(labels)}   e.g.={sample}")
