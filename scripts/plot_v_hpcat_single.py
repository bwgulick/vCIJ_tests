"""
Single plot: length (um) vs pressure (PSI), points colored by temperature.
  X = column 2  "Pressure (PSI)"
  Y = "Distance length" (microns)
  color = column 4 "Temp calib" (C), rounded to nearest 25 C, blue(cold)->red(hot)
Points are in the acquisition order of the sheet (top to bottom).
"""
import os

import openpyxl
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, BoundaryNorm

_REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     os.pardir))
FIGDIR = os.path.join(_REPO, "figures")
os.makedirs(FIGDIR, exist_ok=True)

XLSX, SHEET = os.path.join(_REPO, "data", "v_hpcat.xlsx"), "July (2)"
C_PSI, C_T, C_LEN = 1, 3, 10          # 0-based: col2, col4, "Distance length"

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]

P, T, L = [], [], []
for row in ws.iter_rows(min_row=2, values_only=True):
    psi, t, length = row[C_PSI], row[C_T], row[C_LEN]
    if psi is None or length is None or t is None:
        continue
    P.append(float(psi)); T.append(float(t)); L.append(float(length))
P, T, L = np.array(P), np.array(T), np.array(L)

# round temperature to nearest 25 C for coloring
Tr = np.round(T / 25.0) * 25.0

# blue (room temp) -> purple -> red (hottest); no yellow, all saturated/visible
cmap = LinearSegmentedColormap.from_list("blue_red", ["#1f4fd6", "#8e2fb0", "#d62027"])
lo = np.floor(Tr.min() / 25) * 25
hi = np.ceil(Tr.max() / 25) * 25
bounds = np.arange(lo - 12.5, hi + 12.5 + 25, 25)   # 25 C bins
norm = BoundaryNorm(bounds, cmap.N)

fig, ax = plt.subplots(figsize=(9, 6.5))
sc = ax.scatter(P, L, c=Tr, cmap=cmap, norm=norm, s=80,
                edgecolor="black", linewidth=0.6, zorder=3)

ax.set_xlabel("Pressure (PSI)", fontsize=12)
ax.set_ylabel("Length (µm)", fontsize=12)
ax.set_title("Vanadium sample length vs pressure (colored by temperature)", fontsize=12)
ax.grid(alpha=0.3)
ax.margins(0.05)

cb = fig.colorbar(sc, ax=ax, pad=0.02, ticks=np.arange(lo, hi + 1, 50))
cb.set_label("Temperature (°C, rounded to 25°)", fontsize=11)

fig.tight_layout()
out = os.path.join(FIGDIR, "v_hpcat_single.png")
fig.savefig(out, dpi=200, bbox_inches="tight")
print(f"{len(P)} points ->", out)
