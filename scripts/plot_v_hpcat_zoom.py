"""
Length (um) vs pressure (PSI), colored by temperature (nearest 25 C, blue->red).
Two fixes so the small length changes at the holds are visible:
  1) broken y-axis: compressed top segment (compression curve) + expanded bottom
     segment (the 7000/5000/3000 holds).
  2) within each constant-pressure hold the points are fanned out horizontally in
     acquisition order (heat-up -> peak -> cool-down) so overlapping dots separate.
     Pressure is CONSTANT within a hold; the horizontal spread is cosmetic only.
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
C_PSI, C_T, C_LEN = 1, 3, 10

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]
P, T, L = [], [], []
for row in ws.iter_rows(min_row=2, values_only=True):
    psi, t, length = row[C_PSI], row[C_T], row[C_LEN]
    if psi is None or length is None or t is None:
        continue
    P.append(float(psi)); T.append(float(t)); L.append(float(length))
P, T, L = np.array(P), np.array(T), np.array(L)
Tr = np.round(T / 25.0) * 25.0

# ---- fan out points that share the same pressure (the heat-cycle holds) ----
FAN = 420.0     # +/- PSI cosmetic spread within a hold
Xplot = P.astype(float).copy()
i = 0
while i < len(P):
    j = i
    while j + 1 < len(P) and P[j + 1] == P[i]:
        j += 1
    n = j - i + 1
    if n > 2:                                   # a hold -> fan out in time order
        Xplot[i:j + 1] = P[i] + np.linspace(-FAN, FAN, n)
    i = j + 1

# ---- colors: blue (cold) -> purple -> red (hot), no yellow ----
cmap = LinearSegmentedColormap.from_list("blue_red", ["#1f4fd6", "#8e2fb0", "#d62027"])
lo = np.floor(Tr.min() / 25) * 25
hi = np.ceil(Tr.max() / 25) * 25
bounds = np.arange(lo - 12.5, hi + 12.5 + 25, 25)
norm = BoundaryNorm(bounds, cmap.N)

# ---- broken y-axis: two stacked panels sharing x ----
fig, (axhi, axlo) = plt.subplots(2, 1, sharex=True, figsize=(9.5, 8.0),
                                 gridspec_kw=dict(height_ratios=[1, 1.3], hspace=0.06))

for ax in (axhi, axlo):
    # faint path within each hold to show heat->cool order
    k = 0
    while k < len(P):
        m = k
        while m + 1 < len(P) and P[m + 1] == P[k]:
            m += 1
        if m - k + 1 > 2:
            ax.plot(Xplot[k:m + 1], L[k:m + 1], "-", color="0.7", lw=0.9, zorder=1)
        k = m + 1
    ax.scatter(Xplot, L, c=Tr, cmap=cmap, norm=norm, s=85,
               edgecolor="black", linewidth=0.6, zorder=3)
    ax.grid(alpha=0.3)

# y ranges: top = full compression curve (>=502), bottom = the hold band only.
# break sits at ~500-502 where there is NO data, so nothing is hidden.
axhi.set_ylim(502, 850)
axlo.set_ylim(468, 500)

# hide the shared spine + add diagonal break marks
axhi.spines["bottom"].set_visible(False)
axlo.spines["top"].set_visible(False)
axhi.tick_params(labeltop=False, bottom=False)
d = 0.012
kw = dict(transform=axhi.transAxes, color="k", clip_on=False, lw=1)
axhi.plot((-d, +d), (-d, +d), **kw); axhi.plot((1 - d, 1 + d), (-d, +d), **kw)
kw.update(transform=axlo.transAxes)
axlo.plot((-d, +d), (1 - d * 1.7, 1 + d * 1.7), **kw)
axlo.plot((1 - d, 1 + d), (1 - d * 1.7, 1 + d * 1.7), **kw)

# label the holds
for pval, lab in [(7000, "7000 PSI cycle"), (5000, "5000 PSI cycle"), (3000, "3000 PSI cycle")]:
    axlo.annotate(lab, (pval, 468.4), ha="center", va="bottom", fontsize=8, color="0.3")

axlo.set_xlabel("Pressure (PSI)", fontsize=12)
fig.text(0.035, 0.5, "Length (µm)", va="center", rotation="vertical", fontsize=12)
axhi.set_title("Vanadium length vs pressure (broken y-axis; holds fanned out by time)",
               fontsize=11.5)
axlo.text(0.5, -0.16, "within a hold, points are spread horizontally for clarity — "
          "pressure is constant; order = heat-up → peak → cool-down",
          transform=axlo.transAxes, ha="center", fontsize=8, color="0.4")

sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm); sm.set_array([])
cb = fig.colorbar(sm, ax=(axhi, axlo), pad=0.02, ticks=np.arange(lo, hi + 1, 50))
cb.set_label("Temperature (°C, rounded to 25°)", fontsize=11)

out = os.path.join(FIGDIR, "v_hpcat_zoom.png")
fig.savefig(out, dpi=200, bbox_inches="tight")
print(f"{len(P)} points ->", out)
