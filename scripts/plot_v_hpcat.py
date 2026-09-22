"""
Plot sample length (um) vs pressure (PSI) for the v_hpcat compression/heat-cycle run.
Temperature cycling is shown via a color map (Temp calib, C) plus the acquisition-order
path with directional arrows.

Two panels:
  A) length vs pressure  (the requested axes) - honest pressures, points colored by T,
     connected in acquisition order so you can see each constant-pressure heat loop.
  B) "unrolled" view: length vs acquisition step (left axis) with pressure overlaid
     (right axis, gray stairs) so the temperature cycling within each pressure hold is
     unambiguous.
"""
import os

import openpyxl
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

_REPO = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     os.pardir))
FIGDIR = os.path.join(_REPO, "figures")
os.makedirs(FIGDIR, exist_ok=True)

XLSX = os.path.join(_REPO, "data", "v_hpcat.xlsx")
SHEET = "July (2)"

# column indices (0-based) in the sheet
C_PSI, C_W, C_T, C_DPX, C_STD, C_LEN, C_NOTE = 1, 2, 3, 8, 9, 10, 11

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]

P, T, L, S, note = [], [], [], [], []
for row in ws.iter_rows(min_row=2, values_only=True):
    psi, t, length, std = row[C_PSI], row[C_T], row[C_LEN], row[C_STD]
    if psi is None or length is None:
        continue
    P.append(float(psi))
    T.append(float(t) if t is not None else np.nan)
    L.append(float(length))
    S.append(float(std) if std is not None else 0.0)
    note.append(row[C_NOTE])

P = np.array(P); T = np.array(T); L = np.array(L); S = np.array(S)
step = np.arange(1, len(P) + 1)
print(f"{len(P)} points loaded")

cmap = plt.get_cmap("inferno")
norm = plt.Normalize(vmin=np.nanmin(T), vmax=np.nanmax(T))

# identify constant-pressure heat-cycle holds (pressure fixed, T excursion)
def find_holds():
    holds = []
    for pval in (7000.0, 5000.0, 3000.0):
        idx = np.where(P == pval)[0]
        # keep the contiguous block that contains a heating excursion
        hot = idx[T[idx] > 40]
        if len(hot) == 0:
            continue
        block = np.arange(hot[0], hot[-1] + 1)
        holds.append((pval, block))
    return holds

HOLDS = find_holds()

fig, (axA, axB, axC) = plt.subplots(1, 3, figsize=(20, 6.4))

# ---------------------------------------------------------------- Panel A
# ordered path colored by temperature (segment colored by mean T of its endpoints)
pts = np.column_stack([P, L]).reshape(-1, 1, 2)
segs = np.concatenate([pts[:-1], pts[1:]], axis=1)
segT = (T[:-1] + T[1:]) / 2
lc = LineCollection(segs, cmap=cmap, norm=norm, linewidths=2.2, alpha=0.85, zorder=1)
lc.set_array(segT)
axA.add_collection(lc)

# error bars (length std) - light gray
axA.errorbar(P, L, yerr=S, fmt="none", ecolor="0.6", elinewidth=0.8,
             capsize=2, zorder=2, alpha=0.7)

sc = axA.scatter(P, L, c=T, cmap=cmap, norm=norm, s=55,
                 edgecolor="black", linewidth=0.5, zorder=3)

# direction arrows along the path (every few steps)
for i in range(0, len(P) - 1, 3):
    dx, dy = P[i + 1] - P[i], L[i + 1] - L[i]
    if dx == 0 and dy == 0:
        continue
    axA.annotate("", xy=(P[i + 1], L[i + 1]), xytext=(P[i], L[i]),
                 arrowprops=dict(arrowstyle="-|>", color="0.35", lw=0.8, alpha=0.7),
                 zorder=2)

# label start / end and the max-T point
axA.annotate("start\n(0 PSI, RT)", (P[0], L[0]), textcoords="offset points",
             xytext=(8, 6), fontsize=9, color="navy")
axA.annotate("end", (P[-1], L[-1]), textcoords="offset points",
             xytext=(8, -14), fontsize=9, color="navy")
imax = int(np.nanargmax(T))
axA.annotate(f"hottest {T[imax]:.0f} C", (P[imax], L[imax]),
             textcoords="offset points", xytext=(-70, 0), fontsize=9, color="darkred")

axA.set_xlabel("Pressure (PSI)")
axA.set_ylabel("Length (µm)")
axA.set_title("Length vs pressure\n(arrows = acquisition order; color = temperature)")
axA.margins(0.08)
axA.grid(alpha=0.25)

cb = fig.colorbar(sc, ax=axA, pad=0.01)
cb.set_label("Temperature (°C)")

# zoom inset on the hold region (where all the heat cycling happens)
from mpl_toolkits.axes_grid1.inset_locator import inset_axes, mark_inset
axins = inset_axes(axA, width="55%", height="42%", loc="upper right", borderpad=1.2)
axins.add_collection(LineCollection(segs, cmap=cmap, norm=norm, linewidths=1.6,
                                    alpha=0.85))
axins.get_children()[0].set_array(segT)
axins.scatter(P, L, c=T, cmap=cmap, norm=norm, s=30, edgecolor="black", linewidth=0.4)
axins.set_xlim(2600, 7250); axins.set_ylim(466, 512)
axins.grid(alpha=0.25); axins.tick_params(labelsize=7)
mark_inset(axA, axins, loc1=2, loc2=4, fc="none", ec="0.5", lw=0.8)

# ---------------------------------------------------------------- Panel B
# length vs temperature: hysteresis loop per pressure hold (heat-up vs cool-down)
hold_colors = {7000.0: "#d1495b", 5000.0: "#e08e0b", 3000.0: "#2e7d32"}
for pval, block in HOLDS:
    tt, ll = T[block], L[block]
    col = hold_colors.get(pval, "gray")
    # order along the leg is acquisition order (already time-ordered)
    axB.plot(tt, ll, "-o", color=col, ms=5, lw=1.4, alpha=0.9,
             label=f"{int(pval)} PSI (max {tt.max():.0f} °C)")
    # arrows to show heat-up -> cool-down direction
    for i in range(0, len(block) - 1, 2):
        axB.annotate("", xy=(tt[i + 1], ll[i + 1]), xytext=(tt[i], ll[i]),
                     arrowprops=dict(arrowstyle="-|>", color=col, lw=0.8, alpha=0.7))
axB.set_xlabel("Temperature (°C)")
axB.set_ylabel("Length (µm)")
axB.set_title("Thermal cycling at each pressure hold\n(loops = heat-up → cool-down hysteresis)")
axB.grid(alpha=0.25)
axB.legend(fontsize=8, loc="upper right")

# ---------------------------------------------------------------- Panel C
# unrolled: length vs step (color = T) + pressure on twin axis
axC.errorbar(step, L, yerr=S, fmt="none", ecolor="0.6", elinewidth=0.8,
             capsize=2, alpha=0.7, zorder=1)
axC.plot(step, L, "-", color="0.75", lw=1.0, zorder=1)
axC.scatter(step, L, c=T, cmap=cmap, norm=norm, s=45,
            edgecolor="black", linewidth=0.4, zorder=3)
axC.set_xlabel("Acquisition step (time →)")
axC.set_ylabel("Length (µm)")
axC.grid(alpha=0.25)

axP = axC.twinx()
axP.step(step, P, where="mid", color="steelblue", lw=1.4, alpha=0.8)
axP.set_ylabel("Pressure (PSI)", color="steelblue")
axP.tick_params(axis="y", labelcolor="steelblue")
axP.set_ylim(-300, 8000)

for pval, block in HOLDS:
    x0, x1 = step[block[0]] - 0.5, step[block[-1]] + 0.5
    axC.axvspan(x0, x1, color="orange", alpha=0.08, zorder=0)
    tpk = block[int(np.nanargmax(T[block]))]
    axC.annotate(f"{int(pval)} PSI\nmax {T[tpk]:.0f} °C", (step[tpk], L[tpk]),
                 textcoords="offset points", xytext=(0, 12),
                 ha="center", fontsize=8, color="darkred")

axC.set_title("Unrolled view: length & pressure vs time\n(color = temperature; shaded = heat cycles)")

fig.tight_layout()
out = os.path.join(FIGDIR, "v_hpcat_length_vs_pressure.png")
fig.savefig(out, dpi=200, bbox_inches="tight")
print("wrote", out)
