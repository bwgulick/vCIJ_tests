"""
Four vanadium length figures from v_hpcat.xlsx, sheet "July (2)".

  1) v_length_experiment.png : length over the whole experiment (x = acquisition
     order, unlabeled), each dot colored by the pressure it was taken at
     (continuous viridis + colorbar).  The y-axis is BROKEN at 525 um: the top
     third is the coarse ramp region (>525) and the bottom two-thirds is the
     fine region (<525) where almost every point lives, so the heat-cycle
     detail is legible while the curve still reads as one continuous line.

  2-4) v_length_cycle_7000.png / _5000.png / _3000.png : length vs temperature
     for each pressure hold's heat cycle, drawn in shades of red (light = cool,
     dark = hot) so the heat-up/cool-down excursion is visible as a light->dark
     ->light sweep along the connector.  A colorbar gives exact temperature.

Run:  py plot_v_length_four.py
"""
import openpyxl
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize

XLSX, SHEET = "v_hpcat.xlsx", "July (2)"
C_PSI, C_T, C_LEN = 1, 3, 10          # 0-based: Pressure(PSI), Temp calib(C), Distance length(um)

SPLIT = 525.0                          # y-break: coarse (>SPLIT) over fine (<SPLIT)
CYCLE_START = 7                         # first row of the heat-cycling phase; earlier
                                       # rows are the room-temp pressure ramp-up and
                                       # must not leak into the per-hold cycle plots

# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]
P, T, L = [], [], []
for row in ws.iter_rows(min_row=2, values_only=True):
    psi, t, length = row[C_PSI], row[C_T], row[C_LEN]
    if psi is None or length is None or t is None:
        continue
    P.append(float(psi)); T.append(float(t)); L.append(float(length))
P, T, L = np.array(P), np.array(T), np.array(L)
x = np.arange(len(P))                  # acquisition order == "time"


# ===========================================================================
# Figure 1 - whole experiment, broken y-axis, colored by pressure
# ===========================================================================
def figure_experiment():
    cmap = plt.cm.viridis
    norm = Normalize(vmin=0.0, vmax=float(P.max()))

    lo, hi = L.min(), L.max()
    fine_pad = (SPLIT - lo) * 0.10
    coarse_pad = (hi - SPLIT) * 0.05

    # bottom (fine) gets 2/3 of the height, top (coarse) gets 1/3
    fig, (ax_hi, ax_lo) = plt.subplots(
        2, 1, sharex=True, figsize=(11, 6.5),
        gridspec_kw={"height_ratios": [1, 2], "hspace": 0.05})

    for ax in (ax_hi, ax_lo):
        ax.plot(x, L, "--", color="0.5", lw=0.9, zorder=1)     # continuous connector
        sc = ax.scatter(x, L, c=P, cmap=cmap, norm=norm, s=70,
                        edgecolor="black", linewidth=0.6, zorder=3)

    ax_hi.set_ylim(SPLIT, hi + coarse_pad)                     # coarse: the ramp
    ax_lo.set_ylim(lo - fine_pad, SPLIT)                       # fine: the holds

    # hide the shared inner spines so the two panels read as one curve
    ax_hi.spines["bottom"].set_visible(False)
    ax_lo.spines["top"].set_visible(False)
    ax_hi.tick_params(bottom=False)

    # diagonal break marks straddling the cut
    d = 0.012
    kw = dict(transform=ax_hi.transAxes, color="k", clip_on=False, lw=1)
    ax_hi.plot((-d, +d), (-d * 2, +d * 2), **kw)
    ax_hi.plot((1 - d, 1 + d), (-d * 2, +d * 2), **kw)
    kw.update(transform=ax_lo.transAxes)
    ax_lo.plot((-d, +d), (1 - d, 1 + d), **kw)
    ax_lo.plot((1 - d, 1 + d), (1 - d, 1 + d), **kw)

    # x is "time" (acquisition order) - no meaningful tick labels
    ax_lo.set_xlabel("Experiment progression (acquisition order)", fontsize=12)
    ax_lo.tick_params(labelbottom=False)
    for ax in (ax_hi, ax_lo):
        ax.grid(alpha=0.3)
        ax.margins(x=0.03)

    fig.text(0.045, 0.5, "Length (µm)", va="center", rotation="vertical", fontsize=12)
    ax_hi.set_title("Vanadium length over the experiment (color = pressure)", fontsize=13)

    cbar = fig.colorbar(sc, ax=(ax_hi, ax_lo), pad=0.02)
    cbar.set_label("Pressure (PSI)", fontsize=11)

    out = "v_length_experiment.png"
    fig.savefig(out, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"{len(P)} points ->", out)


# ===========================================================================
# Figures 2-4 - length vs temperature per pressure hold, shades of red
# ===========================================================================
def red_ramp():
    """Light (cool) -> dark (hot) single-hue red colormap."""
    return LinearSegmentedColormap.from_list(
        "cool_to_hot", ["#fcbba1", "#fb6a4a", "#cb181d", "#67000d"])


def figure_cycle(hold_psi, ramp):
    idx = np.where((P == hold_psi) & (x >= CYCLE_START))[0]
    if idx.size == 0:
        print(f"no points at {hold_psi} PSI - skipped")
        return
    tt, ll = T[idx], L[idx]
    order = np.argsort(x[idx])          # acquisition order within the hold
    tt, ll = tt[order], ll[order]

    norm = Normalize(vmin=tt.min(), vmax=tt.max())
    fig, ax = plt.subplots(figsize=(8, 6))

    ax.plot(tt, ll, "-", color="0.7", lw=0.9, zorder=1)        # cycle connector
    sc = ax.scatter(tt, ll, c=tt, cmap=ramp, norm=norm, s=90,
                    edgecolor="black", linewidth=0.6, zorder=3)

    ax.set_xlabel("Temperature (°C)", fontsize=12)
    ax.set_ylabel("Length (µm)", fontsize=12)
    ax.set_title(f"Vanadium length vs temperature — {int(hold_psi)} PSI heat cycle",
                 fontsize=13)
    ax.grid(alpha=0.3)
    ax.margins(0.06)

    cbar = fig.colorbar(sc, ax=ax, pad=0.02)
    cbar.set_label("Temperature (°C)  (light = cool → dark = hot)", fontsize=10)

    fig.tight_layout()
    out = f"v_length_cycle_{int(hold_psi)}.png"
    fig.savefig(out, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"{idx.size} points ->", out)


if __name__ == "__main__":
    figure_experiment()
    ramp = red_ramp()
    for hold in (7000, 5000, 3000):
        figure_cycle(hold, ramp)
