"""
Shared style, data loading and sanity checks for the diamond figures.

House style is carried over from the vanadium paper's apps/vplot/vplot_common.py
at its journal (fixed-point) settings:
  * Physical Review B widths (3.375 in single column), text at 9 / 8 / 7.5 pt
    (label / tick / legend), marker & line scale 0.6.
  * matplotlib's default sans font + mathtext (vplot sets no rcParams).
  * Okabe-Ito colour-blind-safe hues; edgeless markers with "T" error caps.
  * Inward ticks, frameless legends, "Pressure (GPa)" x label.
Colour per quantity follows the diamond spec: C11 blue, C44 green, C12
vermillion, K black, G purple; McSkimin is grey dashed on every figure.

Data are read live from Diamond_use_cleaned.xlsx (via diamond_eos.load_workbook)
and checked against diamond_eos.GROUND_TRUTH before anything is drawn.
"""
import os
import sys
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator

import diamond_eos as de

# ----------------------------------------------------------------------
# Sizes (vplot journal mode: FONT_PT + FIXED_SCALE, PRB widths)
# ----------------------------------------------------------------------
PRB_SINGLE_COL_IN = 3.375
FONT_PT = {"label": 9.0, "tick": 8.0, "legend": 7.5, "panel": 9.0}
S = 0.6                                  # vplot FIXED_SCALE
MS = 6 * S                               # marker size
LW_FIT = 1.8 * S                         # fit-line width
LW_REF = 1.2 * S                         # reference / extrapolation lines
CAP = 4 * S                              # error-bar cap length
ELW = 1.2 * S                            # error-bar line width
DPI = 600                                # vplot DPI
XLIM = (0, 13)                           # pressure axis on every figure

# ----------------------------------------------------------------------
# Colours (Okabe-Ito)
# ----------------------------------------------------------------------
COL = {
    "C11": "#0072B2",     # blue
    "C44": "#009E73",     # bluish green
    "C12": "#D55E00",     # vermillion
    "K":   "#000000",     # black
    "G":   "#CC79A7",     # reddish purple
    "Vp100": "#0072B2",   # carries C11
    "Vs100": "#009E73",   # carries C44
    "Vp110": "#E69F00",   # orange - mixes all three
    "Mc":  "#7F7F7F",     # McSkimin, always grey dashed
}
MC_STYLE = dict(color=COL["Mc"], linestyle="--", linewidth=LW_REF)
MC_LABEL = "McSkimin & Andreatch 1972"
HU_LABEL = "Hu et al. 2021"
FS_LABEL = "3rd-order FS fit"
DATA_LABEL = "This work"

# V/V0 has no per-point uncertainty in the workbook; spec default.
SIG_VV0 = 0.001


# ----------------------------------------------------------------------
# Drawing helpers
# ----------------------------------------------------------------------
def new_fig(height, nrows=1, **kw):
    """PRB single-column figure, `height` inches tall, nrows stacked (shared x)."""
    fig, axes = plt.subplots(nrows, 1, figsize=(PRB_SINGLE_COL_IN, height),
                             sharex=True, **kw)
    for ax in np.atleast_1d(axes):
        style_axes(ax)
    return fig, axes


def style_axes(ax):
    """Inward ticks on all four spines, minor ticks on, journal font sizes."""
    ax.xaxis.set_minor_locator(AutoMinorLocator())
    ax.yaxis.set_minor_locator(AutoMinorLocator())
    ax.tick_params(axis="both", which="both", direction="in",
                   top=True, right=True, labelsize=FONT_PT["tick"])
    ax.set_xlim(*XLIM)


def set_labels(ax, ylabel, xlabel=True):
    ax.set_ylabel(ylabel, fontsize=FONT_PT["label"])
    if xlabel:
        ax.set_xlabel("Pressure (GPa)", fontsize=FONT_PT["label"])


def draw_data(ax, x, y, yerr=None, *, color, marker="o", label=None,
              filled=True, zorder=5):
    """Measured points: edgeless markers with vertical "T" error bars."""
    line, caplines, barcols = ax.errorbar(
        x, y, yerr=yerr, fmt=marker, color=color, markersize=MS,
        mfc=color if filled else "white", capsize=CAP, elinewidth=ELW,
        markeredgewidth=0 if filled else ELW, zorder=zorder, label=label)
    # markeredgewidth=0 also zeroes the cap thickness (same property); restore
    # it so the caps render as the flat T-ends (see vplot_common.draw_pts).
    for cap in caplines:
        cap.set_markeredgewidth(ELW)
    # the P = 0 point sits on the left spine; draw it whole, not half-clipped
    for art in (line, *caplines, *barcols):
        art.set_clip_on(False)


def legend(ax, handles=None, labels=None, **kw):
    kw.setdefault("loc", "best")
    if handles is None:
        handles, labels = ax.get_legend_handles_labels()
    return ax.legend(handles, labels, fontsize=FONT_PT["legend"], frameon=False,
                     handlelength=2.2, **kw)


def panel_label(ax, text):
    ax.text(0.03, 0.95, text, transform=ax.transAxes, ha="left", va="top",
            fontsize=FONT_PT["panel"], fontweight="bold")


def figures_dir():
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
    os.makedirs(out, exist_ok=True)
    return out


def save(fig, name):
    """Save fig_<name>.pdf and fig_<name>.png into Diamond/figures/."""
    base = os.path.join(figures_dir(), f"fig_{name}")
    for ext in ("pdf", "png"):
        fig.savefig(f"{base}.{ext}", dpi=DPI, bbox_inches="tight", pad_inches=0.02)
    print(f"  -> {base}.pdf / .png")


# ----------------------------------------------------------------------
# Data + checks
# ----------------------------------------------------------------------
def load():
    """Live workbook data, checked against GROUND_TRUTH, plus velocity sigmas.

    Velocity uncertainty: sigma_V = V * sigma_t / t from the two-way travel
    times and the 1-sigma travel-time resolutions on the 'Inputs' sheet.
    """
    d = de.load_workbook()
    if "tP100" not in d:
        sys.exit("ABORT: workbook not read (fell back to the embedded table); "
                 "the figures must be drawn from the live data.")
    check(d)
    d["sVp100"] = d["Vp100"] * d["sig_tP100"] / d["tP100"]
    d["sVs100"] = d["Vs100"] * d["sig_tS100"] / d["tS100"]
    d["sVp110"] = d["Vp110"] * d["sig_tP110"] / d["tP110"]
    return d


def check(d):
    """Diff the live workbook against the embedded ground-truth table.

    Hard aborts (spec section 4):
      * any Cij (C11, C44, C12) or sigma differing by > 0.5 GPa;
      * the velocity equations failing to reproduce the table's velocities
        from the table's Cij and rho to 0.01 km/s.
    Reported but NOT aborted: rho, V/V0 and the three velocities.  The table
    computes rho from Hu's EoS and back-computes the velocities from the
    K-route Cij; the workbook holds the measured Cook-length density and the
    measured travel-time velocities, which is what the figures plot.  Those
    differ by design (Vp110 by up to ~0.15 km/s: the [110] route gives a
    higher C12 than the K route), so they get a loose sanity bound only.
    """
    gt = de.GROUND_TRUTH
    if len(d["P"]) != len(gt):
        sys.exit(f"ABORT: workbook has {len(d['P'])} points, table has {len(gt)}")
    bad = []
    print("  live workbook vs embedded GROUND_TRUTH (max |diff|):")
    for k in ("P", "C11", "sC11", "C44", "sC44", "C12", "sC12"):
        err = np.max(np.abs(d[k] - gt[k]))
        tol = 0.05 if k == "P" else 0.5
        print(f"    {'OK ' if err <= tol else '!! '}{k:6s} {err:7.3f}  (tol {tol})")
        if err > tol:
            bad.append(k)

    v = de.velocities(gt["C11"], gt["C12"], gt["C44"], gt["rho"])
    for k in ("Vp100", "Vs100", "Vp110"):
        err = np.max(np.abs(v[k] - gt[k]))
        print(f"    {'OK ' if err <= 0.01 else '!! '}{k:6s} {err:7.3f}  "
              "(equations vs table, tol 0.01)")
        if err > 0.01:
            bad.append(f"{k} equations")

    for k, tol in (("rho", 0.02), ("VV0", 0.005),
                   ("Vp100", 0.2), ("Vs100", 0.2), ("Vp110", 0.2)):
        err = np.max(np.abs(d[k] - gt[k]))
        print(f"    {'-- ' if err <= tol else '!! '}{k:6s} {err:7.3f}  "
              f"(measured vs Hu-EoS/K-route table; expected to differ, bound {tol})")
        if err > tol:
            bad.append(k)
    if bad:
        sys.exit(f"ABORT: workbook disagrees with GROUND_TRUTH on {bad}")


def print_table(title, cols):
    """Print the plotted values: cols = {name: array}."""
    names = list(cols)
    print(f"  {title}")
    print("    " + " ".join(f"{n:>9s}" for n in names))
    for row in zip(*(np.asarray(cols[n]) for n in names)):
        print("    " + " ".join(f"{x:9.4f}" if abs(x) < 10 else f"{x:9.2f}"
                                for x in row))
