"""
Shared helpers for the vanadium elastic-property plots.

Data source is VCIJplotdata.xlsx, sheet "All", which holds three datasets in
side-by-side blocks (header on row 0):

  CK block   : P EXP + C11/C12/C44/C'/Anisotropy/GV/GR/GH/K  (Cook's method, 22 pts)
  ncrt block : ncrt P + ncrt C11 ... ncrt K                  (shared error bars)
  poly block : Ppoly, ncrt Ppoly, Kpoly, ncrt Kpoly,
               Gpoly, ncrt Gpoly                             (aggregate, 19 pts)
  FS block   : P FS  + C11/C12/... /K                        (finite strain, 22 pts)

CK and FS have their OWN pressure arrays but SHARE the single ncrt uncertainty
block (the uncertainties are identical for the two derivations).  The poly
columns carry their own single pressure (Ppoly) and their own uncertainties.

This module only loads data and provides the broken-axis machinery; the
plot_*.py scripts do the drawing.
Style note: axes are never made bold unless explicitly requested.
"""
import os
import numpy as np
import pandas as pd

# ----------------------------------------------------------------------
# Data location.  Override with the VPLOT_XLSX env var if the file moves.
# ----------------------------------------------------------------------
DEFAULT_XLSX = os.environ.get(
    "VPLOT_XLSX", r"C:\Users\bgulick\Downloads\VCIJplotdata.xlsx"
)
ALL_SHEET = "All"

# ----------------------------------------------------------------------
# Colour palette - series are distinguished by DATASET (CK vs FS vs poly),
# not by quantity, because each panel now overlays the two derivations.
#
# Default is the Okabe-Ito colour-blind-safe set.  Saturated red-on-blue is
# deliberately avoided: it causes chromostereopsis (a false 3-D "float"
# where the red and blue markers overlap).  The GUI can override COL live
# from PALETTES below or from any custom hex / matplotlib colour name.
# ----------------------------------------------------------------------
COL = {
    "CK":   "#0072B2",   # blue   - Cook's method
    "FS":   "#E69F00",   # orange - finite strain
    "poly": "#009E73",   # green  - polycrystalline aggregate (Kpoly / Gpoly)
}

# marker per dataset (shape is a second, redundant cue - helps in greyscale
# and for viewers who can't separate the hues)
MARK = {"CK": "o", "FS": "s", "poly": "D"}

# ----------------------------------------------------------------------
# Editable legend names per dataset.  These are the "base" tags the GUI lets
# the user rename; the plot scripts prepend the quantity where needed
# (e.g. the moduli panels build "K CK", "GVRH CK", "K PC", "G PC").  The
# tri-plot uses them verbatim.  The GUI writes chosen names into LABEL
# before each render, so renaming a series updates every legend.
# ----------------------------------------------------------------------
LABEL = {"CK": "CK", "FS": "FS", "poly": "PC"}

# ----------------------------------------------------------------------
# Ready-made colour-blind-friendly palettes (CK, FS, poly).  All chosen to
# avoid the saturated red+blue pairing.  The GUI exposes these by name.
# ----------------------------------------------------------------------
PALETTES = {
    "Okabe-Ito (CB-safe)":  {"CK": "#0072B2", "FS": "#E69F00", "poly": "#009E73"},
    "Tol bright (CB-safe)": {"CK": "#4477AA", "FS": "#228833", "poly": "#AA3377"},
    "MATLAB":               {"CK": "#0072BD", "FS": "#D95319", "poly": "#77AC30"},
    "Viridis":              {"CK": "#3B528B", "FS": "#21908C", "poly": "#5DC863"},
    "High contrast":        {"CK": "#000000", "FS": "#E69F00", "poly": "#56B4E9"},
}

# ----------------------------------------------------------------------
# Error-bar visibility, toggled live by the GUI.  When off, that whole
# family of error bars is dropped from every series on the next render.
# ----------------------------------------------------------------------
SHOW_YERR = True    # vertical   (y) error bars
SHOW_XERR = True    # horizontal (x) error bars

# ----------------------------------------------------------------------
# Axis limits, set live by the GUI.  None on either end => autoscale that
# end (so you can pin just a floor or just a ceiling and let the other
# side breathe).
#   XLIM : the shared pressure (x) axis for the whole figure -> [lo, hi]
#   YLIM : per-panel y limits, keyed by panel key (C11, K, ...) -> [lo, hi]
# The plot scripts call apply_xlim / apply_ylim after drawing.
# ----------------------------------------------------------------------
XLIM = [None, None]
YLIM = {}


def apply_xlim(ax):
    """Apply the shared x (pressure) limits, if either end is set."""
    lo, hi = XLIM
    if lo is not None or hi is not None:
        ax.set_xlim(left=lo, right=hi)


def apply_ylim(ax, key):
    """Apply the y limits for panel `key`, if set for that panel."""
    lim = YLIM.get(key)
    if not lim:
        return
    lo, hi = lim
    if lo is not None or hi is not None:
        ax.set_ylim(bottom=lo, top=hi)


def load_all_data(path=None):
    """Return the whole "All" sheet as an all-numeric DataFrame.

    Every column is coerced to float; the blank spacer columns and any
    trailing text rows become NaN and are simply skipped when plotting.
    Blocks of different length (poly has 19 rows vs 22) coexist happily -
    the short columns just carry NaN in the extra rows.
    """
    path = path or DEFAULT_XLSX
    df = pd.read_excel(path, sheet_name=ALL_SHEET, header=0)
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def series_xy(df, xcol, ycol, yerrcol=None, xerrcol=None):
    """Pull (x, y, yerr, xerr) numpy arrays for one series, NaN-masked.

    Rows where x or y is not finite are dropped from every returned array so
    the error-bar lengths always line up with the plotted points.
    """
    x = df[xcol].to_numpy(dtype=float)
    y = df[ycol].to_numpy(dtype=float)
    m = np.isfinite(x) & np.isfinite(y)
    ye = df[yerrcol].to_numpy(dtype=float)[m] if yerrcol else None
    xe = df[xerrcol].to_numpy(dtype=float)[m] if xerrcol else None
    return x[m], y[m], ye, xe


def draw_pts(ax, x, y, yerr=None, xerr=None, *, color, marker="o", label=None):
    """Draw one series as markers with x/y error bars (no connecting line).

    SHOW_YERR / SHOW_XERR gate the vertical / horizontal error bars so the GUI
    can switch each family on or off.  Markers carry no edge (no black border).
    """
    ax.errorbar(
        x, y,
        yerr=yerr if SHOW_YERR else None,
        xerr=xerr if SHOW_XERR else None,
        fmt=marker, color=color,
        markersize=6, capsize=4, elinewidth=1.2,
        markeredgewidth=0, zorder=4, label=label,
    )


def add_break_marks(ax_top, ax_bot, d=0.015, lw=1.0):
    """Draw the diagonal break marks between two vertically-stacked axes."""
    kw = dict(transform=ax_top.transAxes, color="k", clip_on=False, lw=lw)
    ax_top.plot((-d, +d), (-d, +d), **kw)
    ax_top.plot((1 - d, 1 + d), (-d, +d), **kw)
    kw.update(transform=ax_bot.transAxes)
    ax_bot.plot((-d, +d), (1 - d, 1 + d), **kw)
    ax_bot.plot((1 - d, 1 + d), (1 - d, 1 + d), **kw)


def exterminate_ticks(axes):
    """The tick & spine exterminator - MUST be called last.

    `axes` is the top-to-bottom list of stacked panels.  Removes the interior
    spines/ticks so the stack reads as one broken axis: only the very top panel
    keeps its top spine, only the bottom panel keeps its bottom spine + labels.
    """
    # y ticks point inward on both sides of every panel
    for ax in axes:
        ax.tick_params(axis="y", which="both", direction="in", left=True, right=True)

    bot = axes[-1]

    # hide the shared boundary between each stacked pair
    for ax in axes[:-1]:
        ax.spines["bottom"].set_visible(False)
    for ax in axes[1:]:
        ax.spines["top"].set_visible(False)

    # pressure (x) hash marks + labels live ONLY on the bottom panel's bottom
    # axis - no x ticks on any panel top (that is what put stray marks in the
    # gap between the top and middle panels).
    for ax in axes:
        ax.tick_params(axis="x", which="both", direction="in",
                       top=False, bottom=False, labelbottom=False)
    bot.tick_params(axis="x", which="both", direction="in",
                    bottom=True, labelbottom=True)


def output_path(name):
    """Absolute path next to this module for a saved figure."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
