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
import re
import numpy as np
import pandas as pd

# ----------------------------------------------------------------------
# Data location.  Override with the VPLOT_XLSX env var if the file moves.
# ----------------------------------------------------------------------
DEFAULT_XLSX = os.environ.get(
    "VPLOT_XLSX", r"C:\Users\bgulick\Desktop\V_single_figures\VCIJplotdata.xlsx"
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
    "vv0":  "#0072B2",   # blue   - measured V/V0 compression points (this study)
    "ding": "#000000",   # black  - Ding et al. 2007 BM3 reference curve (a line)
}

# marker per dataset (shape is a second, redundant cue - helps in greyscale
# and for viewers who can't separate the hues).  "ding" defaults to a solid
# line ("-") so the reference EOS is drawn as a curve, not scattered points.
MARK = {"CK": "o", "FS": "s", "poly": "D", "vv0": "o", "ding": "-"}

# ----------------------------------------------------------------------
# Editable legend names per dataset.  These are the "base" tags the GUI lets
# the user rename; the plot scripts prepend the quantity where needed
# (e.g. the moduli panels build "K CK", "GVRH CK", "K PC", "G PC").  The
# tri-plot uses them verbatim.  The GUI writes chosen names into LABEL
# before each render, so renaming a series updates every legend.
# ----------------------------------------------------------------------
LABEL = {"CK": "Cook", "FS": "Finite strain", "poly": "Gulick et al. 2025",
         "vv0": "This study", "ding": "Ding et al. 2007"}

# ----------------------------------------------------------------------
# Literature / comparison series - AUTO-DISCOVERED from the workbook.
#
# Every comparison source follows the same header convention as CK / FS: a
# pressure column "P <tag>" plus quantity columns "<C11|C12|C44|...> <tag>".
# discover_series() (called by load_all_data on every read) scans the sheet,
# finds each such <tag> that is not one of the core blocks, and registers it
# here with an auto-assigned colour + marker.  So adding a whole new dataset
# needs NO code change: just paste its columns into the "All" sheet with a new
# tag and it appears on every figure and in the GUI.
#
# A panel/quantity is drawn only where that source's column exists and has
# data, so a source that reports only some constants (e.g. C11 + C44 but no
# C12) simply skips the missing ones.  Markers only, no error bars.
#
# LIT_SERIES is filled at load time; each entry is {"xcol", "suffix"}.
# ----------------------------------------------------------------------
LIT_SERIES = {}

# Tags that are NOT comparison sources: the core CK / FS pressure columns
# ("P EXP" -> CK, "P FS" -> FS).  poly / ncrt blocks use their own non-
# "P <tag>" column names and so are never picked up as sources.
_RESERVED_TAGS = {"EXP", "FS"}

# Optional pretty legend names per tag; any tag not listed falls back to the
# raw tag (and is still editable live in the GUI).  Add a line here to give a
# source a nicer default legend name.
LIT_LABELS = {
    "Ant": "Antonangeli et al.",
    "Kat": "Katahara et al.",
}

# Colours + markers handed out (in column order) to auto-discovered sources.
# Chosen to avoid the core CK/FS/poly hues and the saturated red+blue pairing;
# the GUI can still override any of them per source.
_LIT_COLOR_CYCLE = ["#CC79A7", "#56B4E9", "#F0E442", "#D55E00", "#661100",
                    "#332288", "#117733", "#999999", "#000000"]
_LIT_MARKER_CYCLE = ["^", "v", "P", "X", "*", "p", "h", "<", ">", "8"]


def discover_series(df):
    """Scan a loaded "All" DataFrame for comparison sources and register them.

    Finds every "P <tag>" column whose <tag> is not a core block, and adds a
    LIT_SERIES entry plus a default colour / marker / label / visibility for
    it.  Idempotent: existing keys (including GUI-edited styles) are never
    overwritten, so it is safe to call on every render.  Colours/markers are
    assigned in column order, so a given source keeps the same defaults run to
    run.  Returns the list of source tags found (in column order).
    """
    tags = []
    for col in df.columns:
        m = re.match(r"^P (.+)$", str(col).strip())
        if not m:
            continue
        tag = m.group(1).strip()
        if not tag or tag in _RESERVED_TAGS or tag in tags:
            continue
        tags.append(tag)

    for i, tag in enumerate(tags):
        LIT_SERIES[tag] = {"xcol": f"P {tag}", "suffix": tag}
        COL.setdefault(tag,   _LIT_COLOR_CYCLE[i % len(_LIT_COLOR_CYCLE)])
        MARK.setdefault(tag,  _LIT_MARKER_CYCLE[i % len(_LIT_MARKER_CYCLE)])
        LABEL.setdefault(tag, LIT_LABELS.get(tag, tag))
        SHOW.setdefault(tag,  True)
    return tags

# ----------------------------------------------------------------------
# Per-dataset visibility, toggled live by the GUI.  A series whose key maps
# to False is skipped entirely on the next render - no markers and no legend
# entry.  Keyed the same as COL / MARK / LABEL (CK / FS / poly + each
# literature tag); missing keys default to shown via the `visible` helper.
# ----------------------------------------------------------------------
SHOW = {"CK": True, "FS": True, "poly": True,   # sources added by discover_series
        "vv0": True, "ding": True}              # V/V0 figure series


def visible(key):
    """True if series `key` should be drawn (default True for unknown keys)."""
    return SHOW.get(key, True)

# ----------------------------------------------------------------------
# Per-series draw order (z-order), set live by the GUI.  HIGHER draws on
# top; series that share a value fall back to draw order (later = on top).
# Keyed the same as COL / MARK / LABEL (CK / FS / poly / vv0 / ding + each
# literature tag).  A missing key falls back to DEFAULT_ZORDER via
# zorder_for(), so the stack is unchanged until the user picks a layer.
# The plot scripts pass the series key to draw_pts, which looks it up here.
# ----------------------------------------------------------------------
DEFAULT_ZORDER = 4
ZORDER = {}


def zorder_for(key):
    """Draw z-order for series `key` (higher = on top); default if unset."""
    try:
        return int(ZORDER.get(key, DEFAULT_ZORDER))
    except (TypeError, ValueError):
        return DEFAULT_ZORDER

# ----------------------------------------------------------------------
# Single-quantity registry for the "Single element" figure.  Maps a quantity
# key to (default y-axis label, poly value col or None, poly ncrt col or None).
# For every quantity the CK/FS columns are "<key> CK" / "<key> FS" and the
# shared vertical uncertainty is "ncrt <key>" (used where that column exists).
# Only K and GH have a polycrystalline overlay (Kpoly / Gpoly); the rest set
# the poly columns to None so plot_single skips that series for them.
# SINGLE holds the quantity the single-element plot should draw; the GUI sets
# it live before each render.
# ----------------------------------------------------------------------
SINGLE = "C11"

QUANTITIES = {
    "C11":        ("$C_{11}$ (GPa)",             None,    None),
    "C12":        ("$C_{12}$ (GPa)",             None,    None),
    "C44":        ("$C_{44}$ (GPa)",             None,    None),
    "C'":         ("$C'$ (GPa)",                 None,    None),
    "Anisotropy": ("Anisotropy $A$",            None,    None),
    "GV":         ("Shear modulus $G_V$ (GPa)", None,    None),
    "GR":         ("Shear modulus $G_R$ (GPa)", None,    None),
    "GH":         ("Shear modulus $G_H$ (GPa)", "Gpoly", "ncrt Gpoly"),
    "K":          ("Bulk modulus $K$ (GPa)",    "Kpoly", "ncrt Kpoly"),
}

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

# ----------------------------------------------------------------------
# Per-panel y-axis label overrides, set live by the GUI.  Keyed by panel key
# (C11, K, ...) -> label string.  A blank / missing entry falls back to the
# plot module's own default label, so the GUI can let you trim
# "Shear modulus $G_H$ (GPa)" down to just "GH".  The plot scripts pass the
# default through ylabel_for() when calling set_ylabel.
# ----------------------------------------------------------------------
YLABEL = {}

# ----------------------------------------------------------------------
# Legend placement, set live by the GUI.
#   LEGEND_XY   : (x, y) anchor in axes fraction for the legend's upper-left
#                 corner, or None => let matplotlib auto-place it ("best" /
#                 the module's own default corner).  The GUI fills this in
#                 when you type coordinates OR drag the legend on the canvas.
#   LEGEND_NCOL : number of legend columns, or None => each plot's own default.
# The plot scripts draw the combined legend through place_legend() below.
# ----------------------------------------------------------------------
LEGEND_XY = None
LEGEND_NCOL = None


def place_legend(ax, handles, labels, fontsize, *,
                 default_ncol=1, default_loc="best"):
    """Draw the combined legend on `ax`, honouring the GUI's column count and
    position, and make it draggable so it can be moved with the mouse.

    LEGEND_XY None  => matplotlib's automatic placement (default_loc).
    LEGEND_XY (x,y) => pin the legend's upper-left corner there (axes fraction).
    LEGEND_NCOL None => default_ncol (each plot's own preferred column count).

    Returns the Legend so the GUI can read back a dragged position.
    """
    ncol = max(1, int(LEGEND_NCOL if LEGEND_NCOL else default_ncol))
    if LEGEND_XY is not None:
        leg = ax.legend(handles, labels, fontsize=fontsize, frameon=False,
                        ncol=ncol, loc="upper left",
                        bbox_to_anchor=tuple(LEGEND_XY), borderaxespad=0.0)
    else:
        leg = ax.legend(handles, labels, fontsize=fontsize, frameon=False,
                        ncol=ncol, loc=default_loc)
    try:
        leg.set_draggable(True)
    except Exception:
        pass          # older matplotlib without draggable support
    return leg


# ----------------------------------------------------------------------
# Output resolution and figure size, set live by the GUI.
#   DPI     : dots-per-inch used for every savefig (disk + GUI "Save").
#   FIGSIZE : (width_in, height_in) override, or None to let each plot
#             module use its own default size.  The GUI takes centimetres
#             and converts with CM_PER_IN before storing here.
# ----------------------------------------------------------------------
DPI = 600
FIGSIZE = None
CM_PER_IN = 2.54

# Text / marker scale factor.  1.0 at a module's default figure size; shrinks
# with the figure so fonts and markers stay proportional when the GUI resizes
# it.  Each plot's main() sets this via set_scale() before drawing.
SCALE = 1.0


def figsize_in(default):
    """Return the GUI-chosen figure size (inches), or the module default."""
    return FIGSIZE if FIGSIZE else default


def set_scale(default):
    """Set & return the text/marker scale for the current figure size.

    The factor is the geometric mean of the width and height ratios relative
    to `default`, so a figure half as wide and half as tall scales text to
    half size (not a quarter).  Returns 1.0 when no size override is set.
    """
    global SCALE
    w, h = figsize_in(default)
    dw, dh = default
    SCALE = ((w / dw) * (h / dh)) ** 0.5
    return SCALE


def apply_xlim(ax):
    """Apply the shared x (pressure) limits, if either end is set."""
    lo, hi = XLIM
    if lo is not None or hi is not None:
        ax.set_xlim(left=lo, right=hi)


def ylabel_for(key, default):
    """Return the GUI-set y-axis label for panel `key`, or `default`.

    A blank / unset override means "leave the module's default label", so the
    plot scripts can call ax.set_ylabel(ylabel_for(key, default_label)).
    """
    lab = YLABEL.get(key)
    return lab if lab else default


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
    discover_series(df)          # register any "P <tag>" comparison sources
    return df


# Sheet holding the per-point uncertainty propagation (length -> V, density,
# velocities, Cij).  The V/V0 error bars are read from its 'v/v0 ncrt' column,
# which the workbook computes as 3 L^2 sigma_L / V0 from the length uncertainty
# 'L uncert' (sigma_L ~ 1 um on the 0.704 mm sample).  We only READ this sheet,
# never write it, so its formulas / charts stay intact.
UNC_SHEET = "uncertainty calcs"


def load_vv0(path=None):
    """Return (P, vv0, sig_vv0, sigP) numpy arrays for the compression curve.

    P and V/V0 come from the "All" sheet ("P EXP" / "V/V0").  The vertical
    uncertainty sig_vv0 is the length-propagated "v/v0 ncrt" from the
    "uncertainty calcs" sheet, matched to each pressure; sigP is the shared
    "ncrt P".  If the uncertainty sheet or its column is missing, sig_vv0 is
    returned as None so the points are still drawn (without y error bars).
    """
    path = path or DEFAULT_XLSX
    allv = pd.read_excel(path, sheet_name=ALL_SHEET, header=0)
    P = pd.to_numeric(allv.get("P EXP"), errors="coerce")
    V = pd.to_numeric(allv.get("V/V0"), errors="coerce")
    sP = pd.to_numeric(allv.get("ncrt P"), errors="coerce") if "ncrt P" in allv else None

    m = P.notna() & V.notna()
    P, V = P[m].to_numpy(float), V[m].to_numpy(float)
    sP = sP[m].to_numpy(float) if sP is not None else None

    # map pressure -> propagated sigma(V/V0) from the uncertainty sheet
    ncrt = _vv0_ncrt_map(path)
    if ncrt:
        sig = np.array([ncrt.get(round(p, 4), np.nan) for p in P], float)
        if not np.isfinite(sig).any():
            sig = None
    else:
        sig = None
    return P, V, sig, sP


def _vv0_ncrt_map(path):
    """{round(P,4): sigma_vv0} from the uncertainty sheet, or {} if unavailable.

    Restricts to the compression block (0.5 <= V/v0 <= 1.001) so an unrelated
    reuse of the same columns lower in the sheet is excluded, and pairs each
    row with its pressure from the first "V100 Pressure" column.
    """
    try:
        d = pd.read_excel(path, sheet_name=UNC_SHEET, header=0)
    except Exception:
        return {}

    def find(name):
        for c in d.columns:
            if str(c).strip().lower() == name:
                return c
        return None

    pcol, vcol, ecol = find("v100 pressure"), find("v/v0"), find("v/v0 ncrt")
    if not (pcol and vcol and ecol):
        return {}
    Pn = pd.to_numeric(d[pcol], errors="coerce")
    Vn = pd.to_numeric(d[vcol], errors="coerce")
    En = pd.to_numeric(d[ecol], errors="coerce")
    ok = Pn.notna() & Vn.between(0.5, 1.001) & En.notna()
    return {round(p, 4): e for p, e in zip(Pn[ok], En[ok])}


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


# matplotlib line-style codes.  When one of these is passed to draw_pts as the
# `marker`, the series is drawn as a connecting line (no point markers) instead
# of scattered points - handy for smooth literature curves such as Katahara.
# The GUI exposes these alongside the marker shapes in its style dropdown.
LINE_STYLES = {"-", "--", ":", "-."}


def draw_pts(ax, x, y, yerr=None, xerr=None, *, color, marker="o", label=None,
             key=None):
    """Draw one series as markers, OR as a connecting line.

    If `marker` is a matplotlib line-style code ("-", "--", ":", "-.") the
    series is drawn as a line with no point markers (points are sorted by x so
    the line reads left-to-right); otherwise it is drawn as markers with no
    connecting line, exactly as before.  SHOW_YERR / SHOW_XERR gate the
    vertical / horizontal error bars in both cases.  Markers carry no edge.

    `key` is the series key (CK / FS / vv0 / ...).  Its GUI-chosen z-order
    (higher = drawn on top) is looked up via zorder_for(); pass None to keep
    the default depth.
    """
    ye = yerr if SHOW_YERR else None
    xe = xerr if SHOW_XERR else None
    z = zorder_for(key)

    if marker in LINE_STYLES:
        order = np.argsort(x)               # a line must read left-to-right
        x, y = x[order], y[order]
        ye = ye[order] if ye is not None else None
        xe = xe[order] if xe is not None else None
        ax.errorbar(
            x, y, yerr=ye, xerr=xe,
            fmt=marker, color=color,
            linewidth=1.8 * SCALE, capsize=4 * SCALE, elinewidth=1.2 * SCALE,
            zorder=z, label=label,
        )
        return

    ax.errorbar(
        x, y, yerr=ye, xerr=xe,
        fmt=marker, color=color,
        markersize=6 * SCALE, capsize=4 * SCALE, elinewidth=1.2 * SCALE,
        markeredgewidth=0, zorder=z, label=label,
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
