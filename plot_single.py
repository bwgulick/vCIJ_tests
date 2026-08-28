"""
Single-panel plot of ONE elastic quantity vs pressure.

The quantity is chosen live from the GUI (vc.SINGLE) - e.g. C11, C12, C44,
C', Anisotropy, GV, GR, GH, K.  The panel overlays the same series as the
multi-panel figures, wherever the workbook reports the chosen quantity:

    CK    Cook's method       (P EXP pressures)
    FS    finite strain       (P FS  pressures)
    poly  polycrystalline     (Ppoly pressures)  - only for K (Kpoly) and
                                                    GH (Gpoly)
    + any literature series (vc.LIT_SERIES) that reports the quantity.

Styling matches the tri-/dual-panel plots: markers only (no fit lines), x/y
error bars gated by the GUI toggles, and live colours / markers / legend
names / axis bounds / y-label overrides via vplot_common.
"""
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

import vplot_common as vc

# Rebuilt from the chosen quantity each render so the GUI's bounds/label panel
# (which reads module.PANELS) shows a row for the current element.  Seeded with
# a sensible default for import time.
PANELS = [("C11", "$C_{11}$ (GPa)")]


def main(path=None, show=False):
    global PANELS
    df = vc.load_all_data(path)

    key = vc.SINGLE
    ylab, poly_col, poly_err = vc.QUANTITIES.get(
        key, (f"{key} (GPa)", None, None))
    PANELS = [(key, ylab)]

    s = vc.set_scale((7, 5))          # text/marker scale for the chosen size
    fig, ax = plt.subplots(figsize=vc.figsize_in((7, 5)))

    # shared uncertainty columns - use only if present in this workbook
    ncrt_y = f"ncrt {key}"
    ncrt_y = ncrt_y if ncrt_y in df.columns else None
    xe_shared = "ncrt P" if "ncrt P" in df.columns else None

    # --- CK (Cook's method) ---
    if vc.visible("CK") and f"{key} CK" in df.columns:
        x, y, ye, xe = vc.series_xy(
            df, "P EXP", f"{key} CK", yerrcol=ncrt_y, xerrcol=xe_shared)
        vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                    color=vc.COL["CK"], marker=vc.MARK["CK"], label=vc.LABEL["CK"])

    # --- FS (finite strain) ---
    if vc.visible("FS") and f"{key} FS" in df.columns:
        x, y, ye, xe = vc.series_xy(
            df, "P FS", f"{key} FS", yerrcol=ncrt_y, xerrcol=xe_shared)
        vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                    color=vc.COL["FS"], marker=vc.MARK["FS"], label=vc.LABEL["FS"])

    # --- poly overlay (K -> Kpoly, GH -> Gpoly only) ---
    if poly_col and vc.visible("poly") and poly_col in df.columns:
        pe = poly_err if poly_err and poly_err in df.columns else None
        xep = "ncrt Ppoly" if "ncrt Ppoly" in df.columns else None
        x, y, ye, xe = vc.series_xy(
            df, "Ppoly", poly_col, yerrcol=pe, xerrcol=xep)
        vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                    color=vc.COL["poly"], marker=vc.MARK["poly"], label=vc.LABEL["poly"])

    # --- literature / comparison series (markers only, no error bars) ---
    for k, cfg in vc.LIT_SERIES.items():
        if not vc.visible(k):
            continue
        ycol = f"{key} {cfg['suffix']}"
        if cfg["xcol"] not in df.columns or ycol not in df.columns:
            continue                          # this paper doesn't report it
        x, y, _, _ = vc.series_xy(df, cfg["xcol"], ycol)
        if x.size == 0:
            continue                          # column present but empty
        vc.draw_pts(ax, x, y,
                    color=vc.COL[k], marker=vc.MARK[k], label=vc.LABEL[k])

    ax.set_ylabel(vc.ylabel_for(key, ylab), fontsize=12 * s)
    ax.set_xlabel("Pressure (GPa)", fontsize=12 * s)
    ax.tick_params(axis="both", labelsize=10 * s)
    ax.tick_params(axis="both", which="both", direction="in",
                   top=True, right=True)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.margins(y=0.12)

    vc.apply_ylim(ax, key)              # GUI-set y bounds (blank => auto)
    vc.apply_xlim(ax)                   # GUI-set x bounds (blank => auto)

    handles, labels = ax.get_legend_handles_labels()
    if handles:
        vc.place_legend(ax, handles, labels, 10 * s)

    out = vc.output_path("single_element.png")
    fig.savefig(out, dpi=vc.DPI, bbox_inches="tight")
    print(f"{key} points -> {out}")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
