"""
Dual-panel broken-axis plot of the aggregate moduli vs pressure.

    top    : Bulk modulus  K
    bottom : Shear modulus (Hill)  GH

Each panel overlays three series:
    CK    Cook's method      (P EXP pressures)   red circles
    FS    finite strain      (P FS  pressures)   blue squares
    poly  polycrystalline    (Ppoly pressures)   purple diamonds
          -> Kpoly on the K panel, Gpoly on the GH panel

CK/FS share the ncrt uncertainty block and the ncrt P pressure error.  The
poly points carry their own pressure (Ppoly) and their own uncertainties
(ncrt Ppoly, ncrt Kpoly, ncrt Gpoly).  Markers only, with x/y error bars.
"""
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

import vplot_common as vc

# The legend uses the editable base tags in vc.LABEL verbatim (same three
# entries for both panels), so the figure carries ONE combined legend of three
# series rather than a per-panel "K CK / GVRH CK" set.
#
# (panel key, y-label, CK col, FS col, ncrt col, poly val col, poly ncrt col)
PANELS = [
    ("K",  "Bulk modulus $K$ (GPa)",    "K CK",  "K FS",  "ncrt K",
     "Kpoly", "ncrt Kpoly"),
    ("GH", "Shear modulus $G_H$ (GPa)", "GH CK", "GH FS", "ncrt GH",
     "Gpoly", "ncrt Gpoly"),
]


def _draw_panel(ax, df, key, ck_col, fs_col, ncrt_col, poly_col, poly_err):
    # --- CK ---
    if vc.visible("CK"):
        x, y, ye, xe = vc.series_xy(df, "P EXP", ck_col, yerrcol=ncrt_col, xerrcol="ncrt P")
        vc.draw_pts(ax, x, y, yerr=ye, xerr=xe, color=vc.COL["CK"],
                    marker=vc.MARK["CK"], label=vc.LABEL["CK"], key="CK")

    # --- FS ---
    if vc.visible("FS"):
        x, y, ye, xe = vc.series_xy(df, "P FS", fs_col, yerrcol=ncrt_col, xerrcol="ncrt P")
        vc.draw_pts(ax, x, y, yerr=ye, xerr=xe, color=vc.COL["FS"],
                    marker=vc.MARK["FS"], label=vc.LABEL["FS"], key="FS")

    # --- poly (Kpoly / Gpoly) at its own single pressure ---
    if vc.visible("poly"):
        x, y, ye, xe = vc.series_xy(df, "Ppoly", poly_col,
                                    yerrcol=poly_err, xerrcol="ncrt Ppoly")
        vc.draw_pts(ax, x, y, yerr=ye, xerr=xe, color=vc.COL["poly"],
                    marker=vc.MARK["poly"], label=vc.LABEL["poly"], key="poly")

    # --- literature / comparison series for this quantity (markers only) ---
    # drawn only where the source reports this modulus (K / GH).
    for k, cfg in vc.LIT_SERIES.items():
        if not vc.visible(k):
            continue
        ycol = f"{key} {cfg['suffix']}"
        if cfg["xcol"] not in df.columns or ycol not in df.columns:
            continue
        x, y, _, _ = vc.series_xy(df, cfg["xcol"], ycol)   # no error columns
        if x.size == 0:
            continue
        vc.draw_pts(ax, x, y, color=vc.COL[k], marker=vc.MARK[k],
                    label=vc.LABEL[k], key=k)

    ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
    ax.margins(y=0.15)


def main(path=None, show=False):
    df = vc.load_all_data(path)

    s = vc.set_scale((7, 8))          # text/marker scale for the chosen size
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=vc.figsize_in((7, 8)), sharex=True,
        gridspec_kw={"hspace": 0.08},
    )

    (k_key, k_ylab, k_ck, k_fs, k_ncrt, k_pcol, k_perr) = PANELS[0]
    (g_key, g_ylab, g_ck, g_fs, g_ncrt, g_pcol, g_perr) = PANELS[1]

    _draw_panel(ax1, df, k_key, k_ck, k_fs, k_ncrt, k_pcol, k_perr)
    _draw_panel(ax2, df, g_key, g_ck, g_fs, g_ncrt, g_pcol, g_perr)

    ax1.set_ylabel(vc.ylabel_for(k_key, k_ylab),
                   fontsize=12 * s * vc.AXIS_FONT_SCALE)
    ax2.set_ylabel(vc.ylabel_for(g_key, g_ylab),
                   fontsize=12 * s * vc.AXIS_FONT_SCALE)
    ax2.set_xlabel("Pressure (GPa)", fontsize=12 * s * vc.AXIS_FONT_SCALE)
    ax1.tick_params(axis="both", labelsize=10 * s * vc.AXIS_FONT_SCALE)
    ax2.tick_params(axis="both", labelsize=10 * s * vc.AXIS_FONT_SCALE)

    # GUI-set bounds (blank => auto); sharex propagates the x limits
    vc.apply_ylim(ax1, k_key)
    vc.apply_ylim(ax2, g_key)
    vc.apply_xlim(ax2)

    # single combined legend (both panels' series) on the top panel, matching
    # the tri-plot.  Dedup by label so nothing repeats.
    handles, labels = [], []
    for ax in (ax1, ax2):
        for h, l in zip(*ax.get_legend_handles_labels()):
            if l not in labels:
                handles.append(h); labels.append(l)
    vc.place_legend(ax1, handles, labels, 9 * s,
                    default_ncol=3, default_loc="upper left")

    # diagonal break marks straddling the interior spine (the y axis jumps
    # between the K and G_H panels)
    vc.add_break_marks(ax1, ax2)

    # clean interior spines/ticks - MUST be called after the break marks
    vc.exterminate_ticks([ax1, ax2])

    out = vc.output_path("moduli_dual.png")
    fig.savefig(out, dpi=vc.DPI, bbox_inches="tight")
    print(f"CK/FS/poly points -> {out}")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
