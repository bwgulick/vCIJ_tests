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

# (panel key, y-label, CK column, FS column, ncrt column, poly value col, poly ncrt col)
PANELS = [
    ("K",  "Bulk modulus $K$ (GPa)",        "K CK",  "K FS",  "ncrt K",
     "Kpoly", "ncrt Kpoly", "$K_{poly}$"),
    ("GH", "Shear modulus $G_H$ (GPa)",     "GH CK", "GH FS", "ncrt GH",
     "Gpoly", "ncrt Gpoly", "$G_{poly}$"),
]


def _draw_panel(ax, df, ck_col, fs_col, ncrt_col, poly_col, poly_err, poly_label):
    # --- CK ---
    x, y, ye, xe = vc.series_xy(df, "P EXP", ck_col, yerrcol=ncrt_col, xerrcol="ncrt P")
    vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                color=vc.COL["CK"], marker=vc.MARK["CK"], label="CK (Cook)")

    # --- FS ---
    x, y, ye, xe = vc.series_xy(df, "P FS", fs_col, yerrcol=ncrt_col, xerrcol="ncrt P")
    vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                color=vc.COL["FS"], marker=vc.MARK["FS"], label="FS (finite strain)")

    # --- poly (Kpoly / Gpoly) at its own single pressure ---
    x, y, ye, xe = vc.series_xy(df, "Ppoly", poly_col,
                                yerrcol=poly_err, xerrcol="ncrt Ppoly")
    vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                color=vc.COL["poly"], marker=vc.MARK["poly"], label=poly_label)

    ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
    ax.margins(y=0.15)


def main(path=None, show=False):
    df = vc.load_all_data(path)

    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(7, 8), sharex=True,
        gridspec_kw={"hspace": 0.08},
    )

    (k_key, k_ylab, k_ck, k_fs, k_ncrt, k_pcol, k_perr, k_plab) = PANELS[0]
    (g_key, g_ylab, g_ck, g_fs, g_ncrt, g_pcol, g_perr, g_plab) = PANELS[1]

    _draw_panel(ax1, df, k_ck, k_fs, k_ncrt, k_pcol, k_perr, k_plab)
    _draw_panel(ax2, df, g_ck, g_fs, g_ncrt, g_pcol, g_perr, g_plab)

    ax1.set_ylabel(k_ylab, fontsize=12)
    ax2.set_ylabel(g_ylab, fontsize=12)
    ax2.set_xlabel("Pressure (GPa)", fontsize=12)

    # GUI-set bounds (blank => auto); sharex propagates the x limits
    vc.apply_ylim(ax1, k_key)
    vc.apply_ylim(ax2, g_key)
    vc.apply_xlim(ax2)

    ax1.legend(fontsize=9, frameon=False, loc="upper left", ncol=3)
    ax2.legend(fontsize=9, frameon=False, loc="upper left", ncol=3)

    # clean interior spines/ticks (no diagonal break marks - K and G_H are
    # different quantities, not a single broken axis)
    vc.exterminate_ticks([ax1, ax2])

    fig.suptitle("Vanadium aggregate moduli: CK vs FS with poly overlay",
                 fontsize=13, y=0.93)
    out = vc.output_path("moduli_dual.png")
    fig.savefig(out, dpi=200, bbox_inches="tight")
    print(f"CK/FS/poly points -> {out}")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
