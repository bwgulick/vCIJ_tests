"""
Tri-panel broken-axis plot of the single-crystal elastic constants vs pressure.

    top    : C11
    middle : C12
    bottom : C44

Each panel overlays the two derivations of the same quantity:
    CK  Cook's method   (P EXP pressures)   red circles
    FS  finite strain   (P FS  pressures)   blue squares

Both carry vertical error bars from the shared ncrt block and horizontal
error bars from the shared pressure uncertainty (ncrt P).  Markers only -
no fit lines - so the two datasets can be compared point-for-point.
"""
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

import vplot_common as vc

# (quantity key, y-axis label).  Column names are built as "<key> CK" / "<key> FS"
# and "ncrt <key>" from these keys.
PANELS = [
    ("C11", "$C_{11}$ (GPa)"),
    ("C12", "$C_{12}$ (GPa)"),
    ("C44", "$C_{44}$ (GPa)"),
]


def main(path=None, show=False):
    df = vc.load_all_data(path)

    s = vc.set_scale((7, 9))          # text/marker scale for the chosen size
    fig, axes = plt.subplots(
        3, 1, figsize=vc.figsize_in((7, 9)), sharex=True,
        gridspec_kw={"hspace": 0.08},
    )

    for ax, (key, ylab) in zip(axes, PANELS):
        ncrt = f"ncrt {key}"

        # --- CK (Cook's method) ---
        if vc.visible("CK"):
            x, y, ye, xe = vc.series_xy(
                df, "P EXP", f"{key} CK", yerrcol=ncrt, xerrcol="ncrt P")
            vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                        color=vc.COL["CK"], marker=vc.MARK["CK"],
                        label=vc.LABEL["CK"], key="CK")

        # --- FS (finite strain) ---
        if vc.visible("FS"):
            x, y, ye, xe = vc.series_xy(
                df, "P FS", f"{key} FS", yerrcol=ncrt, xerrcol="ncrt P")
            vc.draw_pts(ax, x, y, yerr=ye, xerr=xe,
                        color=vc.COL["FS"], marker=vc.MARK["FS"],
                        label=vc.LABEL["FS"], key="FS")

        # --- literature / comparison series (markers only, no error bars) ---
        # drawn per panel only where that source reports the constant.
        for k, cfg in vc.LIT_SERIES.items():
            if not vc.visible(k):
                continue                      # deselected in the GUI
            ycol = f"{key} {cfg['suffix']}"
            if cfg["xcol"] not in df.columns or ycol not in df.columns:
                continue                      # this paper doesn't report this constant
            x, y, _, _ = vc.series_xy(df, cfg["xcol"], ycol)   # no error columns
            if x.size == 0:
                continue                      # column present but empty
            vc.draw_pts(ax, x, y,
                        color=vc.COL[k], marker=vc.MARK[k],
                        label=vc.LABEL[k], key=k)

        ax.set_ylabel(vc.ylabel_for(key, ylab), fontsize=12 * s)
        ax.tick_params(axis="both", labelsize=10 * s)
        ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
        ax.margins(y=0.15)
        vc.apply_ylim(ax, key)          # GUI-set y bounds (blank => auto)

    axes[-1].set_xlabel("Pressure (GPa)", fontsize=12 * s)
    vc.apply_xlim(axes[-1])             # shared x bounds (sharex propagates)

    # single combined legend (dedup handles) on the top panel.  Collect across
    # ALL panels so a literature series that appears only on C12/C44 (not C11)
    # still gets a legend entry.
    handles, labels = [], []
    for ax in axes:
        for h, l in zip(*ax.get_legend_handles_labels()):
            if l not in labels:
                handles.append(h); labels.append(l)
    vc.place_legend(axes[0], handles, labels, 10 * s)

    # break marks between each stacked pair, then exterminate
    vc.add_break_marks(axes[0], axes[1])
    vc.add_break_marks(axes[1], axes[2])
    vc.exterminate_ticks(list(axes))

    out = vc.output_path("cij_tri.png")
    fig.savefig(out, dpi=vc.DPI, bbox_inches="tight")
    print(f"CK/FS points -> {out}")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
