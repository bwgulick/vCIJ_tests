"""
Fig 6 - Tri-panel broken-axis plot of the measured sound velocities vs
pressure, built like the Cij triplot (fig2_cij_tri.py).

    top    : Vp[110]
    middle : Vp[100]
    bottom : Vs[100]

Panel order follows the magnitudes (~18.4 > ~17.5 > ~12.85 km/s), so the stack
reads as one broken y axis and each velocity gets its own zoomed range.

Each panel:
    markers : measured velocities ('C12 analysis' cols K, L, M), sigma from the
              two-way travel times and their 1-sigma resolutions ('Inputs')
    solid   : FS 3O line - Cij(f) and rho(f) on the strain grid, then
              V = sqrt(C/rho) (diamond_eos.velocities)

Vp[110] is plotted as MEASURED.  The FS line carries the K-route C12, which is
lower than the [110] route gives, so the measured Vp[110] sits ~0.06-0.15 km/s
above its line - the same disagreement that led to the K route for C12.
"""
import os
import sys
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.ticker import MaxNLocator

import diamond_eos as de
import dplot_common as dc

# break marks + interior tick/spine removal come straight from the vanadium code
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                os.pardir, "apps", "vplot"))
import vplot_common as vc  # noqa: E402

# (velocity key, y-axis label), top to bottom
PANELS = [
    ("Vp110", r"$V_P\,[110]$ (km/s)"),
    ("Vp100", r"$V_P\,[100]$ (km/s)"),
    ("Vs100", r"$V_S\,[100]$ (km/s)"),
]


def main(d=None, show=False):
    print("Fig 6 - velocities")
    d = d if d is not None else dc.load()
    fig, axes = dc.new_fig(5.5, nrows=3, gridspec_kw={"hspace": 0.08})

    fs = de.fs_grid(de.FITTED)
    v_fs = de.velocities(fs["C11"], fs["C12"], fs["C44"], fs["rho"])

    for ax, (key, ylab) in zip(axes, PANELS):
        ax.plot(fs["P"], v_fs[key], color=dc.COL[key], linewidth=dc.LW_FIT)
        dc.draw_data(ax, d["P"], d[key], d["s" + key], color=dc.COL[key])
        ax.set_ylabel(ylab, fontsize=dc.FONT_PT["label"])
        ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
        ax.margins(y=0.15)

    axes[-1].set_xlabel("Pressure (GPa)", fontsize=dc.FONT_PT["label"])
    fig.align_ylabels(axes)
    grey = "0.3"
    handles = [Line2D([], [], color=grey, marker="o", linestyle="none",
                      markersize=dc.MS, markeredgewidth=0),
               Line2D([], [], color=grey, linewidth=dc.LW_FIT)]
    dc.legend(axes[0], handles, [dc.DATA_LABEL, dc.FS_LABEL], loc="lower right")

    # break marks between each stacked pair, then exterminate (must be last)
    vc.add_break_marks(axes[0], axes[1])
    vc.add_break_marks(axes[1], axes[2])
    vc.exterminate_ticks(list(axes))

    table = {"P": d["P"]}
    for key, _ in PANELS:
        table[key] = d[key]
        table["s" + key] = d["s" + key]
        table[key + " FS"] = np.interp(d["P"], fs["P"], v_fs[key])
    dc.print_table("plotted velocities (measured) and FS lines", table)
    dc.save(fig, "velocities")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
