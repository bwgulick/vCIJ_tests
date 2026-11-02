"""
Fig 2 - Tri-panel broken-axis plot of the elastic constants vs pressure,
built like the vanadium paper's apps/vplot/plot_cij_tri.py.

    top    : C11
    middle : C44
    bottom : C12

Panel order follows the magnitudes, so the stack reads as one broken y axis:
diamond has C11 > C44 > C12 (vanadium's C11 > C12 > C44 put C12 in the middle).

Each panel:
    markers  : measured C11, C44 ('C12 analysis' cols N, O) and K-route C12
               (col X), each with its vertical sigma (cols R, S, Y)
    solid    : 3rd-order Eulerian finite-strain line, from diamond_eos.FITTED
               (KS0, K0' fixed to Hu et al. 2021)
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

# (quantity key, y-axis label), top to bottom
PANELS = [
    ("C11", "$C_{11}$ (GPa)"),
    ("C44", "$C_{44}$ (GPa)"),
    ("C12", "$C_{12}$ (GPa)"),
]


def main(d=None, show=False):
    print("Fig 2 - Cij triplot")
    d = d if d is not None else dc.load()
    fig, axes = dc.new_fig(5.5, nrows=3, gridspec_kw={"hspace": 0.08})

    fs = de.fs_grid(de.FITTED)
    for ax, (key, ylab) in zip(axes, PANELS):
        ax.plot(fs["P"], fs[key], color=dc.COL[key], linewidth=dc.LW_FIT)
        dc.draw_data(ax, d["P"], d[key], d["s" + key], color=dc.COL[key])
        ax.set_ylabel(ylab, fontsize=dc.FONT_PT["label"])
        ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
        ax.margins(y=0.15)

    axes[-1].set_xlabel("Pressure (GPa)", fontsize=dc.FONT_PT["label"])
    fig.align_ylabels(axes)          # C11's 4-digit ticks push its label out
    grey = "0.3"
    handles = [Line2D([], [], color=grey, marker="o", linestyle="none",
                      markersize=dc.MS, markeredgewidth=0),
               Line2D([], [], color=grey, linewidth=dc.LW_FIT)]
    dc.legend(axes[0], handles, [dc.DATA_LABEL, dc.FS_LABEL], loc="lower right")

    # break marks between each stacked pair, then exterminate (must be last)
    vc.add_break_marks(axes[0], axes[1])
    vc.add_break_marks(axes[1], axes[2])
    vc.exterminate_ticks(list(axes))

    dc.print_table("plotted Cij (measured) and FS line at the same P", {
        "P": d["P"],
        "C11": d["C11"], "C11 FS": np.interp(d["P"], fs["P"], fs["C11"]),
        "C44": d["C44"], "C44 FS": np.interp(d["P"], fs["P"], fs["C44"]),
        "C12": d["C12"], "C12 FS": np.interp(d["P"], fs["P"], fs["C12"])})
    dc.save(fig, "cij_tri")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
