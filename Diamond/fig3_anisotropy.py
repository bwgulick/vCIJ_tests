"""
Fig 3 - Zener anisotropy A = 2 C44 / (C11 - C12) vs pressure.

    markers  : pointwise A from the measured Cij (K-route C12), sigma from
               independent quadrature of sC11, sC12, sC44 (this ignores the
               C11/C12 correlation built in by the K route; a ~15% effect on
               sigma_A, well inside the marker scatter)
    solid    : A(P) from the FS 3O fitted Cij curves (diamond_eos.FITTED)
    grey --  : McSkimin 1972 c0/c0' run through the same FS form with
               Mc K0, K0' - an extrapolation beyond their 0.15 GPa range

A = 1 (elastically isotropic) lies off the spec's 1.15-1.35 range, so no
reference line is drawn; say it in the caption instead.
"""
import numpy as np
import matplotlib.pyplot as plt

import diamond_eos as de
import dplot_common as dc


def sigma_A(C11, C12, C44, s11, s12, s44):
    D = C11 - C12
    return np.sqrt((2 * s44 / D) ** 2 + (2 * C44 / D**2) ** 2 * (s11**2 + s12**2))


def main(d=None, show=False):
    print("Fig 3 - anisotropy")
    d = d if d is not None else dc.load()
    fig, ax = dc.new_fig(2.6)

    fs = de.fs_grid(de.FITTED)
    mc = de.fs_grid(de.MCSKIMIN)
    A_fs = de.anisotropy(fs["C11"], fs["C12"], fs["C44"])
    A_mc = de.anisotropy(mc["C11"], mc["C12"], mc["C44"])
    A = de.anisotropy(d["C11"], d["C12"], d["C44"])
    sA = sigma_A(d["C11"], d["C12"], d["C44"], d["sC11"], d["sC12"], d["sC44"])

    ax.plot(mc["P"], A_mc, **dc.MC_STYLE, label=f"{dc.MC_LABEL} (extrap.)")
    ax.plot(fs["P"], A_fs, color=dc.COL["K"], linewidth=dc.LW_FIT,
            label=dc.FS_LABEL)
    dc.draw_data(ax, d["P"], A, sA, color=dc.COL["C44"], label=dc.DATA_LABEL)

    dc.set_labels(ax, "Anisotropy $A$")
    ax.set_ylim(1.15, 1.35)
    dc.legend(ax, loc="upper left")

    dc.print_table("plotted A", {
        "P": d["P"], "A": A, "sigma A": sA,
        "A FS": np.interp(d["P"], fs["P"], A_fs),
        "A McSk": np.interp(d["P"], mc["P"], A_mc)})
    dc.save(fig, "anisotropy")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
