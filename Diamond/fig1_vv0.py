"""
Fig 1 - V/V0 compression vs pressure with BM3 EoS lines.

    markers  : measured V/V0 = rho0/rho from the Cook-length density
               ('C12 analysis' col J), +/- SIG_VV0 (no per-point sigma exists)
    solid    : BM3 with Hu et al. 2021  K0 = 439.2, K0' = 3.6
    grey --  : BM3 with McSkimin 1972   K0 = 442.3, K0' = 4.03

Lines are built on a fine Eulerian-strain grid: V/V0 = (1-2f)^-1.5 vs P(f).
The points sitting above the Hu line is the known ~2-3% under-compression of
the Cook-length volumes (see the workbook README).
"""
import numpy as np
import matplotlib.pyplot as plt

import diamond_eos as de
import dplot_common as dc


def eos_line(k0, k0p, fmax=-0.0095, n=400):
    f = np.linspace(0.0, fmax, n)
    return de.bm3_pressure(f, k0, k0p), (1 - 2 * f) ** -1.5


def main(d=None, show=False):
    print("Fig 1 - V/V0")
    d = d if d is not None else dc.load()
    fig, ax = dc.new_fig(2.6)

    hu, mc = de.FITTED, de.MCSKIMIN
    P, vv = eos_line(hu.k0, hu.k0p)
    ax.plot(P, vv, color=dc.COL["K"], linewidth=dc.LW_FIT,
            label=f"{dc.HU_LABEL} ({hu.k0:.1f}, {hu.k0p:.1f})")
    P, vv = eos_line(mc.k0, mc.k0p)
    ax.plot(P, vv, **dc.MC_STYLE,
            label=f"McSkimin 1972 ({mc.k0:.1f}, {mc.k0p:.2f})")

    dc.draw_data(ax, d["P"], d["VV0"], np.full_like(d["P"], dc.SIG_VV0),
                 color=dc.COL["C11"], label=dc.DATA_LABEL)

    dc.set_labels(ax, "$V/V_0$")
    ax.set_ylim(0.970, 1.002)
    dc.legend(ax, loc="lower left", title="BM3 EoS ($K_0$, $K_0'$)",
              title_fontsize=dc.FONT_PT["legend"], alignment="left")

    dc.print_table("plotted V/V0 (measured) vs Hu BM3 at the same P", {
        "P": d["P"], "V/V0": d["VV0"],
        "Hu BM3": np.interp(d["P"], *eos_line(hu.k0, hu.k0p))})
    dc.save(fig, "vv0")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
