"""
Fig 4 - Aggregate bulk (K) and shear (G) moduli vs pressure, one panel.

    black solid   : K(P) = 439.2 + 3.6 P, Hu et al. 2021.  K from this work's
                    Cij is IDENTICAL by construction (C12 = (3K - C11)/2), so
                    it is not drawn separately.
    purple solid  : G(P) = 533 + 2.3 P, Hu et al. 2021
    purple dashed : G_VRH from the FS 3O fitted Cij
    purple points : G_VRH from the measured Cij, sigma propagated numerically
                    from sC11, sC12, sC44
    grey dashed   : McSkimin 1972 K and G_VRH, FS-extrapolated; grey points
                    mark their P = 0 values

The widening gap between Hu G and G_VRH is the under-stiffening of the [100]
data, worth showing rather than hiding.
"""
import numpy as np
import matplotlib.pyplot as plt

import diamond_eos as de
import dplot_common as dc


def sigma_Gvrh(C11, C12, C44, s11, s12, s44, h=1e-3):
    """Numerical first-order propagation (independent sigmas)."""
    g = lambda a, b, c: de.vrh_moduli(a, b, c)["Gvrh"]
    d11 = (g(C11 + h, C12, C44) - g(C11 - h, C12, C44)) / (2 * h)
    d12 = (g(C11, C12 + h, C44) - g(C11, C12 - h, C44)) / (2 * h)
    d44 = (g(C11, C12, C44 + h) - g(C11, C12, C44 - h)) / (2 * h)
    return np.sqrt((d11 * s11) ** 2 + (d12 * s12) ** 2 + (d44 * s44) ** 2)


def main(d=None, show=False):
    print("Fig 4 - K and G")
    d = d if d is not None else dc.load()
    fig, ax = dc.new_fig(2.8)

    fs = de.fs_grid(de.FITTED)
    mc = de.fs_grid(de.MCSKIMIN)
    G_fs = de.vrh_moduli(fs["C11"], fs["C12"], fs["C44"])["Gvrh"]
    mc_agg = de.vrh_moduli(mc["C11"], mc["C12"], mc["C44"])
    G = de.vrh_moduli(d["C11"], d["C12"], d["C44"])["Gvrh"]
    sG = sigma_Gvrh(d["C11"], d["C12"], d["C44"], d["sC11"], d["sC12"], d["sC44"])
    Pl = fs["P"]

    ax.plot(mc["P"], mc_agg["K"], **dc.MC_STYLE, label=f"{dc.MC_LABEL} (extrap.)")
    ax.plot(mc["P"], mc_agg["Gvrh"], **dc.MC_STYLE)
    ax.plot([0, 0], [mc_agg["K"][0], mc_agg["Gvrh"][0]], "o", color=dc.COL["Mc"],
            markersize=dc.MS, markeredgewidth=0, zorder=6)

    ax.plot(Pl, de.hu_K(Pl), color=dc.COL["K"], linewidth=dc.LW_FIT,
            label=f"$K$, {dc.HU_LABEL}")
    ax.plot(Pl, de.hu_G(Pl), color=dc.COL["G"], linewidth=dc.LW_FIT,
            label=f"$G$, {dc.HU_LABEL}")
    ax.plot(fs["P"], G_fs, color=dc.COL["G"], linewidth=dc.LW_FIT,
            linestyle="--", label=f"$G_{{VRH}}$, {dc.FS_LABEL}")
    dc.draw_data(ax, d["P"], G, sG, color=dc.COL["G"],
                 label=f"$G_{{VRH}}$, {dc.DATA_LABEL}")

    ax.text(9.0, float(de.hu_K(9.0)) - 5, "$K$", ha="center", va="top",
            fontsize=dc.FONT_PT["label"])
    ax.text(9.0, float(de.hu_G(9.0)) + 4, "$G$", ha="center", va="bottom",
            color=dc.COL["G"], fontsize=dc.FONT_PT["label"])

    dc.set_labels(ax, "Modulus (GPa)")
    ax.set_ylim(420, 570)
    dc.legend(ax, loc="center left", bbox_to_anchor=(0.0, 0.55))

    dc.print_table("plotted K, G", {
        "P": d["P"], "K Hu": de.hu_K(d["P"]), "G Hu": de.hu_G(d["P"]),
        "Gvrh": G, "sigma G": sG, "Gvrh FS": np.interp(d["P"], fs["P"], G_fs)})
    print(f"  McSkimin P=0: K = {mc_agg['K'][0]:.1f}, Gvrh = {mc_agg['Gvrh'][0]:.1f}")
    dc.save(fig, "moduli")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
