"""
Fig 5 - Isotropic engineering moduli vs pressure, three stacked panels.

    (a) Poisson ratio  nu = (3K - 2G) / (2(3K + G))
    (b) Young's modulus E = 9KG / (3K + G)
    (c) Pugh ratio G/K   (right axis: the usual B/G = K/G; the brittle/ductile
        threshold B/G = 1.75 is far off-scale - diamond is strongly brittle)

Each panel shows both choices of aggregate G (the open question in the spec):
    black solid   : Hu et al. 2021 K and G
    purple dashed : Hu K with G_VRH from the FS 3O fitted Cij
    purple points : Hu K with G_VRH from the measured Cij
K is Hu's in all three: K from the Cij is identical by construction.
"""
import numpy as np
import matplotlib.pyplot as plt

import diamond_eos as de
import dplot_common as dc

PANELS = [
    ("nu", r"Poisson ratio $\nu$", de.poisson),
    ("E", "$E$ (GPa)", de.youngs),
    ("GK", "$G/K$", de.pugh_GK),
]


def _recip(x):
    """G/K <-> B/G for the secondary axis (matplotlib probes x = 0)."""
    with np.errstate(divide="ignore"):
        return 1.0 / np.asarray(x, float)


def main(d=None, show=False):
    print("Fig 5 - engineering moduli")
    d = d if d is not None else dc.load()
    fig, axes = dc.new_fig(5.4, nrows=3, gridspec_kw={"hspace": 0.08})

    fs = de.fs_grid(de.FITTED)
    Pl = fs["P"]
    K_l, G_hu = de.hu_K(Pl), de.hu_G(Pl)
    G_fs = de.vrh_moduli(fs["C11"], fs["C12"], fs["C44"])["Gvrh"]
    K_d = de.hu_K(d["P"])
    G_d = de.vrh_moduli(d["C11"], d["C12"], d["C44"])["Gvrh"]

    table = {"P": d["P"]}
    for ax, tag, (key, ylab, fn) in zip(axes, "abc", PANELS):
        ax.plot(Pl, fn(K_l, G_hu), color=dc.COL["K"], linewidth=dc.LW_FIT,
                label=f"{dc.HU_LABEL} $K$, $G$")
        ax.plot(fs["P"], fn(de.hu_K(fs["P"]), G_fs), color=dc.COL["G"],
                linewidth=dc.LW_FIT, linestyle="--",
                label=f"$G_{{VRH}}$, {dc.FS_LABEL}")
        dc.draw_data(ax, d["P"], fn(K_d, G_d), color=dc.COL["G"],
                     label=f"$G_{{VRH}}$, {dc.DATA_LABEL}")
        ax.set_ylabel(ylab, fontsize=dc.FONT_PT["label"])
        ax.margins(y=0.15)
        dc.panel_label(ax, f"({tag})")
        table[f"{key} Hu"] = fn(de.hu_K(d["P"]), de.hu_G(d["P"]))
        table[f"{key} VRH"] = fn(K_d, G_d)

    sec = axes[2].secondary_yaxis("right", functions=(_recip, _recip))
    sec.set_ylabel("$B/G$", fontsize=dc.FONT_PT["label"])
    sec.tick_params(direction="in", labelsize=dc.FONT_PT["tick"])
    axes[2].tick_params(axis="y", which="both", right=False)

    axes[-1].set_xlabel("Pressure (GPa)", fontsize=dc.FONT_PT["label"])
    dc.legend(axes[0], loc="lower right")

    dc.print_table("plotted engineering moduli (Hu G vs measured G_VRH)", table)
    dc.save(fig, "engineering")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
