"""
V/V0 (compression) vs pressure.

Overlays two series:
    vv0   the measured V/V0 points (this study), read from the "All" sheet
          ("P EXP" / "V/V0"), with:
            * vertical error bars = the length-propagated sigma(V/V0) taken
              from the "uncertainty calcs" sheet ("v/v0 ncrt" column, which the
              workbook builds as 3 L^2 sigma_L / V0 from the length uncertainty
              sigma_L ~ 1 um).  Gated by the GUI's vertical-error toggle.
            * horizontal error bars = the shared "ncrt P".  Gated by the GUI's
              horizontal-error toggle.
    ding  a smooth 3rd-order Birch-Murnaghan reference curve for vanadium using
          the Ding et al. (2007) parameters K0 = 158 GPa, K' = 3.9, drawn over
          the pressure span of the data.  Drawn as a line (MARK["ding"]="-").

Both series' colours / markers / legend names / visibility are live-editable in
the GUI (series keys "vv0" and "ding"), exactly like CK / FS / poly.
"""
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

import vplot_common as vc

# Ding et al. (2007) vanadium BM3 EOS parameters (edit here to refit the curve).
K0_DING = 158.0   # GPa   bulk modulus at P = 0
KP_DING = 3.9     #        pressure derivative K'

# Panel registry so the GUI's bounds/label panel shows a V/V0 row.
PANELS = [("V/V0", "$V/V_0$")]


def _bm3_pressure(x, k0, kp):
    """3rd-order Birch-Murnaghan pressure (GPa) at compression x = V/V0."""
    f1 = x ** (-7.0 / 3.0) - x ** (-5.0 / 3.0)
    f2 = 1.0 + 0.75 * (kp - 4.0) * (x ** (-2.0 / 3.0) - 1.0)
    return 1.5 * k0 * f1 * f2


def _ding_curve(pmax, k0=K0_DING, kp=KP_DING, n=400):
    """(P, V/V0) sampled from the BM3 EOS over 0..pmax GPa."""
    from scipy.optimize import brentq
    P = np.linspace(0.0, pmax, n)
    vv0 = np.array([
        1.0 if p <= 0 else brentq(lambda x: _bm3_pressure(x, k0, kp) - p, 0.3, 1.0)
        for p in P
    ])
    return P, vv0


def main(path=None, show=False):
    s = vc.set_scale((7, 5))
    fig, ax = plt.subplots(figsize=vc.figsize_in((7, 5)))

    # --- measured V/V0 points (this study) ---
    P, vv0, sig, sigP = vc.load_vv0(path)
    if vc.visible("vv0") and P.size:
        vc.draw_pts(ax, P, vv0, yerr=sig, xerr=sigP,
                    color=vc.COL["vv0"], marker=vc.MARK["vv0"],
                    label=vc.LABEL["vv0"])

    # --- Ding et al. 2007 BM3 reference curve, over the data's pressure span ---
    if vc.visible("ding"):
        pmax = float(np.nanmax(P)) if P.size else 150.0
        cx, cy = _ding_curve(pmax if pmax > 0 else 150.0)
        lab = f"{vc.LABEL['ding']} ($K_0$={K0_DING:.0f}, $K'$={KP_DING:.1f})"
        vc.draw_pts(ax, cx, cy, color=vc.COL["ding"],
                    marker=vc.MARK["ding"], label=lab)

    ax.set_ylabel(vc.ylabel_for("V/V0", "$V/V_0$"), fontsize=12 * s)
    ax.set_xlabel("Pressure (GPa)", fontsize=12 * s)
    ax.tick_params(axis="both", labelsize=10 * s)
    ax.tick_params(axis="both", which="both", direction="in",
                   top=True, right=True)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.margins(y=0.08)

    vc.apply_ylim(ax, "V/V0")
    vc.apply_xlim(ax)

    handles, labels = ax.get_legend_handles_labels()
    if handles:
        vc.place_legend(ax, handles, labels, 10 * s)

    out = vc.output_path("vv0.png")
    fig.savefig(out, dpi=vc.DPI, bbox_inches="tight")
    print(f"V/V0 -> {out}")
    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
