"""
V/V0 vs pressure for vanadium from the 3rd-order Birch-Murnaghan EOS,
using the parameters of Ding et al. (2007):  K0 = 158 GPa, K' = 3.9.

The BM3 pressure as a function of compression x = V/V0 is

    P(x) = (3/2) K0 [x^(-7/3) - x^(-5/3)]
                 * { 1 + (3/4)(K' - 4)[x^(-2/3) - 1] }

We invert P(x) numerically at each target pressure to get V/V0(P), then plot.
"""
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import brentq

# --- Ding et al. 2007 vanadium EOS parameters ---
K0 = 158.0   # GPa   isothermal bulk modulus at P = 0
KP = 3.9     #        pressure derivative K' = dK/dP


def bm3_pressure(x, k0=K0, kp=KP):
    """3rd-order Birch-Murnaghan pressure (GPa) at compression x = V/V0."""
    f1 = x ** (-7.0 / 3.0) - x ** (-5.0 / 3.0)
    f2 = 1.0 + 0.75 * (kp - 4.0) * (x ** (-2.0 / 3.0) - 1.0)
    return 1.5 * k0 * f1 * f2


def vv0_at(P, k0=K0, kp=KP):
    """Invert BM3 for V/V0 at pressure P (GPa)."""
    if P <= 0:
        return 1.0
    # V/V0 is between a small compression and 1; bracket generously.
    return brentq(lambda x: bm3_pressure(x, k0, kp) - P, 0.3, 1.0)


def main(pmax=150.0, show=False):
    P = np.linspace(0.0, pmax, 400)
    vv0 = np.array([vv0_at(p) for p in P])

    fig, ax = plt.subplots(figsize=(7, 5))
    ax.plot(P, vv0, color="#0072B2", lw=2.0,
            label=f"BM3  $K_0$={K0:.0f} GPa, $K'$={KP:.1f}\n(Ding et al. 2007)")

    ax.set_xlabel("Pressure (GPa)", fontsize=12)
    ax.set_ylabel("$V/V_0$", fontsize=12)
    ax.tick_params(axis="both", which="both", direction="in",
                   top=True, right=True, labelsize=10)
    ax.set_xlim(0, pmax)
    ax.margins(y=0.04)
    ax.legend(fontsize=10, frameon=False)

    out = "vv0_ding2007.png"
    fig.savefig(out, dpi=300, bbox_inches="tight")
    print(f"saved -> {out}")

    # a small table of reference points
    print("\n P (GPa)   V/V0")
    for p in [0, 25, 50, 75, 100, 125, 150]:
        if p <= pmax:
            print(f" {p:6.0f}   {vv0_at(float(p)):.4f}")

    if show:
        plt.show()
    return fig


if __name__ == "__main__":
    main()
