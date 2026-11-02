"""
diamond_eos.py
--------------
Equations and parameters for the diamond single-crystal elasticity figures
(Gulick & Li). 3rd-order Eulerian (Birch) finite-strain ("FS 3O") model, with C12
from the bulk-modulus route (C12 = (3K - C11)/2) using K from Hu et al. 2021.

Everything a plotting script needs: parameter sets, the finite-strain equations,
aggregate (VRH) moduli, anisotropy, engineering moduli, velocities, a fine-grid
FS-line builder, a loader for Diamond_use_cleaned.xlsx, and the embedded
ground-truth table with a self-test.

Sign convention for the Eulerian strain follows the project workbook:
    f = 0.5 * (1 - (rho/rho0)**(2/3))        # f <= 0 under compression
Units: C,K,G,E in GPa; P in GPa; rho in g/cm^3; velocities in km/s.
"""
from __future__ import annotations
from dataclasses import dataclass
import os
import numpy as np

# --------------------------------------------------------------------------- #
# Parameter sets
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class ElasticParams:
    name: str
    c11_0: float; c11p: float
    c44_0: float; c44p: float
    c12_0: float; c12p: float
    k0: float;    k0p: float
    rho0: float = 3.512
    g0: float | None = None
    g0p: float | None = None


# FS 3O fit, anchored to Hu's K (C12_0, C12p are DERIVED from K and C11):
#   C12_0 = (3*K0  - C11_0)/2 = 121.2
#   C12p  = (3*K0p - C11p )/2 = 3.833
FITTED = ElasticParams(
    name="FS 3O (Hu-anchored)",
    c11_0=1075.2, c11p=3.135,
    c44_0=576.5,  c44p=1.899,
    c12_0=121.2,  c12p=3.833,
    k0=439.2,     k0p=3.6,      # Hu et al. 2021 (fixed)
    g0=533.0,     g0p=2.3,      # Hu et al. 2021 aggregate shear modulus
)

# McSkimin & Andreatch 1972 (ultrasonic, to 0.15 GPa) -- benchmark / extrapolation
MCSKIMIN = ElasticParams(
    name="McSkimin 1972",
    c11_0=1079.0, c11p=5.98,
    c44_0=578.0,  c44p=2.98,
    c12_0=124.0,  c12p=3.06,
    k0=(1079 + 2*124)/3,          # 442.3
    k0p=(5.98 + 2*3.06)/3,        # 4.03
)

# convenience: Hu aggregate EoS as simple linear laws
def hu_K(P):  return 439.2 + 3.6 * np.asarray(P, float)
def hu_G(P):  return 533.0 + 2.3 * np.asarray(P, float)

# --------------------------------------------------------------------------- #
# Finite-strain core
# --------------------------------------------------------------------------- #
def eulerian_strain(rho, rho0=3.512):
    """f = 0.5*(1 - (rho/rho0)**(2/3));  f <= 0 under compression."""
    return 0.5 * (1.0 - (np.asarray(rho, float) / rho0) ** (2.0 / 3.0))

def density_from_f(f, rho0=3.512):
    """Inverse of eulerian_strain: rho = rho0*(1-2f)**1.5."""
    return rho0 * (1.0 - 2.0 * np.asarray(f, float)) ** 1.5

def bm3_pressure(f, k0, k0p):
    """3rd-order Birch-Murnaghan pressure, project sign convention (f<=0)."""
    f = np.asarray(f, float)
    return -3.0 * k0 * (1 - 2*f) ** 2.5 * (1 + 1.5 * (4 - k0p) * f) * f

def cij_of_f(f, c0, cp, n, k0, k0p):
    """
    FS 3O effective elastic constant on a strain grid.
      n = 3 for the longitudinal C11, n = 1 for C44 and C12.
      ca = 3*k0*(cp - n) - 7*c0
      Cij(f) = (1-2f)**3.5 * (c0 - ca*f) + n * P(f)
    The +nP term uses the EoS derivative k0p (not the per-constant cp).
    """
    f = np.asarray(f, float)
    ca = 3.0 * k0 * (cp - n) - 7.0 * c0
    return (1 - 2*f) ** 3.5 * (c0 - ca * f) + n * bm3_pressure(f, k0, k0p)

def fs_grid(params: ElasticParams, fmax=-0.0095, npts=400):
    """
    Build smooth FS curves vs pressure for a parameter set.
    Returns a dict with P and the three Cij (and rho) along the grid.
    """
    f = np.linspace(0.0, fmax, npts)
    P = bm3_pressure(f, params.k0, params.k0p)
    rho = density_from_f(f, params.rho0)
    C11 = cij_of_f(f, params.c11_0, params.c11p, 3, params.k0, params.k0p)
    C44 = cij_of_f(f, params.c44_0, params.c44p, 1, params.k0, params.k0p)
    C12 = cij_of_f(f, params.c12_0, params.c12p, 1, params.k0, params.k0p)
    return dict(f=f, P=P, rho=rho, C11=C11, C44=C44, C12=C12)

# --------------------------------------------------------------------------- #
# Derived properties
# --------------------------------------------------------------------------- #
def vrh_moduli(C11, C12, C44):
    """Voigt-Reuss-Hill aggregate G and the (exact, cubic) bulk modulus K."""
    C11, C12, C44 = map(np.asarray, (C11, C12, C44))
    Gv = (C11 - C12 + 3*C44) / 5.0
    Gr = 5.0 * (C11 - C12) * C44 / (4.0 * C44 + 3.0 * (C11 - C12))
    Gvrh = 0.5 * (Gv + Gr)
    K = (C11 + 2*C12) / 3.0           # identically Hu K here (C12 built from it)
    return dict(Gv=Gv, Gr=Gr, Gvrh=Gvrh, K=K)

def anisotropy(C11, C12, C44):
    """Zener anisotropy A = 2*C44/(C11-C12);  A=1 is isotropic."""
    return 2.0 * np.asarray(C44) / (np.asarray(C11) - np.asarray(C12))

def poisson(K, G):   return (3*np.asarray(K) - 2*np.asarray(G)) / (2*(3*np.asarray(K) + np.asarray(G)))
def youngs(K, G):    return 9*np.asarray(K)*np.asarray(G) / (3*np.asarray(K) + np.asarray(G))
def pugh_GK(K, G):   return np.asarray(G) / np.asarray(K)           # G/K  (brittle if > ~0.57)
def pugh_BG(K, G):   return np.asarray(K) / np.asarray(G)           # B/G  (brittle if < 1.75)

def velocities(C11, C12, C44, rho):
    """Measured-mode velocities (km/s) for C in GPa, rho in g/cm^3."""
    C11, C12, C44, rho = map(np.asarray, (C11, C12, C44, rho))
    Vp100 = np.sqrt(C11 / rho)
    Vs100 = np.sqrt(C44 / rho)
    Vp110 = np.sqrt((C11 + C12 + 2*C44) / (2*rho))
    return dict(Vp100=Vp100, Vs100=Vs100, Vp110=Vp110)

# --------------------------------------------------------------------------- #
# Data loading
# --------------------------------------------------------------------------- #
# Embedded ground truth (per measured point). Columns documented in dtype.names.
GROUND_TRUTH = np.array(
    [(0.00, 1.0000, 3.512, 17.500, 12.842, 18.310, 1075.5, 1.9, 579.2, 1.5, 121.1, 2.3),
     (3.86, 0.9914, 3.543, 17.519, 12.814, 18.354, 1087.2, 1.9, 581.7, 1.5, 136.0, 2.3),
     (5.70, 0.9874, 3.557, 17.525, 12.829, 18.393, 1092.4, 1.9, 585.4, 1.5, 143.4, 2.3),
     (7.17, 0.9843, 3.568, 17.536, 12.848, 18.431, 1097.3, 1.9, 589.0, 1.5, 148.9, 2.3),
     (8.55, 0.9814, 3.579, 17.534, 12.851, 18.453, 1100.3, 1.9, 591.0, 1.5, 154.8, 2.3),
     (9.70, 0.9790, 3.587, 17.535, 12.857, 18.474, 1103.0, 1.9, 593.0, 1.5, 159.6, 2.3),
     (10.80, 0.9767, 3.596, 17.537, 12.864, 18.495, 1105.8, 1.9, 595.0, 1.5, 164.2, 2.3),
     (11.67, 0.9749, 3.602, 17.556, 12.889, 18.530, 1110.3, 2.0, 598.4, 1.5, 166.7, 2.3),
     (12.46, 0.9733, 3.608, 17.562, 12.898, 18.550, 1112.9, 2.0, 600.3, 1.6, 169.6, 2.3)],
    dtype=[("P", "f8"), ("VV0", "f8"), ("rho", "f8"),
           ("Vp100", "f8"), ("Vs100", "f8"), ("Vp110", "f8"),
           ("C11", "f8"), ("sC11", "f8"), ("C44", "f8"), ("sC44", "f8"),
           ("C12", "f8"), ("sC12", "f8")],
)

WORKBOOK = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "Diamond_use_cleaned.xlsx")

def load_workbook(path=WORKBOOK):
    """
    Pull the per-point data live from the workbook's 'C12 analysis' sheet
    (rows 2-10). Falls back to GROUND_TRUTH if openpyxl or the file is missing.
    Column map: A=P B=tP100 C=tS100 D=tP110 (two-way, us) J=rho K=Vp110
                L=Vp100 M=Vs100 N=C11 O=C44 R=sC11 S=sC44 X=C12 Y=sC12
    Also returns the 1-sigma travel-time resolutions from 'Inputs'!B14:B16
    as sig_tP100 / sig_tS100 / sig_tP110 (scalars).
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["C12 analysis"]
        g = lambda c, r: ws[f"{c}{r}"].value
        rows = range(2, 11)
        d = {k: np.array([g(c, r) for r in rows], float) for k, c in
             dict(P="A", tP100="B", tS100="C", tP110="D",
                  rho="J", Vp110="K", Vp100="L", Vs100="M",
                  C11="N", C44="O", sC11="R", sC44="S", C12="X", sC12="Y").items()}
        d["VV0"] = FITTED.rho0 / d["rho"]
        inp = wb["Inputs"]
        for k, cell in dict(sig_tP100="B14", sig_tS100="B15", sig_tP110="B16").items():
            d[k] = float(inp[cell].value)
        return d
    except Exception as e:  # noqa
        print(f"[load_workbook] falling back to embedded table ({e})")
        gt = GROUND_TRUTH
        return {n: gt[n].copy() for n in gt.dtype.names}

# --------------------------------------------------------------------------- #
# Self-test
# --------------------------------------------------------------------------- #
def _selftest():
    gt = GROUND_TRUTH
    # FS curves sampled at the data pressures (via each point's measured density)
    f = eulerian_strain(gt["rho"], FITTED.rho0)
    C11 = cij_of_f(f, FITTED.c11_0, FITTED.c11p, 3, FITTED.k0, FITTED.k0p)
    C44 = cij_of_f(f, FITTED.c44_0, FITTED.c44p, 1, FITTED.k0, FITTED.k0p)
    C12 = cij_of_f(f, FITTED.c12_0, FITTED.c12p, 1, FITTED.k0, FITTED.k0p)
    v = velocities(gt["C11"], gt["C12"], gt["C44"], gt["rho"])
    ok = True
    for lbl, got, ref, tol in [
        ("C11 fit", C11, gt["C11"], 2.0), ("C44 fit", C44, gt["C44"], 3.0),
        ("Vp100", v["Vp100"], gt["Vp100"], 0.01),
        ("Vs100", v["Vs100"], gt["Vs100"], 0.01),
        ("Vp110", v["Vp110"], gt["Vp110"], 0.01)]:
        err = np.max(np.abs(np.asarray(got) - ref))
        flag = "OK " if err <= tol else "!! "
        ok &= err <= tol
        print(f"  {flag}{lbl:8s} max|diff| = {err:.3f} (tol {tol})")
    A = anisotropy(gt["C11"], gt["C12"], gt["C44"])
    agg = vrh_moduli(gt["C11"], gt["C12"], gt["C44"])
    K, G = hu_K(gt["P"]), hu_G(gt["P"])
    print(f"  A(0)={A[0]:.3f} A(max)={A[-1]:.3f} | nu(0)={poisson(K,G)[0]:.4f} "
          f"E(0)={youngs(K,G)[0]:.0f} GPa | G/K(0)={pugh_GK(K,G)[0]:.3f} "
          f"B/G(0)={pugh_BG(K,G)[0]:.3f} | Gvrh(0)={agg['Gvrh'][0]:.1f}")
    print("  SELFTEST", "PASS" if ok else "FAIL")
    return ok

if __name__ == "__main__":
    _selftest()
