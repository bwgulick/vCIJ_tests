# Diamond Elasticity — Figure Generation Spec (for Claude Code)

**Project:** Single-crystal elasticity of diamond to ~12.5 GPa (Gulick & Li), ultrasonic
interferometry on [100] and [110] crystals, Stony Brook USCA-1000, Wang–Chen–Qi Al₂O₃
pressure scale. C12 is obtained from the bulk modulus route (C12 = (3K − C11)/2) using
K from Hu et al. 2021, and a 3rd-order Eulerian (Birch) finite-strain ("FS 3O") fit.

**Task:** reproduce the figure set from the vanadium paper, re-derived for diamond. Produce
one script per figure (or one module + a driver), saving vector PDF + 300-dpi PNG.

---

## 0. House style (match the vanadium paper)

If the vanadium paper or its plotting code is in the repo, **read it first and match it**:
fonts, figure size, color order, marker set, tick direction, legend placement, line weights.
If it is not available, use these defaults:

- `matplotlib`, constrained layout, figure ~3.4 in (single column) or 7 in (double).
- Font: a serif stack (e.g. `STIXGeneral` / Times) with `mathtext` for symbols; size 9–10.
- Ticks inward, on all four spines, minor ticks on.
- Pressure axis `P (GPa)` from 0 to 13 on every figure.
- Data = markers **with vertical (y) error bars, caps ~2 pt**; fits = solid lines;
  extrapolations/other sources = dashed.
- Consistent color map across all figures: **C11 = blue, C44 = green, C12 = orange/red,
  K = black, G = purple**. Keep McSkimin = grey dashed everywhere.
- Save as `fig_<name>.pdf` and `fig_<name>.png`.

---

## 1. Inputs

### 1a. Live data — `Diamond_use_cleaned.xlsx` (same folder)

Sheet **`C12 analysis`**, rows 2–10 (9 pressure points). Columns:

| col | quantity | col | quantity |
|----|----|----|----|
| A | P (GPa) | R | σ C11 |
| J | density ρ (g/cm³) | S | σ C44 |
| L | Vp[100] (km/s) | W | Hu K(P) |
| M | Vs[100] (km/s) | X | **C12** (K-route) |
| K | Vp[110] (km/s) | Y | σ C12 |
| N | C11 (GPa) | O | C44 (GPa) |

Sheet **`Finite strain fit`**: fitted parameters in `B4=C11₀, B5=C11′, B6=C44₀, B7=C44′,
B10=KS0, B11=K′S0, B14=C12₀, B15=C12′`.

Read with `openpyxl(..., data_only=True)`. The embedded tables below are the ground-truth
copy — if a read disagrees with them, stop and flag it.

### 1b. Fitted parameters (FS 3O, Hu-anchored) — the curves use these

```
rho0   = 3.512      # g/cm^3 (zero-P single-crystal density)
C11_0  = 1075.2 ;  C11p = 3.135
C44_0  = 576.5  ;  C44p = 1.899
C12_0  = 121.2  ;  C12p = 3.833     # derived: C12=(3K-C11)/2, C12'=(3K'-C11')/2
KS0    = 439.2  ;  K0p  = 3.6       # Hu et al. 2021 (FIXED, not fitted)
G0     = 533.0  ;  G0p  = 2.3       # Hu et al. 2021 aggregate shear modulus
```

### 1c. Reference values (for comparison / extrapolation lines)

```
# McSkimin & Andreatch 1972 (ultrasonic, 0.15 GPa) -- the benchmark
Mc_C11=1079 ; Mc_C11p=5.98
Mc_C12=124  ; Mc_C12p=3.06
Mc_C44=578  ; Mc_C44p=2.98
Mc_K0=(1079+2*124)/3 = 442.3 ; Mc_K0p=(5.98+2*3.06)/3 = 4.03
# Hu et al. 2021 bulk & shear EoS: K(P)=439.2+3.6P , G(P)=533+2.3P
```

### 1d. Embedded ground-truth table (per point)

```
P      V/V0    rho    Vp100   Vs100   Vp110    C11     sC11  C44    sC44  C12    sC12
0.00   1.0000  3.512  17.500  12.842  18.310   1075.5  1.9   579.2  1.5   121.1  2.3
3.86   0.9914  3.543  17.519  12.814  18.354   1087.2  1.9   581.7  1.5   136.0  2.3
5.70   0.9874  3.557  17.525  12.829  18.393   1092.4  1.9   585.4  1.5   143.4  2.3
7.17   0.9843  3.568  17.536  12.848  18.431   1097.3  1.9   589.0  1.5   148.9  2.3
8.55   0.9814  3.579  17.534  12.851  18.453   1100.3  1.9   591.0  1.5   154.8  2.3
9.70   0.9790  3.587  17.535  12.857  18.474   1103.0  1.9   593.0  1.5   159.6  2.3
10.80  0.9767  3.596  17.537  12.864  18.495   1105.8  1.9   595.0  1.5   164.2  2.3
11.67  0.9749  3.602  17.556  12.889  18.530   1110.3  2.0   598.4  1.5   166.7  2.3
12.46  0.9733  3.608  17.562  12.898  18.550   1112.9  2.0   600.3  1.6   169.6  2.3
```

---

## 2. Equations — put these in one `diamond_eos.py` module

**Eulerian strain** (this project's sign convention, f ≤ 0 under compression):
```
f(V) = 0.5 * (1 - (rho/rho0)**(2/3))          # = 0.5*(1 - (V0/V)**(2/3))
rho(f) = rho0 * (1 - 2f)**1.5                  # inverse, for building fine grids
```

**3rd-order Birch–Murnaghan pressure** (this project's form):
```
P(f; K0,K0p) = -3*K0*(1-2f)**2.5 * (1 + 1.5*(4-K0p)*f) * f
```

**FS 3O elastic constant** along a grid of f (n = 3 for C11, n = 1 for C44 and C12):
```
ca(c0,cp,n,K0) = 3*K0*(cp - n) - 7*c0
Cij(f) = (1-2f)**3.5 * (c0 - ca*f) + n * P(f; K0,K0p)
```
To draw an FS line vs pressure: make `f_grid = linspace(0, f_max, 400)` (f_max ≈ -0.009),
compute `P(f_grid)` and `Cij(f_grid)`, plot Cij vs P. Use `KS0, K0p` for this project's
lines; use `Mc_K0, Mc_K0p` with the McSkimin c0/c0′ for the McSkimin extrapolation lines.

**Aggregate moduli (VRH) from single-crystal Cij** — for G from your own data:
```
Gv  = (C11 - C12 + 3*C44)/5
Gr  = 5*(C11-C12)*C44 / (4*C44 + 3*(C11-C12))
Gvrh = 0.5*(Gv + Gr)
K   = (C11 + 2*C12)/3            # NOTE: identically = Hu K here, since C12 was built from it
```

**Zener anisotropy:** `A = 2*C44 / (C11 - C12)`  (A = 1 is elastically isotropic)

**Isotropic engineering moduli** (use aggregate K and G):
```
nu    = (3*K - 2*G) / (2*(3*K + G))     # Poisson ratio
E     = 9*K*G / (3*K + G)               # Young's modulus
pugh  = G/K                             # Pugh ratio (brittle if G/K > ~0.57, i.e. B/G < 1.75)
# also report B/G = K/G for the usual Pugh convention
```

**Velocities** (km/s, with ρ in g/cm³ and C in GPa, V=sqrt(C/ρ) gives km/s):
```
Vp100 = sqrt(C11/rho)
Vs100 = sqrt(C44/rho)
Vp110 = sqrt((C11 + C12 + 2*C44)/(2*rho))
```

---

## 3. Figures

### Fig 1 — V/V₀ compression + EoS
- x = P (GPa), y = V/V₀ (0.97–1.00).
- Markers: your V/V₀ data (col = embedded `V/V0`), small y error bars (use ±0.001 if no
  per-point σ available).
- Lines: BM3 EoS **(solid)** with Hu `KS0=439.2, K0p=3.6`; **(grey dashed)** McSkimin
  `442.3, 4.03`. Build the line by inverting `P(f)` on a fine f grid and plotting
  `1/(1-2f)**1.5` vs `P(f)` (that is V/V₀). Annotate K₀, K₀′ for each.

### Fig 2 — Cij "triplot" (one panel, all three)
- x = P, y = Cij (GPa), a single axes with C11, C44, C12 together (they are all well-behaved
  — do **not** make three separate panels).
- Each Cij: markers + **vertical error bars** (σ from the table), plus its **FS 3O line**
  (section 2). Color per house style; one legend entry per constant.
- Secondary y-axis is unnecessary; 0–1200 GPa holds all three. Label C12 near ~150, C44
  near ~590, C11 near ~1100.

### Fig 3 — Anisotropy
- x = P, y = A = 2C44/(C11−C12), range ~1.15–1.35.
- **Solid line**: FS 3O fit A(P) from your fitted Cij curves.
- **Grey dashed line**: McSkimin extrapolation — build Cij(P) from McSkimin c0/c0′ via the
  same FS form (with `Mc_K0, Mc_K0p`), then A(P). Mark that it is an extrapolation beyond
  0.15 GPa.
- Optional markers: pointwise A from the measured Cij (table). Horizontal reference line at
  A = 1 (isotropic).

### Fig 4 — K and G
- x = P, y = modulus (GPa), 420–570.
- **K(P)**: Hu `439.2 + 3.6P` (solid black). (Note in caption that K from your Cij is
  identical by construction, since C12 was derived from this K.)
- **G(P)**: plot **both** — Hu `533 + 2.3P` (solid purple) **and** G_VRH from your fitted
  Cij (purple dashed). They diverge upward with P (~536→545 for VRH vs 533→562 Hu); that gap
  is the known under-stiffening in the [100] data, worth showing.
- Add McSkimin K₀ (442.3) as a grey point at P = 0 for context.

### Fig 5 — Engineering moduli (Poisson, Young's, Pugh)
Three small stacked panels (shared x = P) or a 1×3 row:
- **ν(P)** ≈ 0.068 → 0.082 (diamond's famously low Poisson ratio).
- **E(P)** ≈ 1138 → 1215 GPa.
- **Pugh G/K** ≈ 1.21 → 1.16 (equivalently B/G ≈ 0.82; well below the 1.75 ductile
  threshold — diamond is strongly brittle). Add the 1.75 B/G line if plotting B/G.
All computed from aggregate Hu K and Hu G (use G_VRH variant as a dashed companion if useful).

### Fig 6 — Velocities (all three)
- x = P, y = velocity (km/s). Two y-ranges differ a lot, so either one panel with a broken
  axis or (cleaner) **two panels**: top = Vp[100] & Vp[110] (~17.5 and ~18.3–18.5),
  bottom = Vs[100] (~12.8–12.9).
- Each velocity: markers (table) + **FS 3O line** (compute Cij(f) and ρ(f) on the grid, then
  V(f) from section 2). Vp[100] and Vs[100] are nearly flat; Vp[110] rises more steeply
  (it carries C12) — expected.

---

## 4. Deliverables & checks
- `diamond_eos.py` (equations, parameters), one script per figure, a `run_all.py`.
- Print a small table of each figure's plotted values and diff against the embedded
  ground-truth table; abort if any Cij differs by > 0.5 GPa or any velocity by > 0.01 km/s.
- Caption each figure with the fit parameters used and the data provenance (Hu K, FS 3O).
- Keep McSkimin styling (grey dashed) identical across Figs 1, 3, 4 so it reads as one
  reference throughout.

**Open question for the user:** confirm the vanadium-paper figure dimensions/fonts, and
whether Fig 5 should use aggregate Hu G or VRH G from the Cij for the engineering moduli.
