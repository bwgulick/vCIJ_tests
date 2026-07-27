# Vanadium Cij — Global Minimization: build guide & audit

Companion to **`Vanadium_Cij_finite_strain_all_equations_GlobalMin.xlsx`**
(a copy of your `(1)` file with a new **`Global Min`** sheet added; your charts and all
Solver models are preserved; original file untouched).

---

## 1. How your workbook is actually wired (verified)

- All **8 combinations (4 FS + 4 CK) are live** across the 4 FS sheets. Each FS sheet holds
  BOTH the finite-strain fit block and the Cook fit block for its dropped-velocity case.
- `Summary` is a **static paste** of results (not live) — it cannot drive Solver.
- `Cook updated` holds only one combination + stray junk past column AL (`"tttt"`, `"o"`, `0`).
- Your **Solver models are saved** as hidden per-sheet defined names. Each FS sheet's saved
  Solver optimizes its **Cook block** (adjusts the AP/AQ/AR/AS cells; objective = CK "All sumsq").

### Source-cell map (what the Global Min sheet reads)
Rows = C11, C12, C44, C11', C12', C44'.

| combo | sheet | C11 | C12 | C44 | C11' | C12' | C44' | FS misfit | CK misfit |
|-------|-------|-----|-----|-----|------|------|------|-----------|-----------|
| NoVs111 FS | FS No Vs 111 | BG25 | BG26 | BG27 | BI25 | BI26 | BI27 | BK32 | — |
| NoVp111 FS | FS No Vp 111 | BC25 | BC26 | BC27 | BE25 | BE26 | BE27 | BH29 | — |
| NoVs100 FS | FS No Vs 100 | BF25 | BF26 | BF27 | BH25 | BH26 | BH27 | BK29 | — |
| NoVp100 FS | FS No Vp 100 | BH25 | BH26 | BH27 | BJ25 | BJ26 | BJ27 | BM30 | — |
| NoVs111 CK | FS No Vs 111 | AP26 | AP27 | AP28 | AR26 | AR27 | AR28 | — | AV30 |
| NoVp111 CK | FS No Vp 111 | AQ28 | AQ29 | AQ30 | AQ31 | AQ32 | AQ33 | — | AQ34 |
| NoVs100 CK | FS No Vs 100 | AQ26 | AQ27 | AQ28 | AQ29 | AQ30 | AQ31 | — | AQ32 |
| NoVp100 CK | FS No Vp 100 | AS26 | AS27 | AS28 | AS29 | AS30 | AS31 | — | AS32 |

---

## 2. Errors / issues found

1. **`#REF!` errors (176 cells) — all in vestigial columns, none feed the fits.**
   Verified: every residual column feeding the SUMSQ objectives is clean. Root cause: a Vs111
   velocity column was deleted and dependents left dangling. Safe to delete these columns:
   - `Cook updated`: AD, AH   · `FS No Vs 111`: AB   · `FS No Vp 111`: AA, AG
   - `FS No Vs 100`: Z, CH   · `FS No Vp 100`: Y
2. **`Cook updated` precision inconsistency**: K-integrand uses `0.6666` while J/L use
   `0.66666666666667` (2/3). Also `L1` label says "C46" (no such cubic constant; means C44).
   These K/L columns are NOT used in the length calc (only J is) — cosmetic.
3. **NOT a bug** (checked): `S` "L uncert" referencing rows +27 (e.g. `S3`→`O30`) is intentional —
   it pulls from the parallel sensitivity table in rows 29–50.

---

## 3. Deliverable A — consistency objective (built, non-destructive)

On the `Global Min` sheet:
- **B4:I9** = the 8 combos × 6 constants (live).
- **J4:J9** = weights (default C11/C12/C44 = 1, derivatives = 0). Change to taste.
- **D11** = normalize toggle. `0` = raw GPa² (C11 dominates); `1` = each divided by mean²
  so all constants contribute comparably (recommended when mixing C11 with C44').
- **Objective cells (point Solver at these, To: Min):**
  - `E13` WITHIN-FS spread (make the 4 FS agree)
  - `E14` WITHIN-CK spread (make the 4 Cook agree)
  - `E15` CK vs FS (between-method means)
  - `E16` ALL 8 total spread  (= within-FS + within-CK + between)

Current values (raw): WITHIN-FS ≈ 51.4, WITHIN-CK ≈ 42.0, CKvsFS ≈ 0.23, ALL ≈ 93.9 GPa².
The spread is driven almost entirely by the **NoVp100** combo's low C11 (235–236 vs ~240.8).

### Making Solver's variables actually shared (required for A)
Right now each FS sheet has its **own copy** of the input panel (rows 28–49), all identical.
Solver changing one sheet's cell won't touch the others. To calibrate globally, pick a master
sheet (say `FS No Vs 111`) and on the **other three** FS sheets replace the constant with a
reference, e.g. on `FS No Vp 111!B48` enter `='FS No Vs 111'!B48`. Do this for whichever inputs
you want Solver to move. Candidate shared inputs (row on every FS sheet):

`B29` (L0 111-P 0.623) · `C29` (L0 111-S 0.607) · `B38`/`D38` (L0 100 0.704) ·
`B31`/`B40` (density 6.09) · `B32`/`B41` (γ 1.23) · `B30`/`B39` (α) · `B45` (ΔC11 3) ·
`B46` (ΔC12/C44 1) · `B48` (K 151.14) · `B49` (K' 3.47).

> Recommendation: **do not** free all ~18 inputs — that's more knobs than data and will overfit.
> Start with the bench-length corrections (`B29`,`C29`,`B38`) and maybe `K`,`K'`; keep density,
> γ, T fixed at measured values. Add bounds/constraints in Solver to keep values physical.

### Workflow (outer loop, because the per-combo fits are nested)
1. Adjust shared inputs (by hand or via Solver on `E13`/`E14`/`E16`).
2. Re-run each per-combo Solver (they re-fit that combo's Cij to its data).
3. Re-read `E13`–`E16`. Repeat until the objective stops dropping.

---

## 4. Deliverable B — joint shared-parameter fit (rigorous)

Instead of 4 independent fits, fit **one** Cij set to all 4 datasets at once, so they agree by
construction. The live joint objective is already on the sheet:
- `B24` = Σ of the 4 FS data-misfit totals · `C24` = Σ of the 4 CK totals · `B25` = grand total.

To turn these into a true joint fit:
1. Put one shared 6-cell block on `Global Min` (C11,C12,C44,C11',C12',C44').
2. On each FS sheet, repoint its FS param cells (the BC25/BE25… group per the table above) to
   that shared block, e.g. `FS No Vp 111!BC25` → `='Global Min'!$<C11cell>`.
3. Solver: **Set Objective `B24`, To Min, By Changing** the 6 shared cells. (Do the same with
   the CK param cells → `C24`; both together → `B25`.)

Note: a joint fit's natural variables are the **Cij themselves** (not the physical inputs) —
that's what "fit one curve to all data" means. Use A for physical-input self-consistency and B
for a single best-fit Cij set; they answer different questions.

---

## 5. Solver quick-reference

Data ▸ Solver ▸ GRG Nonlinear ▸ Set Objective = (E13 / E14 / E15 / E16 / B24 / C24 / B25) ▸
To: Min ▸ By Changing Variable Cells = (your shared inputs, or the shared Cij for B) ▸
add bounds as constraints ▸ Solve. Un-tick "Make Unconstrained Variables Non-Negative" if any
variable can be negative.
