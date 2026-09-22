# Step 3: combined-final uncertainty rollup on Summary / GLOBAL Summary / Super Summary.
# Emits a TSV (sheet<TAB>cell<TAB>F<TAB>formula) applied via _apply_spec.ps1.
# All blocks placed in confirmed-empty far-right columns (no collision with existing content).
#
# Per averaged constant the block reports:
#   value       = the sheet's existing all-combo mean (M col)
#   sig_fit     = sqrt(SUM of the 8 combo fit-SE^2)/8   [fit-error of the mean]
#   sig_scatter = STDEV(8 combo values)/sqrt(8)         [between-combo scatter of the mean]
#   sig_total   = sqrt(sig_fit^2 + sig_scatter^2)        [combined 1-sigma +/-]
# SolverAid SE cells (written at anchor col GR=200): value col GS(201), SE col GT(202).
#   params: CK SE = GT44..49, FS SE = GT70..75 (idx 0..5 = C11,C12,C44,C11',C12',C44')
#   moduli sigma: CK = GS54..57, FS = GS80..83 (K0,GV0,GR0,GH0)

import os

_HERE = os.path.dirname(os.path.abspath(__file__))

rows = []
def emit(sheet, cell, formula): rows.append(f"{sheet}\t{cell}\tF\t{formula}")

PARAMS = ["C11","C12","C44","C11'","C12'","C44'"]
MODS   = ["K0","GV0","GR0","GH0"]
COMBO  = list("BCDEFGHI")          # source combo value columns (4 FS + 4 CK)
HCOLS  = ["AB","AC","AD","AE","AF","AG","AH","AI"]   # helper per-combo modulus columns

def sumsq8(sheets, gt_fs, gt_ck):
    refs = [f"'{s}'!{gt_fs}" for s in sheets] + [f"'{s}'!{gt_ck}" for s in sheets]
    return "SUMSQ(" + ",".join(refs) + ")"

def build(sheet, sheets, crow0, Mcol,
          qc, vc, fc, sc, tc, r0,          # rollup block cols + header row
          hlabel, hr0):                     # helper: label col + first row (4 rows)
    """crow0 = row of C11 in the sheet's constant table. Rollup title at r0-1."""
    n = len(COMBO)                          # 8
    sq = f"SQRT({n})"
    crC11, crC12, crC44 = crow0, crow0+1, crow0+2

    emit(sheet, f"{qc}{r0-1}", "UNCERTAINTY ROLLUP  (1-sigma, GPa)")
    emit(sheet, f"{qc}{r0}", "Quantity"); emit(sheet, f"{vc}{r0}", "value")
    emit(sheet, f"{fc}{r0}", "sig_fit");  emit(sheet, f"{sc}{r0}", "sig_scatter")
    emit(sheet, f"{tc}{r0}", "sig_total(+/-)")

    # 6 fitted parameters
    for idx, name in enumerate(PARAMS):
        cr, outr = crow0+idx, r0+1+idx
        emit(sheet, f"{qc}{outr}", name)
        emit(sheet, f"{vc}{outr}", f"={Mcol}{cr}")
        emit(sheet, f"{fc}{outr}", f"=SQRT({sumsq8(sheets, f'GT{70+idx}', f'GT{44+idx}')})/{n}")
        emit(sheet, f"{sc}{outr}", f"=STDEV({COMBO[0]}{cr}:{COMBO[-1]}{cr})/{sq}")
        emit(sheet, f"{tc}{outr}", f"=SQRT({fc}{outr}^2+{sc}{outr}^2)")

    # helper: per-combo moduli (own columns HCOLS), one row per modulus
    hK, hGV, hGR, hGH = hr0, hr0+1, hr0+2, hr0+3
    emit(sheet, f"{hlabel}{hK}",  "K0 per combo")
    emit(sheet, f"{hlabel}{hGV}", "GV0 per combo")
    emit(sheet, f"{hlabel}{hGR}", "GR0 per combo")
    emit(sheet, f"{hlabel}{hGH}", "GH0 per combo")
    for j, hcol in enumerate(HCOLS):
        src = COMBO[j]
        c11, c12, c44 = f"{src}{crC11}", f"{src}{crC12}", f"{src}{crC44}"
        D = f"({c11}-{c12})"
        emit(sheet, f"{hcol}{hK}",  f"=({c11}+2*{c12})/3")
        emit(sheet, f"{hcol}{hGV}", f"=({c11}-{c12}+3*{c44})/5")
        emit(sheet, f"{hcol}{hGR}", f"=5*{D}*{c44}/(3*{D}+4*{c44})")
        emit(sheet, f"{hcol}{hGH}", f"=({hcol}{hGV}+{hcol}{hGR})/2")

    # 4 moduli rollup
    mod_hr = {"K0":hK,"GV0":hGV,"GR0":hGR,"GH0":hGH}
    mod_gs = {"K0":(80,54),"GV0":(81,55),"GR0":(82,56),"GH0":(83,57)}
    M11, M12, M44 = f"{Mcol}{crC11}", f"{Mcol}{crC12}", f"{Mcol}{crC44}"
    DM = f"({M11}-{M12})"
    mod_val = {"K0":f"=({M11}+2*{M12})/3", "GV0":f"=({M11}-{M12}+3*{M44})/5",
               "GR0":f"=5*{DM}*{M44}/(3*{DM}+4*{M44})"}
    base = r0+1+len(PARAMS)
    outrow = {m: base+i for i,m in enumerate(MODS)}
    for m in MODS:
        outr, hr = outrow[m], mod_hr[m]
        gs_fs, gs_ck = mod_gs[m]
        emit(sheet, f"{qc}{outr}", m)
        if m == "GH0":
            emit(sheet, f"{vc}{outr}", f"=({vc}{outrow['GV0']}+{vc}{outrow['GR0']})/2")
        else:
            emit(sheet, f"{vc}{outr}", mod_val[m])
        emit(sheet, f"{fc}{outr}", f"=SQRT({sumsq8(sheets, f'GS{gs_fs}', f'GS{gs_ck}')})/{n}")
        emit(sheet, f"{sc}{outr}", f"=STDEV({HCOLS[0]}{hr}:{HCOLS[-1]}{hr})/{sq}")
        emit(sheet, f"{tc}{outr}", f"=SQRT({fc}{outr}^2+{sc}{outr}^2)")

    note_r = outrow["GH0"] + 2
    emit(sheet, f"{qc}{note_r}",
         "NOTE: sig_fit=sqrt(sum SE^2)/8 (8 combos: 4 FS+4 CK); sig_scatter=STDEV/sqrt(8); "
         "value=all-combo mean. FS&CK on a sheet share data (correlated) so scatter is the "
         "robust term. Re-run _solveraid.ps1 after any re-fit. Helper table to the right.")
    return outrow

# ---- Summary (4 regular combo sheets); rollup U..Y r3+, helper label AA data AB..AI r5+ ----
REG = ["No Vs 111","No Vp 111","No Vs 100","No Vp 100"]
build("Summary", REG, 4, "M", "U","V","W","X","Y", 4, "AA", 5)

# ---- GLOBAL Summary (4 global combo sheets); rollup S..W r3+, helper AA/AB..AI r5+ ----
GLB = ["Global No Vs 111","Global No Vp 111","Global No Vs 100","Global No Vp 100"]
build("GLOBAL Summary", GLB, 2, "M", "S","T","U","V","W", 3, "AA", 5)

# ---- Super Summary: reference the two rollups (single source of truth); cols X..AB ----
# Summary rollup:  value V, total Y, rows 5..14   |  GLOBAL: value T, total W, rows 4..13
QALL = PARAMS + MODS
emit("Super Summary", "X2", "COMBINED FINAL (1-sigma, GPa)  <- from Summary & GLOBAL Summary rollups")
emit("Super Summary", "X3", "Quantity"); emit("Super Summary", "Y3", "Reg value")
emit("Super Summary", "Z3", "Reg +/-"); emit("Super Summary", "AA3", "Global value")
emit("Super Summary", "AB3", "Global +/-")
for i, q in enumerate(QALL):
    r, sr, gr = 4+i, 5+i, 4+i
    emit("Super Summary", f"X{r}", q)
    emit("Super Summary", f"Y{r}", f"='Summary'!V{sr}")
    emit("Super Summary", f"Z{r}", f"='Summary'!Y{sr}")
    emit("Super Summary", f"AA{r}", f"='GLOBAL Summary'!T{gr}")
    emit("Super Summary", f"AB{r}", f"='GLOBAL Summary'!W{gr}")
emit("Super Summary", "X15",
     "NOTE: Reg = mean of 8 per-pressure combo fits (4 FS+4 CK); Global = mean of 8 global fits. "
     "+/- combines fit SE and between-combo scatter. See Summary!U3 and 'GLOBAL Summary'!S2.")

with open(os.path.join(_HERE, "_uncert_step3_spec.tsv"),"w",encoding="utf-8") as f:
    f.write("\n".join(rows))
print(f"wrote {len(rows)} cell ops for step3")
