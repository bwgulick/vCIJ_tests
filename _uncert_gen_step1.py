# Generates a TSV spec of per-point uncertainty formulas (Step 1) for the 4 combo sheets.
# Output columns: sheet<TAB>cell<TAB>kind(F/T)<TAB>content
import openpyxl
from openpyxl.utils import get_column_letter as CL

R0, R1 = 2, 23                      # data rows
START = 150                         # first new column (ET)
path = r"Cij analysis/Vanadium_Cij_Brian_Claude.xlsx"

# Per-sheet config: Cij experimental cells, and 3 measurement channels (m_cell, time_cell)
# with coefficient vectors a[C11],a[C44],a[C12] over the 3 channels.
sheets = {
 "No Vs 111": dict(C11="AG", C12="AI", C44="AH",
    ch=[("AG","F"),("AH","G"),("AF","H")],           # p100, s100, p111
    a11=[1,0,0], a44=[0,1,0], a12=[-0.5,-2,1.5],
    vs111_vel=None, vs111_time="I"),
 "No Vp 111": dict(C11="AH", C12="AJ", C44="AI",
    ch=[("AH","F"),("AI","G"),("AF","I")],           # p100, s100, s111
    a11=[1,0,0], a44=[0,1,0], a12=[1,1,-3],
    vs111_vel="AC", vs111_time="I"),
 "No Vs 100": dict(C11="AH", C12="AJ", C44="AI",
    ch=[("AH","F"),("AF","H"),("AG","I")],           # p100, p111, s111
    a11=[1,0,0], a44=[-0.5,0.5,1], a12=[0.5,0.5,-2],
    vs111_vel="AD", vs111_time="I"),
 "No Vp 100": dict(C11="AL", C12="AK", C44="AJ",
    ch=[("AJ","G"),("AH","H"),("AI","I")],           # s100, p111, s111
    a11=[-2,1,2], a44=[1,0,0], a12=[-1,1,-1],
    vs111_vel="AD", vs111_time="I"),
}

# column layout (offsets from START)
names = ["relA2","B1","B2","B3","VarC11","VarC12","VarC44","Cov1112","Cov1144","Cov1244",
         "GRg1","GRg3","GHg1","GHg2","GHg3",
         "sigC11","sigC12","sigC44","sigK","sigGV","sigGR","sigGH"]
COL = {n: CL(START+i) for i,n in enumerate(names)}

def num(x):
    # tidy coefficient formatting
    if x == int(x): return str(int(x))
    return repr(x)

rows = []
def emit(sheet, cell, kind, content): rows.append(f"{sheet}\t{cell}\t{kind}\t{content}")

for sheet, cfg in sheets.items():
    # headers
    emit(sheet, f"{COL['relA2']}1", "T", "UNCERTAINTY BLOCK (per-point, 1sigma) -->")
    for n in names[1:]:
        emit(sheet, f"{COL[n]}1", "T", n)
    C11,C12,C44 = cfg["C11"],cfg["C12"],cfg["C44"]
    ch = cfg["ch"]; a11,a44,a12 = cfg["a11"],cfg["a44"],cfg["a12"]
    for r in range(R0,R1+1):
        relA2 = f"{COL['relA2']}{r}"
        B = [f"{COL['B'+str(k+1)]}{r}" for k in range(3)]
        V11 = f"{COL['VarC11']}{r}"; V12=f"{COL['VarC12']}{r}"; V44=f"{COL['VarC44']}{r}"
        X1112=f"{COL['Cov1112']}{r}"; X1144=f"{COL['Cov1144']}{r}"; X1244=f"{COL['Cov1244']}{r}"
        GRg1=f"{COL['GRg1']}{r}"; GRg3=f"{COL['GRg3']}{r}"
        GHg1=f"{COL['GHg1']}{r}"; GHg2=f"{COL['GHg2']}{r}"; GHg3=f"{COL['GHg3']}{r}"
        c11=f"{C11}{r}"; c12=f"{C12}{r}"; c44=f"{C44}{r}"

        # relA2
        emit(sheet, relA2, "F", f"=(0.02/$B$40)^2+9*(0.001/$B$38)^2+({'Q'}{r}/{'M'}{r})^2")
        # channel bases B_x = 4*m_x^2*(0.0002/t_x)^2
        for k,(mc,tc) in enumerate(ch):
            emit(sheet, B[k], "F", f"=4*({mc}{r})^2*(0.0002/{tc}{r})^2")
        # variances
        def varf(cell, a):
            terms=[f"({cell})^2*{relA2}"]
            for k in range(3):
                if a[k]!=0: terms.append(f"{num(a[k]*a[k])}*{B[k]}")
            return "="+"+".join(terms)
        emit(sheet, V11, "F", varf(c11,a11))
        emit(sheet, V12, "F", varf(c12,a12))
        emit(sheet, V44, "F", varf(c44,a44))
        # covariances
        def covf(cA,cB,aA,aB):
            terms=[f"({cA})*({cB})*{relA2}"]
            for k in range(3):
                coef=aA[k]*aB[k]
                if coef!=0: terms.append(f"{num(coef)}*{B[k]}")
            return "="+"+".join(terms)
        emit(sheet, X1112, "F", covf(c11,c12,a11,a12))
        emit(sheet, X1144, "F", covf(c11,c44,a11,a44))
        emit(sheet, X1244, "F", covf(c12,c44,a12,a44))
        # GR gradient helpers: denom=3*(C11-C12)+4*C44
        denom=f"(3*(({c11})-({c12}))+4*({c44}))"
        emit(sheet, GRg1, "F", f"=20*({c44})^2/{denom}^2")
        emit(sheet, GRg3, "F", f"=15*(({c11})-({c12}))^2/{denom}^2")
        # GH gradient = average of GV grad (1/5,-1/5,3/5) and GR grad (GRg1,-GRg1,GRg3)
        emit(sheet, GHg1, "F", f"=(0.2+{GRg1})/2")
        emit(sheet, GHg2, "F", f"=(-0.2-{GRg1})/2")
        emit(sheet, GHg3, "F", f"=(0.6+{GRg3})/2")
        # sigmas for Cij
        emit(sheet, f"{COL['sigC11']}{r}", "F", f"=SQRT({V11})")
        emit(sheet, f"{COL['sigC12']}{r}", "F", f"=SQRT({V12})")
        emit(sheet, f"{COL['sigC44']}{r}", "F", f"=SQRT({V44})")
        # sigK: g=(1/3,2/3,0)
        emit(sheet, f"{COL['sigK']}{r}", "F",
             f"=SQRT({V11}/9+4*{V12}/9+(4/9)*{X1112})")
        # sigGV: g=(1/5,-1/5,3/5)
        emit(sheet, f"{COL['sigGV']}{r}", "F",
             f"=SQRT(({V11}+{V12}+9*{V44})/25+(-2*{X1112}+6*{X1144}-6*{X1244})/25)")
        # sigGR: quadratic form with (GRg1,-GRg1,GRg3)
        emit(sheet, f"{COL['sigGR']}{r}", "F",
             f"=SQRT({GRg1}^2*({V11}+{V12})-2*{GRg1}^2*{X1112}+{GRg3}^2*{V44}"
             f"+2*{GRg1}*{GRg3}*({X1144}-{X1244}))")
        # sigGH: full quadratic form with (GHg1,GHg2,GHg3)
        emit(sheet, f"{COL['sigGH']}{r}", "F",
             f"=SQRT({GHg1}^2*{V11}+{GHg2}^2*{V12}+{GHg3}^2*{V44}"
             f"+2*{GHg1}*{GHg2}*{X1112}+2*{GHg1}*{GHg3}*{X1144}+2*{GHg2}*{GHg3}*{X1244})")

    # Fix broken AB (VS111 ncrt): reference correct Vs111 velocity if measured, else clear
    for r in range(R0,R1+1):
        vv=cfg["vs111_vel"]; tt=cfg["vs111_time"]
        if vv is None:
            emit(sheet, f"AB{r}", "T", "")   # Vs111 omitted on this sheet -> clear #REF!
        else:
            emit(sheet, f"AB{r}", "F", f"=SQRT((Q{r}/M{r})^2+(0.0002/{tt}{r})^2)*{vv}{r}")

with open("_uncert_step1_spec.tsv","w",encoding="utf-8") as f:
    f.write("\n".join(rows))
print(f"wrote {len(rows)} cell ops for step1; new cols {COL['relA2']}..{COL['sigGH']}")
