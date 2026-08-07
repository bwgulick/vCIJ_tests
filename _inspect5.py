import openpyxl
wv = openpyxl.load_workbook("Cij analysis/Vanadium_Cij_Brian_Claude.xlsx", data_only=True)
s = wv["No Vp 100"]
print("r | P | len | Vs100(AE) sZ | Vp111(AC) sAA | Vs111(AD) sAB | C11(AL) sFI | C44(AJ) sFK | C12(AK) sFJ | Vp100=2*.712/F")
lastrow=0
for r in range(2,26):
    P=s[f"A{r}"].value
    if P is None and s[f"M{r}"].value is None: continue
    lastrow=r
    F=s[f"F{r}"].value
    vp100 = 2*0.712/F if F else None
    def g(c): 
        v=s[f"{c}{r}"].value
        return f"{v:.3f}" if isinstance(v,(int,float)) else str(v)
    print(f"{r:>2} P={g('A'):<7} L={g('M'):<6} Vs100={g('AE')} s={g('Z')} | Vp111={g('AC')} s={g('AA')} | Vs111={g('AD')} s={g('AB')} | C11={g('AL')} s={g('FI')} | C44={g('AJ')} s={g('FK')} | C12={g('AK')} s={g('FJ')} | Vp100={vp100:.3f}")
print("last data row:", lastrow)
