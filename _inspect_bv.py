import openpyxl
wf = openpyxl.load_workbook("Cij analysis/Vanadium_Cij_Brian_Claude.xlsx", data_only=False)
wv = openpyxl.load_workbook("Cij analysis/Vanadium_Cij_Brian_Claude.xlsx", data_only=True)
f = wf["No Vp 100"]; v = wv["No Vp 100"]

print("=== BV/BW/BX headers + r2 formula + values, vs AL/AJ/AK ===")
for col,label in [("BV","C11?"),("BW","C44?"),("BX","C12?")]:
    print(col, label, "hdr:", repr(f[col+"1"].value), "| r2:", repr(f[col+"2"].value), "| val r2:", v[col+"2"].value)
print("---")
for col,label in [("AL","C11"),("AJ","C44"),("AK","C12")]:
    print(col, label, "val r2:", v[col+"2"].value)

print()
print("=== value check row2: CB(K), CG(GV), CH(GR), CI(GH) and existing sig FL..FO ===")
for col in ["CB","CG","CH","CI","FL","FM","FN","FO"]:
    print(col, f[col+"1"].value, "=", v[col+"2"].value)

print()
print("=== compare BV vs AL, BW vs AJ, BX vs AK across rows 2..23 ===")
maxdiff = {"C11":0,"C44":0,"C12":0}
for r in range(2,24):
    for a,b,k in [("BV","AL","C11"),("BW","AJ","C44"),("BX","AK","C12")]:
        va, vb = v[f"{a}{r}"].value, v[f"{b}{r}"].value
        if isinstance(va,(int,float)) and isinstance(vb,(int,float)):
            maxdiff[k]=max(maxdiff[k], abs(va-vb))
print("max |BV-AL|, |BW-AJ|, |BX-AK| over rows 2-23:", maxdiff)
