import zipfile, shutil, os, re
from xml.sax.saxutils import escape
import openpyxl

SRC = r"C:\Users\bgulick\Downloads\Vanadium_Cij_finite_strain_all_equations(1).xlsx"
OUT = r"C:\Users\bgulick\Downloads\Vanadium_Cij_finite_strain_all_equations_GlobalMin.xlsx"

# ---- live source-cell map: column -> (sheet, [C11,C12,C44,C11',C12',C44']) ----
cols = {
 'B': ("FS No Vs 111", ['BG25','BG26','BG27','BI25','BI26','BI27']),
 'C': ("FS No Vp 111", ['BC25','BC26','BC27','BE25','BE26','BE27']),
 'D': ("FS No Vs 100", ['BF25','BF26','BF27','BH25','BH26','BH27']),
 'E': ("FS No Vp 100", ['BH25','BH26','BH27','BJ25','BJ26','BJ27']),
 'F': ("FS No Vs 111", ['AP26','AP27','AP28','AR26','AR27','AR28']),
 'G': ("FS No Vp 111", ['AQ28','AQ29','AQ30','AQ31','AQ32','AQ33']),
 'H': ("FS No Vs 100", ['AQ26','AQ27','AQ28','AQ29','AQ30','AQ31']),
 'I': ("FS No Vp 100", ['AS26','AS27','AS28','AS29','AS30','AS31']),
}
qlabels = ['C11','C12','C44',"C11'","C12'","C44'"]
FS_cols = ['B','C','D','E']
CK_cols = ['F','G','H','I']

# sanity: compute current values + objective from cached workbook
wbv = openpyxl.load_workbook(SRC, data_only=True)
grid = {}  # (col,qi)->value
for col,(sh,cells) in cols.items():
    for qi,cell in enumerate(cells):
        grid[(col,qi)] = wbv[sh][cell].value
print("Current live values pulled:")
for qi,q in enumerate(qlabels):
    fs=[grid[(c,qi)] for c in FS_cols]; ck=[grid[(c,qi)] for c in CK_cols]
    print(f"  {q:5s} FS={['%.4f'%x for x in fs]}  CK={['%.4f'%x for x in ck]}")
def devsq(xs):
    m=sum(xs)/len(xs); return sum((x-m)**2 for x in xs)
def avg(xs): return sum(xs)/len(xs)
w=[1,1,1,0,0,0]
for norm in (0,1):
    of=oc=op=oa=0.0
    for qi in range(6):
        fs=[grid[(c,qi)] for c in FS_cols]; ck=[grid[(c,qi)] for c in CK_cols]; allv=fs+ck
        km,lm,mm=avg(fs),avg(ck),avg(allv)
        f=(devsq(fs)/km**2 if norm else devsq(fs))
        c=(devsq(ck)/lm**2 if norm else devsq(ck))
        p=((km-lm)**2/mm**2 if norm else (km-lm)**2)
        a=(devsq(allv)/mm**2 if norm else devsq(allv))
        of+=w[qi]*f; oc+=w[qi]*c; op+=w[qi]*p; oa+=w[qi]*a
    print(f"norm={norm}: WITHIN-FS={of:.6g}  WITHIN-CK={oc:.6g}  CKvsFS={op:.6g}  ALL={oa:.6g}")

# ---- build cells dict ----
cells = {}  # 'A1' -> ('s', text) string | ('f', formula) | ('n', number)
def S(ref,text): cells[ref]=('s',text)
def F(ref,formula): cells[ref]=('f',formula)
def N(ref,num): cells[ref]=('n',num)

S('A1','GLOBAL MINIMIZATION  (live consistency objective)')
S('A2','Values pulled live from the 4 FS sheets. Point Solver at E13/E14/E15/E16. See notes col S.')
# header row 3
S('A3','Quantity')
hdr={'B':'NoVs111 FS','C':'NoVp111 FS','D':'NoVs100 FS','E':'NoVp100 FS',
     'F':'NoVs111 CK','G':'NoVp111 CK','H':'NoVs100 CK','I':'NoVp100 CK',
     'J':'weight','K':'FS mean','L':'CK mean','M':'All mean',
     'N':'within-FS SS','O':'within-CK SS','P':'CKvsFS','Q':'ALL SS'}
for c,t in hdr.items(): S(c+'3',t)
# data rows 4..9
for qi,q in enumerate(qlabels):
    r=4+qi
    S('A%d'%r,q)
    for col,(sh,clist) in cols.items():
        F('%s%d'%(col,r), "'%s'!%s"%(sh,clist[qi]))
    N('J%d'%r, w[qi])
    F('K%d'%r, "AVERAGE(B%d:E%d)"%(r,r))
    F('L%d'%r, "AVERAGE(F%d:I%d)"%(r,r))
    F('M%d'%r, "AVERAGE(B%d:I%d)"%(r,r))
    F('N%d'%r, "J{r}*IF($D$11=1,DEVSQ(B{r}:E{r})/K{r}^2,DEVSQ(B{r}:E{r}))".format(r=r))
    F('O%d'%r, "J{r}*IF($D$11=1,DEVSQ(F{r}:I{r})/L{r}^2,DEVSQ(F{r}:I{r}))".format(r=r))
    F('P%d'%r, "J{r}*IF($D$11=1,(K{r}-L{r})^2/M{r}^2,(K{r}-L{r})^2)".format(r=r))
    F('Q%d'%r, "J{r}*IF($D$11=1,DEVSQ(B{r}:I{r})/M{r}^2,DEVSQ(B{r}:I{r}))".format(r=r))
# toggle + objectives
S('A11','Normalize? 1=relative, 0=absolute (raw GPa^2)')
N('D11',0)
S('A13','Objective: WITHIN-FS spread (make 4 FS agree)');   F('E13','SUM(N4:N9)')
S('A14','Objective: WITHIN-CK spread (make 4 Cook agree)'); F('E14','SUM(O4:O9)')
S('A15','Objective: CK vs FS (between-method means)');       F('E15','SUM(P4:P9)')
S('A16','Objective: ALL 8 (total spread CK+FS)');            F('E16','SUM(Q4:Q9)')
# notes
notes = [
 (3, 'HOW TO USE:'),
 (4, '1) weights J4:J9 pick which constants count (default C11/C12/C44=1, derivs=0).'),
 (5, '2) D11=0 uses raw GPa^2; D11=1 normalizes each by mean^2 so all count equally.'),
 (6, '3) Data > Solver: Set Objective = E13 (or E14/E15/E16), To: Min.'),
 (7, '4) By Changing Cells = your SHARED physical inputs (see guide): repoint them'),
 (8, '   to one master first, else Solver only affects one sheet.'),
 (9, '5) Re-run the per-combo Solvers after, then re-read these objectives.'),
 (11,'Between = ALL - (withinFS+withinCK); minimizing E16 drives full agreement.'),
]
for r,t in notes: S('S%d'%r,t)

# ---- Deliverable B: joint-fit data-misfit objective (live) ----
S('A18','JOINT-FIT objective (Deliverable B) - per-combo data-misfit totals (live)')
S('A19','combo'); S('B19','FS misfit'); S('C19','CK misfit')
misfit = {
 'NoVs111': ("FS No Vs 111",'BK32','AV30'),
 'NoVp111': ("FS No Vp 111",'BH29','AQ34'),
 'NoVs100': ("FS No Vs 100",'BK29','AQ32'),
 'NoVp100': ("FS No Vp 100",'BM30','AS32'),
}
r=20
for combo,(sh,fs,ck) in misfit.items():
    S('A%d'%r,combo); F('B%d'%r,"'%s'!%s"%(sh,fs)); F('C%d'%r,"'%s'!%s"%(sh,ck)); r+=1
S('A24','SUM = joint objective'); F('B24','SUM(B20:B23)'); F('C24','SUM(C20:C23)')
S('A25','Joint ALL (FS+CK)'); F('B25','B24+C24')
notes2=[
 (14,'JOINT FIT (rigorous): repoint all 4 FS param blocks to ONE shared 6-cell'),
 (15,'  set, then Solver: min B24 by changing those 6 shared Cij. Same for CK (C24).'),
 (16,'  Grand joint fit: min B25. Variables here are Cij (inherent to a joint fit).'),
]
for rr,t in notes2: S('S%d'%rr,t)

# ---- serialize worksheet xml ----
from openpyxl.utils import column_index_from_string
def cell_xml(ref):
    typ,val=cells[ref]
    if typ=='s':
        return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'%(ref,escape(str(val)))
    if typ=='f':
        return '<c r="%s"><f>%s</f></c>'%(ref,escape(val))
    return '<c r="%s"><v>%s</v></c>'%(ref,val)
# group by row
rows={}
for ref in cells:
    m=re.match(r'([A-Z]+)(\d+)',ref); col,row=m.group(1),int(m.group(2))
    rows.setdefault(row,[]).append((column_index_from_string(col),ref))
maxrow=max(rows);
sheetdata=[]
for r in sorted(rows):
    cellrefs=[ref for _,ref in sorted(rows[r])]
    sheetdata.append('<row r="%d">%s</row>'%(r,''.join(cell_xml(x) for x in cellrefs)))
ws_xml=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
 '<dimension ref="A1:S%d"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
 '<sheetFormatPr defaultRowHeight="15"/><cols>'
 '<col min="1" max="1" width="42"/><col min="19" max="19" width="70"/></cols>'
 '<sheetData>%s</sheetData></worksheet>')%(maxrow,''.join(sheetdata))

# ---- repackage zip, add sheet7, patch workbook/rels/content-types, drop calcChain ----
zin=zipfile.ZipFile(SRC)
names=zin.namelist()
wbxml=zin.read('xl/workbook.xml').decode('utf8')
relsxml=zin.read('xl/_rels/workbook.xml.rels').decode('utf8')
ctxml=zin.read('[Content_Types].xml').decode('utf8')

# add sheet element at end of <sheets>
wbxml=wbxml.replace('</sheets>','<sheet name="Global Min" sheetId="7" r:id="rId11"/></sheets>')
# force full recalc on load
wbxml=wbxml.replace('<calcPr calcId="191029"/>','<calcPr calcId="191029" fullCalcOnLoad="1"/>')
# drop calcChain rel + content-type (we will not write calcChain.xml)
relsxml=re.sub(r'<Relationship Id="rId10"[^>]*calcChain[^>]*/>','',relsxml)
relsxml=relsxml.replace('</Relationships>',
   '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/></Relationships>')
ctxml=re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>','',ctxml)
ctxml=ctxml.replace('</Types>',
   '<Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')

zout=zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    if item.filename=='xl/calcChain.xml': continue
    data=zin.read(item.filename)
    if item.filename=='xl/workbook.xml': data=wbxml.encode('utf8')
    elif item.filename=='xl/_rels/workbook.xml.rels': data=relsxml.encode('utf8')
    elif item.filename=='[Content_Types].xml': data=ctxml.encode('utf8')
    zout.writestr(item,data)
zout.writestr('xl/worksheets/sheet7.xml',ws_xml.encode('utf8'))
zout.close(); zin.close()
print('\nWROTE:',OUT)
print('size',os.path.getsize(OUT))
