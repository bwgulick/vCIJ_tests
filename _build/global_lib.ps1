# Global fit builders: one shared 6-param Cij block fit across all 4 velocity combos.
#  Global_FS   : shared per-point length block (like the individual FS), model Cij shared,
#                measured Cij computed 4 ways; objective = dP^2 + sum over 4 combos of dCij^2.
#  Global_Cook : each combo keeps its OWN analytic Cook length (different integrand), but all
#                4 share the 6 Cij params; objective = sum of the 4 Cook combo misfits.

function ColL([int]$c){
    $s=''
    while($c -gt 0){ $m=($c-1)%26; $s=[char](65+$m)+$s; $c=[int](($c-$m-1)/26) }
    return $s
}

function Build-GlobalFS {
    param($wb,$afterSheet,[int]$n=22,[double[]]$seedParams)
    $ws=$wb.Worksheets.Add([System.Reflection.Missing]::Value,$afterSheet); $ws.Name='Global_FS'
    $hdr=@('Oil bar','P (GPa)','tP100','tS100','tP111','tS111',
      'L fit','L111','V0/V','rho','f','Vp100','Vs100','Vp111','Vs111','rVp111^2','rVs111^2',
      'C11 nVs111','C12 nVs111','C44 nVs111','C11 nVp111','C12 nVp111','C44 nVp111',
      'C11 nVs100','C12 nVs100','C44 nVs100','C11 nVp100','C12 nVp100','C44 nVp100',
      'P_FS','C11 mod','C12 mod','C44 mod','dP','gated sum','K bulk','GV','GR','GH')
    for($c=0;$c -lt $hdr.Count;$c++){ $ws.Cells.Item(3,1+$c).Value2=$hdr[$c] }
    for($i=0;$i -lt $n;$i++){
        $r=4+$i; $oil=30*$i; $seedL=0.704-($oil/630.0)*0.015
        $ws.Cells.Item($r,1).Formula="=Align!A$r"; $ws.Cells.Item($r,2).Formula="=Align!B$r"
        $ws.Cells.Item($r,3).Formula="=Align!C$r"; $ws.Cells.Item($r,4).Formula="=Align!D$r"
        $ws.Cells.Item($r,5).Formula="=Align!F$r"; $ws.Cells.Item($r,6).Formula="=Align!G$r"
        if($r -eq 4){ $ws.Cells.Item($r,7).Formula="=Lfin_100" } else { $ws.Cells.Item($r,7).Value2=[double]$seedL }
        $ws.Cells.Item($r,8).Formula="=G$r*Lfin_111/Lfin_100"
        $ws.Cells.Item($r,9).Formula="=(`$G`$4/G$r)^3"
        $ws.Cells.Item($r,10).Formula="=rho_100*I$r"
        $ws.Cells.Item($r,11).Formula="=0.5*(1-I$r^(2/3))"
        $ws.Cells.Item($r,12).Formula="=2*G$r/C$r"; $ws.Cells.Item($r,13).Formula="=2*G$r/D$r"
        $ws.Cells.Item($r,14).Formula="=2*H$r/E$r"; $ws.Cells.Item($r,15).Formula="=2*H$r/F$r"
        $ws.Cells.Item($r,16).Formula="=J$r*N$r^2"; $ws.Cells.Item($r,17).Formula="=J$r*O$r^2"
        # measured 4 combos
        $ws.Cells.Item($r,18).Formula="=J$r*L$r^2"                 # R  NoVs111 C11
        $ws.Cells.Item($r,19).Formula="=1.5*P$r-0.5*R$r-2*T$r"    # S  NoVs111 C12
        $ws.Cells.Item($r,20).Formula="=J$r*M$r^2"                # T  NoVs111 C44
        $ws.Cells.Item($r,21).Formula="=J$r*L$r^2"                # U  NoVp111 C11
        $ws.Cells.Item($r,22).Formula="=U$r+W$r-3*Q$r"           # V  NoVp111 C12
        $ws.Cells.Item($r,23).Formula="=J$r*M$r^2"                # W  NoVp111 C44
        $ws.Cells.Item($r,24).Formula="=J$r*L$r^2"                # X  NoVs100 C11
        $ws.Cells.Item($r,25).Formula="=X$r+Z$r-3*Q$r"           # Y  NoVs100 C12
        $ws.Cells.Item($r,26).Formula="=Q$r+0.5*P$r-0.5*X$r"     # Z  NoVs100 C44
        $ws.Cells.Item($r,27).Formula="=3*Q$r+AB$r-AC$r"         # AA NoVp100 C11
        $ws.Cells.Item($r,28).Formula="=P$r-Q$r-AC$r"            # AB NoVp100 C12
        $ws.Cells.Item($r,29).Formula="=J$r*M$r^2"               # AC NoVp100 C44
        # model (shared params col AP=42)
        $ws.Cells.Item($r,30).Formula="=-3*`$AP`$11*((1-2*K$r)^2.5)*(1+1.5*(4-`$AP`$12)*K$r)*K$r"   # AD P_FS
        $ws.Cells.Item($r,31).Formula="=((1-2*K$r)^3.5)*(`$AP`$4-`$AP`$14*K$r)+3*AD$r"   # AE C11m
        $ws.Cells.Item($r,32).Formula="=((1-2*K$r)^3.5)*(`$AP`$5-`$AP`$15*K$r)+1*AD$r"   # AF C12m
        $ws.Cells.Item($r,33).Formula="=((1-2*K$r)^3.5)*(`$AP`$6-`$AP`$16*K$r)+1*AD$r"   # AG C44m
        $ws.Cells.Item($r,34).Formula="=AD$r-B$r"                                          # AH dP
        $ws.Cells.Item($r,35).Formula=("=IF(AND(A$r>=CutoffLowBar,A$r<=CutoffOilBar),AH$r^2"+
          "+(AE$r-R$r)^2+(AF$r-S$r)^2+(AG$r-T$r)^2"+
          "+(AE$r-U$r)^2+(AF$r-V$r)^2+(AG$r-W$r)^2"+
          "+(AE$r-X$r)^2+(AF$r-Y$r)^2+(AG$r-Z$r)^2"+
          "+(AE$r-AA$r)^2+(AF$r-AB$r)^2+(AG$r-AC$r)^2,0)")                                 # AI gated
        $ws.Cells.Item($r,36).Formula="=(AE$r+2*AF$r)/3"
        $ws.Cells.Item($r,37).Formula="=(AE$r-AF$r+3*AG$r)/5"
        $ws.Cells.Item($r,38).Formula="=5*(AE$r-AF$r)*AG$r/(3*(AE$r-AF$r)+4*AG$r)"
        $ws.Cells.Item($r,39).Formula="=(AK$r+AL$r)/2"
    }
    $last=3+$n
    Write-ParamBlock $ws 42 $seedParams $last 35   # params col AP=42, objective sums AI (col35)
    Write-OutputBlock $ws 45 42 39                 # outputs at col45, params col42, GH per-row col39
    return $ws
}

function Build-GlobalCook {
    param($wb,$afterSheet,[int]$n=22,[double[]]$seedParams)
    $ws=$wb.Worksheets.Add([System.Reflection.Missing]::Value,$afterSheet); $ws.Name='Global_Cook'
    # shared cols
    $ws.Cells.Item(3,1).Value2='Oil bar'; $ws.Cells.Item(3,2).Value2='P (GPa)'
    $ws.Cells.Item(3,3).Value2='tP100'; $ws.Cells.Item(3,4).Value2='tS100'
    $ws.Cells.Item(3,5).Value2='tP111'; $ws.Cells.Item(3,6).Value2='tS111'
    $ws.Cells.Item(3,7).Value2='Vp100r'; $ws.Cells.Item(3,8).Value2='Vs100r'
    $ws.Cells.Item(3,9).Value2='Vp111r'; $ws.Cells.Item(3,10).Value2='Vs111r'
    for($i=0;$i -lt $n;$i++){
        $r=4+$i
        $ws.Cells.Item($r,1).Formula="=Align!A$r"; $ws.Cells.Item($r,2).Formula="=Align!B$r"
        $ws.Cells.Item($r,3).Formula="=Align!C$r"; $ws.Cells.Item($r,4).Formula="=Align!D$r"
        $ws.Cells.Item($r,5).Formula="=Align!F$r"; $ws.Cells.Item($r,6).Formula="=Align!G$r"
        $ws.Cells.Item($r,7).Formula="=2*Lfin_100/C$r"; $ws.Cells.Item($r,8).Formula="=2*Lfin_100/D$r"
        $ws.Cells.Item($r,9).Formula="=2*Lfin_111/E$r"; $ws.Cells.Item($r,10).Formula="=2*Lfin_111/F$r"
    }
    # 4 combo blocks, 22 cols each, starting col 11
    $combos=@(
      @{name='NoVs111'; v='=[I]{r}^2-(4/3)*[H]{r}^2';                  c11='=[rho]{r}*[vp100]{r}^2'; c44='=[rho]{r}*[vs100]{r}^2'; c12='=1.5*[rvp]{r}-0.5*[c11]{r}-2*[c44]{r}'},
      @{name='NoVp111'; v='=[G]{r}^2+(2/3)*[H]{r}^2-2*[J]{r}^2';       c11='=[rho]{r}*[vp100]{r}^2'; c44='=[rho]{r}*[vs100]{r}^2'; c12='=[c11]{r}+[c44]{r}-3*[rvs]{r}'},
      @{name='NoVs100'; v='=(1/3)*[I]{r}^2-(4/3)*[J]{r}^2+(2/3)*[G]{r}^2'; c11='=[rho]{r}*[vp100]{r}^2'; c44='=[rvs]{r}+0.5*[rvp]{r}-0.5*[c11]{r}'; c12='=[c11]{r}+[c44]{r}-3*[rvs]{r}'},
      @{name='NoVp100'; v='=[I]{r}^2-(4/3)*[H]{r}^2';                  c44='=[rho]{r}*[vs100]{r}^2'; c12='=[rvp]{r}-[rvs]{r}-[c44]{r}'; c11='=3*[rvs]{r}+[c12]{r}-[c44]{r}'}
    )
    $gateCols=@()
    for($k=0;$k -lt 4;$k++){
        $b=11+$k*22
        $cb=$combos[$k]
        # offsets
        $vc=ColL($b+0); $int=ColL($b+1); $trap=ColL($b+2); $intg=ColL($b+3); $len=ColL($b+4)
        $l111=ColL($b+5); $rho=ColL($b+6); $f=ColL($b+7); $vp100=ColL($b+8); $vs100=ColL($b+9)
        $vp111=ColL($b+10); $vs111=ColL($b+11); $rvp=ColL($b+12); $rvs=ColL($b+13)
        $c11=ColL($b+14); $c12=ColL($b+15); $c44=ColL($b+16); $pfs=ColL($b+17)
        $c11m=ColL($b+18); $c12m=ColL($b+19); $c44m=ColL($b+20); $gate=ColL($b+21)
        $gateCols += $gate
        # headers
        $hh=@("vc2 $($cb.name)","integrand","trap","integral","Cook L","L111","rho","f",
              "Vp100","Vs100","Vp111","Vs111","rVp111^2","rVs111^2","C11","C12","C44",
              "P_FS","C11mod","C12mod","C44mod","gated")
        for($j=0;$j -lt $hh.Count;$j++){ $ws.Cells.Item(3,$b+$j).Value2=$hh[$j] }
        for($i=0;$i -lt $n;$i++){
            $r=4+$i
            $repl=@{ '[I]'="I"; '[H]'="H"; '[G]'="G"; '[J]'="J"; '[rho]'=$rho; '[vp100]'=$vp100;
                     '[vs100]'=$vs100; '[rvp]'=$rvp; '[rvs]'=$rvs; '[c11]'=$c11; '[c12]'=$c12; '[c44]'=$c44 }
            # vcombo2
            $vf=$cb.v; foreach($kk in $repl.Keys){ $vf=$vf.Replace($kk,$repl[$kk]) }; $vf=$vf.Replace('{r}',"$r")
            $ws.Cells.Item($r,$b+0).Formula=$vf
            $ws.Cells.Item($r,$b+1).Formula="=4*Lfin_100^2/$vc$r"
            if($r -eq 4){ $ws.Cells.Item($r,$b+2).Value2=[double]0; $ws.Cells.Item($r,$b+3).Value2=[double]0 }
            else {
                $ws.Cells.Item($r,$b+2).Formula="=0.5*(B$r-B$($r-1))*($int$r+$int$($r-1))"
                $ws.Cells.Item($r,$b+3).Formula="=$intg$($r-1)+$trap$r"
            }
            $ws.Cells.Item($r,$b+4).Formula="=Lfin_100/(1+$intg$r*scal_100)"
            $ws.Cells.Item($r,$b+5).Formula="=$len$r*Lfin_111/Lfin_100"
            $ws.Cells.Item($r,$b+6).Formula="=rho_100*(Lfin_100/$len$r)^3"
            $ws.Cells.Item($r,$b+7).Formula="=0.5*(1-(Lfin_100/$len$r)^2)"
            $ws.Cells.Item($r,$b+8).Formula="=2*$len$r/C$r"
            $ws.Cells.Item($r,$b+9).Formula="=2*$len$r/D$r"
            $ws.Cells.Item($r,$b+10).Formula="=2*$l111$r/E$r"
            $ws.Cells.Item($r,$b+11).Formula="=2*$l111$r/F$r"
            $ws.Cells.Item($r,$b+12).Formula="=$rho$r*$vp111$r^2"
            $ws.Cells.Item($r,$b+13).Formula="=$rho$r*$vs111$r^2"
            # measured (order matters for NoVp100: c44,c12,c11)
            $mc11=$cb.c11; $mc12=$cb.c12; $mc44=$cb.c44
            foreach($kk in $repl.Keys){ $mc11=$mc11.Replace($kk,$repl[$kk]); $mc12=$mc12.Replace($kk,$repl[$kk]); $mc44=$mc44.Replace($kk,$repl[$kk]) }
            $ws.Cells.Item($r,$b+14).Formula=$mc11.Replace('{r}',"$r")
            $ws.Cells.Item($r,$b+15).Formula=$mc12.Replace('{r}',"$r")
            $ws.Cells.Item($r,$b+16).Formula=$mc44.Replace('{r}',"$r")
            # model (shared params col AP=42 -> we place params AFTER blocks; use named cells instead)
            $ws.Cells.Item($r,$b+17).Formula="=-3*gK0*((1-2*$f$r)^2.5)*(1+1.5*(4-gK0p)*$f$r)*$f$r"
            $ws.Cells.Item($r,$b+18).Formula="=((1-2*$f$r)^3.5)*(gC11_0-gC11a*$f$r)+3*$pfs$r"
            $ws.Cells.Item($r,$b+19).Formula="=((1-2*$f$r)^3.5)*(gC12_0-gC12a*$f$r)+1*$pfs$r"
            $ws.Cells.Item($r,$b+20).Formula="=((1-2*$f$r)^3.5)*(gC44_0-gC44a*$f$r)+1*$pfs$r"
            $ws.Cells.Item($r,$b+21).Formula="=IF(AND(A$r>=CutoffLowBar,A$r<=CutoffOilBar),($pfs$r-B$r)^2+($c11m$r-$c11$r)^2+($c12m$r-$c12$r)^2+($c44m$r-$c44$r)^2,0)"
        }
    }
    # shared param block far right (col 101), with named ranges g*
    $pc=101; $last=3+$n
    $ws.Cells.Item(3,$pc).Value2='SHARED FIT PARAMS'
    $plbl=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'"); $pn=@('gC11_0','gC12_0','gC44_0','gC11p','gC12p','gC44p')
    $PL=ColL($pc)
    for($i=0;$i -lt 6;$i++){
        $ws.Cells.Item(4+$i,$pc-1).Value2=[string]$plbl[$i]
        $ws.Cells.Item(4+$i,$pc).Value2=[double]$seedParams[$i]
        $wb.Names.Add($pn[$i], "=Global_Cook!`$$PL`$$(4+$i)") | Out-Null
    }
    $ws.Cells.Item(11,$pc-1).Value2='K0';   $ws.Cells.Item(11,$pc).Formula="=(gC11_0+2*gC12_0)/3"; $wb.Names.Add('gK0',"=Global_Cook!`$$PL`$11")|Out-Null
    $ws.Cells.Item(12,$pc-1).Value2="K0'";  $ws.Cells.Item(12,$pc).Formula="=(gC11p+2*gC12p)/3";   $wb.Names.Add('gK0p',"=Global_Cook!`$$PL`$12")|Out-Null
    $ws.Cells.Item(14,$pc-1).Value2='C11a'; $ws.Cells.Item(14,$pc).Formula="=3*gK0*(gC11p-3)-7*gC11_0"; $wb.Names.Add('gC11a',"=Global_Cook!`$$PL`$14")|Out-Null
    $ws.Cells.Item(15,$pc-1).Value2='C12a'; $ws.Cells.Item(15,$pc).Formula="=3*gK0*(gC12p-1)-7*gC12_0"; $wb.Names.Add('gC12a',"=Global_Cook!`$$PL`$15")|Out-Null
    $ws.Cells.Item(16,$pc-1).Value2='C44a'; $ws.Cells.Item(16,$pc).Formula="=3*gK0*(gC44p-1)-7*gC44_0"; $wb.Names.Add('gC44a',"=Global_Cook!`$$PL`$16")|Out-Null
    $sumParts = ($gateCols | ForEach-Object { "SUM($_`4:$_$last)" }) -join '+'
    $ws.Cells.Item(18,$pc-1).Value2='OBJECTIVE'; $ws.Cells.Item(18,$pc).Formula="=$sumParts"
    # outputs
    $oc=$pc+3; $ws.Cells.Item(3,$oc).Value2='OUTPUTS'
    $b0=11  # first block for shear at P=0
    $c11m0=ColL($b0+18); $c12m0=ColL($b0+19); $c44m0=ColL($b0+20)
    $outs=@(@('C11_0','=gC11_0'),@('C12_0','=gC12_0'),@('C44_0','=gC44_0'),@("C11'",'=gC11p'),@("C12'",'=gC12p'),@("C44'",'=gC44p'),
      @('K0','=gK0'),@("K0'",'=gK0p'),
      @('GV0',"=($c11m0`4-$c12m0`4+3*$c44m0`4)/5"),
      @('GR0',"=5*($c11m0`4-$c12m0`4)*$c44m0`4/(3*($c11m0`4-$c12m0`4)+4*$c44m0`4)"),
      @('GH0 (Hill)',"=((($c11m0`4-$c12m0`4+3*$c44m0`4)/5)+(5*($c11m0`4-$c12m0`4)*$c44m0`4/(3*($c11m0`4-$c12m0`4)+4*$c44m0`4)))/2"),
      @('misfit',"=$($PL)18"))
    for($i=0;$i -lt $outs.Count;$i++){ $ws.Cells.Item(4+$i,$oc).Value2=[string]$outs[$i][0]; $ws.Cells.Item(4+$i,$oc+1).Formula=$outs[$i][1] }
    return @{ ws=$ws; paramCol=$pc; last=$last }
}

# shared helpers for Global_FS param/output blocks (params at fixed col)
function Write-ParamBlock($ws,[int]$pc,[double[]]$seed,[int]$last,[int]$gateCol){
    $GC=ColL($gateCol); $PL=ColL($pc)
    $ws.Cells.Item(3,$pc).Value2='FIT PARAMS'
    $lbl=@('C11_0','C12_0','C44_0',"C11'","C12'","C44'")
    for($i=0;$i -lt 6;$i++){ $ws.Cells.Item(4+$i,$pc-1).Value2=[string]$lbl[$i]; $ws.Cells.Item(4+$i,$pc).Value2=[double]$seed[$i] }
    $ws.Cells.Item(11,$pc-1).Value2='K0';   $ws.Cells.Item(11,$pc).Formula="=(`$$PL`$4+2*`$$PL`$5)/3"
    $ws.Cells.Item(12,$pc-1).Value2="K0'";  $ws.Cells.Item(12,$pc).Formula="=(`$$PL`$7+2*`$$PL`$8)/3"
    $ws.Cells.Item(14,$pc-1).Value2='C11a'; $ws.Cells.Item(14,$pc).Formula="=3*`$$PL`$11*(`$$PL`$7-3)-7*`$$PL`$4"
    $ws.Cells.Item(15,$pc-1).Value2='C12a'; $ws.Cells.Item(15,$pc).Formula="=3*`$$PL`$11*(`$$PL`$8-1)-7*`$$PL`$5"
    $ws.Cells.Item(16,$pc-1).Value2='C44a'; $ws.Cells.Item(16,$pc).Formula="=3*`$$PL`$11*(`$$PL`$9-1)-7*`$$PL`$6"
    $ws.Cells.Item(18,$pc-1).Value2='OBJECTIVE (misfit)'; $ws.Cells.Item(18,$pc).Formula="=SUM(${GC}4:$GC$last)"
}
function Write-OutputBlock($ws,[int]$oc,[int]$pc,[int]$ghCol){
    $PL=ColL($pc); $GH=ColL($ghCol); $GV=ColL($ghCol-2); $GR=ColL($ghCol-1)
    $ws.Cells.Item(3,$oc).Value2='OUTPUTS'
    $outs=@(@('C11_0',"=`$$PL`$4"),@('C12_0',"=`$$PL`$5"),@('C44_0',"=`$$PL`$6"),
      @("C11'","=`$$PL`$7"),@("C12'","=`$$PL`$8"),@("C44'","=`$$PL`$9"),
      @('K0',"=`$$PL`$11"),@("K0'","=`$$PL`$12"),
      @('GV0',"=${GV}4"),@('GR0',"=${GR}4"),@('GH0 (Hill)',"=${GH}4"),@('misfit',"=`$$PL`$18"))
    for($i=0;$i -lt $outs.Count;$i++){ $ws.Cells.Item(4+$i,$oc).Value2=[string]$outs[$i][0]; $ws.Cells.Item(4+$i,$oc+1).Formula=$outs[$i][1] }
}
