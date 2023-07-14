VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XcodeX4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim X1 As New XcodeX1
Dim X3 As New XcodeX3
Dim errorCount As Integer, ThisTempLine As String, nofixError As Integer


Public Sub AutoCorrectX_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim modified As Boolean, ce As Range, thisComment As String, pcount As Double, pfullcount As Double
        errorCount = 0: nofixError = 0
        
        XCodeProgress.setPP pp
        XCodeProgress.Show
        XCodeProgress.ProgressBar_val pcount
        pfullcount = pp.Range(pp.Range("B3"), pp.Range("B2").End(xlDown)).Count
           
        
        thermoMaskCheck pp
        ioexpFCUparamsCheck pp
        device_def_check pp
        dev_idpg_bell_mask pp
        
        For Each ce In pp.Range(pp.Range("B3"), pp.Range("B2").End(xlDown))
            ThisTempLine = pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text)
            thisComment = ""
            If Len(ThisTempLine) > 0 And (Not InStr(ThisTempLine, "//") = 1) Then
                
                If InStr(ThisTempLine, "//") > 2 Then
                    thisComment = " " & Mid(ThisTempLine, InStr(ThisTempLine, "//"), Len(ThisTempLine))
                    ThisTempLine = pp.WorksheetFunction.Trim(Mid(ThisTempLine, 1, InStr(ThisTempLine, "//") - 1))
                End If
                
                symbolCheck1 ThisTempLine, "#", ce.Text
                symbolCheck1 ThisTempLine, "$", ce.Text
                symbolCheck2 ThisTempLine, "=", ce.Text
                symbolCheck2 ThisTempLine, ".", ce.Text
                symbolCheck2 ThisTempLine, ",", ce.Text
                symbolCheck2 ThisTempLine, "-", ce.Text
                symbolCheck2 ThisTempLine, "_", ce.Text
                
                All_dataTypeCheck ThisTempLine, ce.Text
                             
                                
                
                
                
                If pp.WorksheetFunction.Trim(ThisTempLine & thisComment) <> ce.Offset(0, 1).Text Then
                    ce.Offset(0, 1).FormulaR1C1 = pp.WorksheetFunction.Trim(ThisTempLine & thisComment)
                End If
        
            End If
            
            pcount = pcount + 1
            XCodeProgress.ProgressBar_val (pcount * 100 / pfullcount)
            XCodeProgress.Caption = "Auto Correct Wizard | Checking " & CInt(pcount * 100 / pfullcount) & "%, So far " & errorCount & " mistakes fixed"
            XCodeProgress.Label1.Caption = ce.Offset(0, 1).Text
            
        Next ce
        
        If errorCount > 0 Or nofixError > 0 Then
            If errorCount < 5 Then
                MsgBox "Auto Corrected " & errorCount & " errors. Almost there keep eye on those errors", vbExclamation, "XCODE | Error_Checker"
            ElseIf errorCount < 10 Then
                MsgBox "Auto Corrected " & errorCount & " errors. Excuse me! little bit out of control here!", vbExclamation, "XCODE | Error_Checker"
            ElseIf errorCount < 20 Then
                MsgBox "Auto Corrected " & errorCount & " errors. I am embarrassed to be your compiler", vbExclamation, "XCODE | Error_Checker"
            Else
                MsgBox "Auto Corrected " & errorCount & " errors. Please don't do XCODE", vbExclamation, "XCODE | Error_Checker"
            End If
        Else
            MsgBox "No Syntax erros found, good job", vbInformation, "XCODE | Error_Checker"
        End If
        X3.Color_ThisXcode_Function pp
        
    End If
End Sub


Private Sub thermoMaskCheck(pp As Application)
    Dim thermo_func As String, thermo_line As String, thisRCU As String
    thisRCU = X1.getRCUnumber(pp)
    thermo_func = Replace(pp.Range("Device" & thisRCU & "!N25").Text, " ", "")
    thermo_line = Replace(Replace(pp.Range("Device" & thisRCU & "!J25").Text, " ", ""), ";", "")
    thermo_line = Right(thermo_line, 2) & Left(thermo_line, 2)
    
    If Len(thermo_func) > 0 And Not (LCase(Left(thermo_line, 1)) = "8" Or LCase(Left(thermo_line, 1)) = "c") Then
        MSG1 = MsgBox("Enable proper mask for " & thermo_func & "?", vbYesNo + vbQuestion, "Enable Proper Mask?")
        If MSG1 = vbYes Then
            thermo_line = "8" & Right(thermo_line, 3)
            thermo_line = Right(thermo_line, 2) & " " & Left(thermo_line, 2) & ";"
            pp.Range("Device" & thisRCU & "!J25").FormulaR1C1 = thermo_line
        End If
    End If
End Sub

Private Sub ioexpFCUparamsCheck(pp As Application)
    Dim IOwithFCU As Boolean, thisRCU As String, FCUType As String
    thisRCU = X1.getRCUnumber(pp)
    IOwithFCU = False
    
    For fcu = 1 To 6
        FCUType = Replace(Replace(pp.Range("Settings" & thisRCU & "!B" & (60 + (fcu - 1) * 9)).Text, " ", ""), ";", "")
        
        If FCUType <> "00" Then
            If FCUType = "01" Then
                If Replace(pp.Range("Settings" & thisRCU & "!B" & (61 + (fcu - 1) * 9)).Text, " ", "") <> "00;" Then
                    IOwithFCU = True
                    Exit For
                End If
            ElseIf Left(FCUType, 1) = "4" Or Left(FCUType, 1) = "8" Then
                If Right(Replace(Replace(pp.Range("Settings" & thisRCU & "!B" & (61 + (fcu - 1) * 9)).Text, " ", ""), ";", ""), 1) <> "0" Then
                    IOwithFCU = True
                    Exit For
                End If
            End If
        End If
    Next fcu
    
    If IOwithFCU Then
        If Replace(pp.Range("Settings" & thisRCU & "!B115").Text, " ", "") = "00;" Then
            MSG1 = MsgBox("Found a FCU having IOEXP as its control output device." & vbNewLine & "Enable proper setting for force updating of expansion upon RCU startup ?", vbYesNo + vbQuestion, "Enable Proper Settings?")
            If MSG1 = vbYes Then
                pp.Range("Settings" & thisRCU & "!B115").FormulaR1C1 = "01;"
            End If
        End If
    End If
End Sub

Private Sub device_def_check(pp As Application)
    Dim ce As Range, thisRCU As String, k As Integer, thisCell As String
    thisRCU = X1.getRCUnumber(pp)
    For Each ce In pp.Range("Device" & thisRCU & "!B4:B" & X1.getDevLimit(pp, thisRCU))
         For k = 1 To 5
            thisCell = ce.Offset(0, k - 1).Text
            If X1.isValidHexString(Replace(thisCell, " ", "")) Then
                
            Else
                nofixError = nofixError + 1
                MsgBox "Invalid hex string in device definition " & ce.Offset(0, -1).Text & vbNewLine & vbNewLine & ce.Offset(0, -1).Text & " [" & ce.Text & " " & ce.Offset(0, 1).Text & " " & ce.Offset(0, 2).Text & " " & ce.Offset(0, 3).Text & " " & ce.Offset(0, 4).Text & " " & ce.Offset(0, 5).Text & "]", vbExclamation, "XCODE | Error_Checker"
            End If
         Next k
    Next ce
End Sub

Private Sub dev_idpg_bell_mask(pp As Application)
    Dim ce As Range, thisRCU As String, idpg_def As String, idpg_line As Integer, bell_line As String, bell_mask As String
    thisRCU = X1.getRCUnumber(pp)
    For Each ce In pp.Range("Device" & thisRCU & "!B4:B" & X1.getDevLimit(pp, thisRCU))
        If Replace(ce.Offset(0, 1).Text, " ", "") = "01" Then
            idpg_def = Replace(ce.Offset(0, 4).Text, " ", "")
            idpg_def = Right(idpg_def, 2) & Left(idpg_def, 2)
            idpg_line = pp.WorksheetFunction.Hex2Dec(idpg_def) + 4
            
            If Len(Replace(pp.Range("Device" & thisRCU & "!N" & (idpg_line + 1)).Text, " ", "")) > 0 Then
                bell_line = Replace(Replace(pp.Range("Device" & thisRCU & "!J" & (idpg_line + 1)).Text, " ", ""), ";", "")
                bell_mask = LCase(Mid(bell_line, 3, 1))
                
                If Not (bell_mask = "c" Or bell_mask = "8") Then
                    MSG1 = MsgBox("Found, IDPG bell not using proper mask." & vbNewLine & "Enable proper mask (c) for managing bell from FW ?", vbYesNo + vbQuestion, "Enable Proper Mask?")
                    If MSG1 = vbYes Then
                        pp.Range("Device" & thisRCU & "!J" & (idpg_line + 1)).FormulaR1C1 = Left(bell_line, 2) & " c" & Right(bell_line, 1) & ";"
                    End If
                End If
            End If
        End If
    Next ce
End Sub


Private Sub symbolCheck1(ThisTempLine As String, symbol As String, lineNum As String)
    If InStr(ThisTempLine, symbol) Then
        i = InStr(ThisTempLine, symbol)
        While i > 1
            If Mid(ThisTempLine, i + 1, 1) = " " Then
                MsgBox "An error fixed at Line " & lineNum & vbNewLine & "Detected a space after [" & symbol & "]", vbExclamation, "XCODE | Error_Checker"
                ThisTempLine = Left(ThisTempLine, i) & Mid(ThisTempLine, i + 2, Len(ThisTempLine))
                errorCount = errorCount + 1
            End If
            i = InStr(i + 1, ThisTempLine, symbol)
        Wend
        
        i = InStr(ThisTempLine, symbol)
        While i > 1
            If Mid(ThisTempLine, i - 1, 1) <> " " Then
                MsgBox "An error fixed at Line " & lineNum & vbNewLine & "Not detecting a space before [" & symbol & "]", vbExclamation, "XCODE | Error_Checker"
                ThisTempLine = Left(ThisTempLine, i - 1) & " " & Mid(ThisTempLine, i, Len(ThisTempLine))
                errorCount = errorCount + 1
            End If
            i = InStr(i + 1, ThisTempLine, symbol)
        Wend
    End If
End Sub

Private Sub symbolCheck2(ThisTempLine As String, symbol As String, lineNum As String)
    If InStr(ThisTempLine, symbol) Then
        i = InStr(ThisTempLine, symbol)
        While i > 1
            If Mid(ThisTempLine, i + 1, 1) = " " Then
                MsgBox "An error fixed at Line " & lineNum & vbNewLine & "Detected a space after [" & symbol & "]", vbExclamation, "XCODE | Error_Checker"
                ThisTempLine = Left(ThisTempLine, i) & Mid(ThisTempLine, i + 2, Len(ThisTempLine))
                errorCount = errorCount + 1
            End If
            i = InStr(i + 1, ThisTempLine, symbol)
        Wend
        
        i = InStr(ThisTempLine, symbol)
        While i > 1
            If Mid(ThisTempLine, i - 1, 1) = " " Then
                MsgBox "An error fixed at Line " & lineNum & vbNewLine & "Detected a space before [" & symbol & "]", vbExclamation, "XCODE | Error_Checker"
                ThisTempLine = Left(ThisTempLine, i - 2) & Mid(ThisTempLine, i, Len(ThisTempLine))
                errorCount = errorCount + 1
            End If
            i = InStr(i + 1, ThisTempLine, symbol)
        Wend
    End If
End Sub

Private Sub dataTypesCheck(ThisTempLine As String, misSpell As String, correctSpell As String, lineNum As String)
    If InStr(ThisTempLine, " $" & misSpell & "=") Then
        MsgBox "An error fixed at Line " & lineNum & vbNewLine & "Invalid parameter $" & misSpell & " replaced with $" & correctSpell, vbExclamation, "XCODE | Error_Checker"
        ThisTempLine = Replace(ThisTempLine, " $" & misSpell & "=", " $" & correctSpell & "=")
        errorCount = errorCount + 1
    End If
End Sub

Private Sub All_dataTypeCheck(ThisTempLine As String, lineNum As String)
    Dim CList() As String, MList() As String, M_dataList() As String
    
    CList = Split("dec,hex,bin,relay,fan,type,display,sec,min,hour,bit,lum,len,reg,dev,ch,ballast,group,milisec,element,ip,var", ",")
    M_dataList = Split("ec,dc,de,xec,cec,fec,rec,eec,dwc,dsc,ddc,dfc,drc,d4c,d3c,def,ded,dex,edc,dce,ddec,deec,decc" & _
    "$ex,hx,he,gex,bex,nex,jex,uex,yex,hwx,hsx,hdx,hfx,hrx,h4x,h3x,hec,hed,hes,hez,ehx,hxe,hhex,heex,hexx" & _
    "$in,bn,bi,nin,hin,gin,vin,bun,bjn,bkn,bln,bon,b9n,b8n,bim,bij,bih,bib,ibn,bni,bbin,biin,binn" & _
    "$relays,realy,raely,elay,rlay,reay,rely,rela,eelay,delay,felay,gelay,telay,5elay,4elay,rwlay,rslay,rdlay,rflay,rrlay,r4lay,r3lay,rekay,repay,reoay,relzy,relxy,relsy,relwy,relqy,relat,relag,relah,relaj,relau,rela7,rela6,erlay,rleay,relya" & _
    "$an,fn,fa,dan,can,van,gan,tan,ran,fzn,fxn,fsn,fwn,fqn,fam,faj,fah,fab,afn,fna,ffan,faan,fann" & _
    "$ype,tpe,tye,typ,rype,fype,gype,hype,yype,6ype,5ype,ttpe,tgpe,thpe,tjpe,tupe,t7pe,t6pe,tyoe,tyle,ty-e,ty0e,typw,typs,typd,typf,typr,typ4,typ3,ytpe,tpye,tyep,ttype,tyype,typpe,typee,types" & _
    "$isplay,dsplay,diplay,dislay,dispay,disply,displa,sisplay,xisplay,cisplay,fisplay,risplay,eisplay,dusplay,djsplay,dksplay,dlsplay,dosplay,d9splay,d8splay,diaplay,dizplay,dixplay,didplay,dieplay,diwplay,disolay,disllay,dis-lay,dis0lay,dispkay,disppay,dispoay,displzy,displxy,displsy,displwy,displqy,displat,displag,displah,displaj,displau,displa7,displa,idsplay,dsiplay,dipslay,dislpay,dispaly,displya,ddisplay,diisplay,dissplay,dispplay,displlay,displaay,displayy" & _
    "$ec,sc,se,aec,zec,xec,eec,wec,swc,ssc,sdc,sfc,src,s4c,s3c,sev,sef,sed,sex,esc,sce,ssec,seec,secc" & _
    "$in,mn,mi,kin,jin,nin,mun,mjn,mkn,mln,mon,m9n,m8n,mim,mij,mih,mib,imn,mni,mmin,miin,minn,mins" & _
    "$our,hur,hor,hou,gour,bour,nour,jour,uour,your,hiur,hkur,hlur,hpur,h0ur,h9ur,hoyr,hohr,hojr,hokr,hoir,ho8r,ho7r,houe,houd,houf,houg,hout,hou5,hou4,ohur,huor,horu,hhour,hoour,houur,hourr,hours" & _
    "$bits,it,bt,bi,nit,hit,git,vit,but,bjt,bkt,blt,bot,b9t,b8t,bir,bif,big,bih,biy,bi6,bi5,ibt,bti,biit,bitt" & _
    "$um,lm,lu,kum,pum,oum,lym,lhm,ljm,lkm,lim,l8m,l7m,luk,luj,lun,ulm,lmu,llum,luum,lumm" & _
    "$en,ln,le,ken,pen,oen,lwn,lsn,ldn,lfn,lrn,l4n,l3n,lem,lej,leh,leb,eln,lne,llen,leen,lenn" & _
    "$eg,rg,re,eeg,deg,feg,geg,teg,5eg,4eg,rwg,rsg,rdg,rfg,rrg,r4g,r3g,ref,rev,reb,reh,rey,ret,erg,rge,rreg,reeg,regg" & _
    "$ev,dv,de,sev,xev,cev,fev,rev,eev,dwv,dsv,ddv,dfv,drv,d4v,d3v,deb,deg,def,edv,dve,ddev,deev,devv" & _
    "$h,c,vh,fh,dh,xh,cg,cb,cn,cj,cu,cy,hc,cch,chh" & _
    "$allast,bllast,balast,ballst,ballat,ballas,nallast,hallast,gallast,vallast,bzllast,bxllast,bsllast,bwllast,bqllast,baklast,baplast,baolast,balkast,balpast,baloast,ballzst,ballxst,ballsst,ballwst,ballqst,ballaat,ballazt,ballaxt,balladt,ballaet,ballawt,ballasr,ballasf,ballasg,ballash,ballasy,ballas6,ballas5,abllast,blalast,balalst,ballsat,ballats,bballast,baallast,balllast,ballaast,ballasst,ballastt,ballasts" & _
    "$roup,goup,grup,grop,grou,froup,vroup,broup,hroup,yroup,troup,geoup,gdoup,gfoup,ggoup,gtoup,g5oup,g4oup,griup,grkup,grlup,grpup,gr0up,gr9up,groyp,grohp,grojp,grokp,groip,gro8p,gro7p,grouo,groul,grou-,grou0,rgoup,gorup,gruop,gropu,ggroup,grroup,grooup,grouup,groupp,groups" & _
    "$ilisec,mlisec,miisec,milsec,miliec,milisc,milise,kilisec,jilisec,nilisec,mulisec,mjlisec,mklisec,mllisec,molisec,m9lisec,m8lisec,mikisec,mipisec,mioisec,milusec,miljsec,milksec,millsec,milosec,mil9sec,mil8sec,miliaec,milizec,milixec,milidec,milieec,miliwec,miliswc,milissc,milisdc,milisfc,milisrc,milis4c,milis3c,milisev,milisef,milised,milisex,imlisec,mliisec,miilsec,milsiec,miliesc,milisce,mmilisec,miilisec,millisec,miliisec,milissec,miliseec,milisecc" & _
    "$lement,eement,elment,eleent,elemnt,elemet,elemen,wlement,slement,dlement,flement,rlement,4lement,3lement,ekement,epement,eoement,elwment,elsment,eldment,elfment,elrment,el4ment,el3ment,elekent,elejent,elenent,elemwnt,elemsnt,elemdnt,elemfnt,elemrnt,elem4nt,elem3nt,elememt,elemejt,elemeht,elemebt,elemenr,elemenf,elemeng,elemenh,elemeny,elemen6,elemen5,leement,eelment,elmeent,eleemnt,elemnet,elemetn,eelement,ellement,eleement,elemment,elemeent,elemennt,elementt" & _
    "$p,i,up,jp,kp,lp,op,9p,8p,io,il,i-,i0,pi,iip,ipp" & _
    "$ar,vr,va,bar,gar,far,car,vzr,vxr,vsr,vwr,vqr,vae,vad,vaf,vag,vat,va5,va4,avr,vra,vvar,vaar,varr", "$")
    
    For k = LBound(CList) To UBound(CList)
        MList = Split(M_dataList(k), ",")
        For j = LBound(MList) To UBound(MList)
            dataTypesCheck ThisTempLine, MList(j), CList(k), lineNum
        Next j
    Next k
End Sub

