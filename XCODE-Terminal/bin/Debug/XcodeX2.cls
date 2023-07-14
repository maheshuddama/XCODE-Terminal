VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XcodeX2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim X1 As New XcodeX1


Public Sub XcodeGen_Function(pp As Application)
    Dim startTime As String
    startTime = X1.getTimeStamp
    
    If X1.pointToXCode(pp) And X1.NoCriticalErrors(pp) Then
        
        pp.ScreenUpdating = False
        Dim thisRCUList As Variant, thisRCU As String, RCUname As String
        thisRCUList = X1.RCUList(pp)
        XCodeProgress.setPP pp ''
        If X1.compatible(pp) = 2 Then
        ''XCodeProgress.Show ''
        
            For L = LBound(thisRCUList) To UBound(thisRCUList)
                thisRCU = CStr(thisRCUList(L))
                
                If thisRCU = "" Then
                    RCUname = ""
                Else
                    RCUname = "RCU" & thisRCU & " - "
                End If
                
                'If X1.pointToXCode(pp) Then
                X1.ClearFunctionSeparaters pp, thisRCU
                X1.add_this_FUNCTION_sheet pp, thisRCU '4.9
                X1.ClearXCODE pp, thisRCU
                X1.print_XCODE_build pp, thisRCU '4.9
                X1.insert_XCODE_new_format pp, thisRCU '4.9
                
                Dim ce1 As Range, ce2 As Range
                Dim lineNum As Integer, Fcount As Integer, pcount As Double, pfullcount As Double
                Dim ThisTempLine As String, ThisHexCode As String
                lineNum = 0: Fcount = 0: pcount = 0: pfullcount = 0
                
                ''XCodeProgress.setPP pp
                XCodeProgress.Show
                XCodeProgress.ProgressBar_val pcount
                pfullcount = pp.Range(pp.Range("XCODE" & thisRCU & "!B2"), pp.Range("XCODE" & thisRCU & "!B2").End(xlDown)).Count
                
                
                For Each ce1 In pp.Range(pp.Range("XCODE" & thisRCU & "!B2"), pp.Range("XCODE" & thisRCU & "!B2").End(xlDown))
                           
                                    
                    If ce1.Offset(0, 1).Text <> "" Then
                        ThisTempLine = LCase(pp.WorksheetFunction.Trim(ce1.Offset(0, 1).Text))
                        
                        If InStr(ThisTempLine, "//") > 2 And Not Left(ThisTempLine, 2) = "0x" Then 'comment detection and filter except direct hex
                            ThisTempLine = pp.WorksheetFunction.Trim(Left(ThisTempLine, InStr(ThisTempLine, "//") - 1))
                        End If
                        
                        If Left(ThisTempLine, 2) = "//" Then
                            'do nothing this is a comment
                        ElseIf Left(ThisTempLine, 2) = "0x" Then
                            If X1.GrabHexString(ThisTempLine, pp) <> "FALSE" Then
                                X1.insertXCODE lineNum, X1.GrabComment(ThisTempLine), pp, thisRCU
                                X1.insertHEXCODE lineNum, X1.GrabHexString(ThisTempLine, pp), pp, thisRCU
                                lineNum = lineNum + 1
                            End If
                        ElseIf isOtherRCUcall(ThisTempLine) Then
                            X1.insertXCODE lineNum, "//" & ThisTempLine, pp, thisRCU
                            X1.insertHEXCODE lineNum, "01 01 00 00 00 00 00 00;", pp, thisRCU
                            lineNum = lineNum + 1
                        ElseIf isAFunction(ThisTempLine) Then
                            If X1.isNotDuplicateFunction(GrabFunctionName(ThisTempLine), pp, thisRCU) Then
                                X1.insertFunction Fcount, lineNum, GrabFunctionName(ThisTempLine), pp, thisRCU
                                Fcount = Fcount + 1
                            Else
                                'MsgBox "Duplicate Function Found " & GrabFunctionName(ThisTempLine) & " at " & ce1.Text 'replaced with errorlog ver 5.0
                                X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & ce1.Text & " | " & GrabFunctionName(ThisTempLine) & "][Duplicate Function]"
                            End If
                        ElseIf isAnIOEXPDirect(ThisTempLine) Then
                            Dim vIOEXPDirect As Variant
                            vIOEXPDirect = GrabIOEXPDirect(ThisTempLine, thisRCU, ce1.Text)
                            For i = LBound(vIOEXPDirect) To UBound(vIOEXPDirect)
                                X1.insertXCODE lineNum, CStr(vIOEXPDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isAnIODEXPDirect(ThisTempLine) Then
                            Dim vIODEXPDirect As Variant
                            vIODEXPDirect = GrabIODEXPDirect(ThisTempLine, thisRCU, ce1.Text)
                            For i = LBound(vIODEXPDirect) To UBound(vIODEXPDirect)
                                X1.insertXCODE lineNum, CStr(vIODEXPDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isARCUDirect(ThisTempLine) Then
                            Dim vRCUDirect As Variant
                            vRCUDirect = GrabRCUDirect(ThisTempLine)
                            For i = LBound(vRCUDirect) To UBound(vRCUDirect)
                                X1.insertXCODE lineNum, CStr(vRCUDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isAMathDirect(ThisTempLine) Then
                            Dim vMathDirect As Variant
                            vMathDirect = GrabMathDirect(ThisTempLine, thisRCU, ce1.Text)
                            For i = LBound(vMathDirect) To UBound(vMathDirect)
                                X1.insertXCODE lineNum, CStr(vMathDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isOtherVarDirect(ThisTempLine, pp, thisRCU) Then
                            Dim vOtherVarDirect As Variant
                            vOtherVarDirect = GrabOtherVarDirect(ThisTempLine, pp, thisRCU)
                            For i = LBound(vOtherVarDirect) To UBound(vOtherVarDirect)
                                X1.insertXCODE lineNum, CStr(vOtherVarDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isASaveDirect(ThisTempLine) Then
                            Dim vSaveDirect As Variant
                            vSaveDirect = GrabSaveDirect(ThisTempLine)
                            For i = LBound(vSaveDirect) To UBound(vSaveDirect)
                                X1.insertXCODE lineNum, CStr(vSaveDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isAModbusDirect(ThisTempLine) Then
                            Dim vModbusDirect As Variant
                            vModbusDirect = GrabModbusDirect(ThisTempLine)
                            For i = LBound(vModbusDirect) To UBound(vModbusDirect)
                                X1.insertXCODE lineNum, CStr(vModbusDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isADimmerDirect(ThisTempLine) Then
                            Dim vDimmerDirect As Variant
                            vDimmerDirect = GrabDimmerDirect(ThisTempLine, thisRCU, ce1.Text)
                            For i = LBound(vDimmerDirect) To UBound(vDimmerDirect)
                                X1.insertXCODE lineNum, CStr(vDimmerDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isADimmerReadDirect(ThisTempLine) Then
                            Dim vDimmerReadDirect As Variant
                            vDimmerReadDirect = GrabDimmerReadDirect(ThisTempLine, thisRCU, ce1.Text)
                            For i = LBound(vDimmerReadDirect) To UBound(vDimmerReadDirect)
                                X1.insertXCODE lineNum, CStr(vDimmerReadDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isADimmerLoad(ThisTempLine) Then
                            X1.insertXCODE lineNum, "load $hex=1", pp, thisRCU
                            lineNum = lineNum + 1
                            X1.insertXCODE lineNum, "//" & ThisTempLine, pp, thisRCU
                            X1.insertHEXCODE lineNum, LoadDimmerConvert(ThisTempLine, pp), pp, thisRCU
                            lineNum = lineNum + 1
                        ElseIf isADaliDirect(ThisTempLine) Then
                            Dim vDaliDirect As Variant
                            vDaliDirect = GrabDaliDirect(ThisTempLine, pp, thisRCU, ce1.Text)
                            For i = LBound(vDaliDirect) + 1 To UBound(vDaliDirect)
                                X1.insertXCODE lineNum, CStr(vDaliDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isADMXDirect(ThisTempLine) Then
                            Dim vDMXDirect As Variant
                            vDMXDirect = GrabDMXDirect(ThisTempLine, pp, thisRCU, ce1.Text)
                            For i = LBound(vDMXDirect) + 1 To UBound(vDMXDirect)
                                X1.insertXCODE lineNum, CStr(vDMXDirect(i)), pp, thisRCU
                                lineNum = lineNum + 1
                            Next i
                        ElseIf isADefinition(ThisTempLine) Then
                            'rcu definitions, get info later
                                
                        Else
                            If Len(Replace(ThisTempLine, " ", "")) < 2 Then
                                'ignore multiple space empty lines and lines less then 2 charactors
                            Else
                                X1.insertXCODE lineNum, ThisTempLine, pp, thisRCU
                                lineNum = lineNum + 1
                            End If
                        End If
                        
                    End If
                    
                    pcount = pcount + 1
                    XCodeProgress.ProgressBar_val (pcount * 50 / pfullcount)
                    XCodeProgress.Caption = RCUname & "Compiling " & CInt(pcount * 50 / pfullcount) & "% - " & X1.getfilename(pp)
                    XCodeProgress.Label1.Caption = ThisTempLine
                    pp.StatusBar = "XCODE | Compiler | Core - 1   " & String(((pcount * 50 / pfullcount) * 0.25), ChrW(9607)) & String(25 - ((pcount * 50 / pfullcount) * 0.25), ChrW(9620)) & " " & RCUname & "Compiling " & CInt(pcount * 50 / pfullcount) & "% - " & X1.getfilename(pp)
                
                Next ce1
                       
                X1.insertHEXCODE -2, ";", pp, thisRCU
                X1.insertHEXCODE -1, ";code", pp, thisRCU
                X1.insertStartupHEXCODE 0, pp.Range("Program" & thisRCU & "!F3").Text, pp, thisRCU
                lineNum = 1: pcount = 0
                XCodeProgress.ProgressBar_val 50
                pfullcount = pp.Range(pp.Range("Program" & thisRCU & "!F4"), pp.Range("Program" & thisRCU & "!F4").End(xlDown)).Count
                
                For Each ce2 In pp.Range(pp.Range("Program" & thisRCU & "!F4"), pp.Range("Program" & thisRCU & "!F4").End(xlDown))
                           
                    ThisTempLine = ce2.Text
                    ThisHexCode = "01 01 00 00 00 00 00 00;"
                    
                    If Left(ThisTempLine, 2) = "//" Then
                        ThisHexCode = ce2.Offset(0, -1).Text
                        X1.insertXCODE lineNum, pp.WorksheetFunction.Trim(Replace(ThisTempLine, "//", "")), pp, thisRCU
                    ElseIf hasAFunction(ThisTempLine) Then
                        ThisHexCode = FunctionConvert(ThisTempLine, ce2.Row, pp, thisRCU) 'ce2 for nearest exit
                    ElseIf isARCU(ThisTempLine) Then
                        ThisHexCode = RCUConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf isADMX(ThisTempLine) Then
                        ThisHexCode = DMXConvert(ThisTempLine, pp, thisRCU)
                    ElseIf isWaitABS(ThisTempLine) Then
                        ThisHexCode = WaitABSConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf hasTime(ThisTempLine) Then
                        ThisHexCode = WaitConvert(ThisTempLine, pp)
                    ElseIf isAnExit(ThisTempLine) Then
                        ThisHexCode = ExitConvert
                    ElseIf isDebug(ThisTempLine) Then
                        ThisHexCode = DebugConvert
                    ElseIf isVersion(ThisTempLine) Then
                        ThisHexCode = VersionConvert(ThisTempLine)
                    ElseIf isALoad(ThisTempLine) Then
                        ThisHexCode = LoadConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf isASend(ThisTempLine) Then
                        ThisHexCode = SendConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf isARead(ThisTempLine) Then
                        ThisHexCode = ReadConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf isASave(ThisTempLine) Then
                        ThisHexCode = SaveConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf isAMath(ThisTempLine) Then
                        ThisHexCode = MathConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    ElseIf isADimFlow(ThisTempLine) Then
                        ThisHexCode = DimFlowConvert(ThisTempLine, pp, thisRCU, ce2.Offset(0, -2).Text)
                    Else
                        'else part goes here, catching poorly written instructions
                        If InStr(ThisTempLine, ".set") > 1 Or InStr(ThisTempLine, ".unset") > 1 Or InStr(ThisTempLine, ".or") > 1 Or InStr(ThisTempLine, ".and") > 1 Or InStr(ThisTempLine, ".xor") > 1 Then  '[dropped from 5.2] Or InStr(ThisTempLine, "$") > 2 Then
                            'MsgBox ThisTempLine & vbNewLine & vbNewLine & "[" & Left(ThisTempLine, InStr(ThisTempLine, ".") - 1) & "] can be an undefined variable!" 'replaced with errorlog ver 5.0
                            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & ce2.Offset(0, -2).Text & " | " & ThisTempLine & "][Undefined variable | " & Left(ThisTempLine, InStr(ThisTempLine, ".") - 1) & "]"
                        ElseIf (InStr(ThisTempLine, "wait ") > 0) And (Not (InStr(ThisTempLine, "$hour") > 5 Or InStr(ThisTempLine, "$min") > 5 Or InStr(ThisTempLine, "$sec") > 5)) Then 'added from 5.2
                            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & ce2.Offset(0, -2).Text & " | " & ThisTempLine & "][Invalid parameter | " & Replace(ThisTempLine, "wait ", "") & "]"
                        End If
                    End If
                    
                    X1.insertHEXCODE lineNum, ThisHexCode, pp, thisRCU
                    lineNum = lineNum + 1
                    
                    pcount = pcount + 1
                    XCodeProgress.ProgressBar_val (50 + (pcount * 50 / pfullcount))
                    XCodeProgress.Caption = RCUname & "Compiling " & 50 + CInt(pcount * 50 / pfullcount) & "% - " & X1.getfilename(pp)
                    XCodeProgress.Label1.Caption = ThisHexCode
                    pp.StatusBar = "XCODE | Compiler | Core - 2   " & String(((50 + (pcount * 50 / pfullcount)) * 0.25), ChrW(9607)) & String(25 - ((50 + (pcount * 50 / pfullcount)) * 0.25), ChrW(9620)) & " " & RCUname & "Compiling " & CInt(50 + (pcount * 50 / pfullcount)) & "% - " & X1.getfilename(pp)
                
                Next ce2
                
                X1.setFunctionSeparaters pp, thisRCU
                X1.FunctionCalibrate pp, thisRCU
                ''Unload XCodeProgress
                ''XCodeProgress.setPP Nothing
                'End If
            
            Next L
            
    
            'if there are more than 1 rcu, linking other rcu
            
            If UBound(thisRCUList) - LBound(thisRCUList) > 0 Then 'And X1.pointToXCode(pp) Then
                For L = LBound(thisRCUList) To UBound(thisRCUList)
                    thisRCU = CStr(thisRCUList(L))
                    RCUname = "RCU" & thisRCU & " - "
                    pcount = 0
                    ''XCodeProgress.setPP pp
                    ''XCodeProgress.Show
                    XCodeProgress.ProgressBar_val pcount
                    pfullcount = pp.Range(pp.Range("Program" & thisRCU & "!D4"), pp.Range("Program" & thisRCU & "!D4").End(xlDown)).Count
                    
                    For Each ce3 In pp.Range(pp.Range("Program" & thisRCU & "!D4"), pp.Range("Program" & thisRCU & "!D4").End(xlDown))
                        ThisTempLine = ce3.Offset(0, 2).Text
                        
                        If isOtherRCUcall(ThisTempLine) Then
                            Dim otherRCU As String
                            otherRCU = X1.getOtherRCU(ThisTempLine, pp)
                            
                            If otherRCU <> "NULL" Then
                                Dim ThisTempLine2 As String, otherRCUIP As String, otherRCUF As String, thisFunction As String, LAarrayF() As String
                                otherRCUIP = "00 00 00 00": otherRCUF = "00 00"
                                For Each ce4 In pp.Range(pp.Range("XCODE" & thisRCU & "!B2"), pp.Range("XCODE" & thisRCU & "!B2").End(xlDown))
                                    ThisTempLine2 = ce4.Offset(0, 1).Text
                                    
                                    If isADefinition(ThisTempLine2) Then
                                    
                                        If InStr(ThisTempLine2, "//") > 20 Then 'filter comment with minimum distance
                                            ThisTempLine2 = Left(ThisTempLine2, InStr(ThisTempLine2, "//") - 1)
                                        End If
                                    
                                        If isAnIPdefinition(ThisTempLine2, otherRCU, thisRCU) Then
                                            otherRCUIP = Grab_IP_addr(Grab_variant(ThisTempLine2, "ip"))
                                            Exit For
                                        End If
                                    End If
                                Next ce4
                                
                                LArrayF = Split(ThisTempLine, " ")
                                
                                For f = LBound(LArrayF) To UBound(LArrayF)
                                    If InStr(LArrayF(f), "#") = 1 Then
                                        thisFunction = Replace(LArrayF(f), "#", "")
                                        otherRCUF = X1.getLineFromFunctionName(thisFunction, 3, pp, otherRCU) '3 input line for nearest exit
                                        otherRCUF = Right(otherRCUF, 2) & " " & Left(otherRCUF, 2)
                                    End If
                                Next f
                                
                                If otherRCUF <> "00 00" And otherRCUIP <> "00 00 00 00" Then
                                    ce3.Offset(0, 1).FormulaR1C1 = "30 " & otherRCUF & " 00 " & otherRCUIP & ";"
                                Else
                                    If otherRCUF = "00 00" Then
                                        'MsgBox """" & ThisTempLine & """ is an invalid function call" 'replaced in ver 5.0 with error log
                                        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | " & ThisTempLine & "][Invalid function call at RCU" & otherRCU & "]"
                                    End If
                                    If otherRCUIP = "00 00 00 00" Then
                                        'MsgBox "Missing RCU" & otherRCU & " definition" 'replaced in ver 5.0 with error log
                                        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Request functions from RCU" & otherRCU & "][Missing RCU" & otherRCU & " definition]"
                                    End If
                                End If
                            
                            Else
                                'MsgBox """" & ThisTempLine & """ has invalid RCU call"
                                X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | " & ThisTempLine & "][Invalid request of RCU]"
                            End If
                            
                        End If
                        
                        pcount = pcount + 1
                        XCodeProgress.ProgressBar_val (pcount * 100 / pfullcount)
                        XCodeProgress.Caption = RCUname & "Mapping " & CInt(pcount * 100 / pfullcount) & "% - " & X1.getfilename(pp)
                        XCodeProgress.Label1.Caption = ThisTempLine
                        pp.StatusBar = "XCODE | Compiler | Core - 3   " & String(((pcount * 100 / pfullcount) * 0.25), ChrW(9607)) & String(25 - ((pcount * 100 / pfullcount) * 0.25), ChrW(9620)) & " " & RCUname & "Mapping " & CInt(pcount * 100 / pfullcount) & "% - " & X1.getfilename(pp)
                        
                    Next ce3
                    
                    ''Unload XCodeProgress
                    ''XCodeProgress.setPP Nothing
                    
                Next L
                
            End If
            
    
            
            Unload XCodeProgress ''
            
            
            'dev size check
            X1.devSizeCheck pp
            'check progbase
            X1.checkProgbase pp
            'set version
            X1.set_version pp '4.9
            'delete temporary sheets
            X1.remove_all_FUNCTION_sheet pp '4.9
            'create hex files
            X1.SaveAllXcode pp
            'error log ver 5.0
            X1.composeErrorReport pp, "Logic compiled with error(s)", startTime
            
        End If
        
        XCodeProgress.setPP Nothing ''
        pp.StatusBar = False
        pp.ScreenUpdating = True
    Else
        X1.composeErrorReport pp, "Logic did not compile due to critical error(s)", startTime
        
        pp.StatusBar = False
        pp.ScreenUpdating = True
    End If
End Sub


'function headers
Private Function isAFunction(inputString As String) As Boolean
    If InStr(LCase(inputString), "function ") = 1 Then
        isAFunction = True
    Else
        isAFunction = False
    End If
End Function

Private Function hasAFunction(inputString As String) As Boolean
    If InStr(LCase(inputString), " #") > 0 Then
        hasAFunction = True
    Else
        hasAFunction = False
    End If
End Function

Private Function GrabFunctionName(inputString As String) As String
    Dim LArray() As String
    LArray = Split(inputString, " ")
    GrabFunctionName = LArray(1)
End Function


'definition header
Private Function isADefinition(inputString As String) As Boolean
    If InStr(LCase(inputString), "define ") = 1 Then
        isADefinition = True
    Else
        isADefinition = False
    End If
End Function

'ip definition
Private Function isAnIPdefinition(inputString As String, otherRCU As String, thisRCU As String) As Boolean
    Dim otherRCUIP As String
    otherRCUIP = Grab_variant(inputString, "ip")
    isAnIPdefinition = False
    If otherRCUIP <> "NULL" And InStr(inputString, "define ") = 1 And InStr(inputString, "rcu" & otherRCU & " ") > 1 Then
        If X1.isAValidIP(otherRCUIP) Then
            isAnIPdefinition = True
        Else
            'not valid IP address
            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Definition of RCU" & otherRCU & " $ip=" & otherRCUIP & "][Invalid IP address]"
        End If
    End If
End Function

Private Function isOtherRCUcall(inputString As String) As Boolean
    'If (InStr(inputString, "run ") = 1 Or InStr(inputString, "reuse ") = 1) And InStr(inputString, " #") > 1 And InStr(inputString, " @rcu") > 1 Then
    If (InStr(inputString, "run ") = 1 Or InStr(inputString, "reuse ") = 1) And InStr(inputString, " @rcu") > 1 Then '4.9 removed #
        isOtherRCUcall = True
    Else
        isOtherRCUcall = False
    End If
End Function

'ioexp direct
Private Function isAnIOEXPDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "ioexp") = 1 And (InStr(Replace(LCase(inputString), " ", ""), ".") = 6 Or InStr(Replace(LCase(inputString), " ", ""), ".") = 7) Then
        isAnIOEXPDirect = True
    Else
        isAnIOEXPDirect = False
    End If
End Function

Private Function GrabIOEXPDirect(inputString As String, thisRCU As String, thisLine As String) As Variant
    Dim ioString As String, LArray() As String, ioName As String, L As Integer
    Dim vGrabIOEXPDirect As Variant
    ReDim vGrabIOEXPDirect(1 To 1)
    'detect ioexp
    If Mid(inputString, InStr(inputString, "ioexp") + 5, 1) = "." Then
        ioName = "ioexp1"
    Else
        ioName = "ioexp" & Mid(inputString, InStr(inputString, "ioexp") + 5, 1)
    End If
    
    If InStr(inputString, "read") = (InStr(inputString, ".") + 1) Then
        If InStr(inputString, " relay") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=4 $len=2"
        ElseIf InStr(inputString, " aout8") > 9 Or InStr(inputString, " analog8") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=43 $len=2"
        ElseIf InStr(inputString, " aout7") > 9 Or InStr(inputString, " analog7") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=39 $len=2"
        ElseIf InStr(inputString, " aout6") > 9 Or InStr(inputString, " analog6") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=35 $len=2"
        ElseIf InStr(inputString, " aout5") > 9 Or InStr(inputString, " analog5") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=31 $len=2"
        ElseIf InStr(inputString, " aout4") > 9 Or InStr(inputString, " analog4") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=27 $len=2"
        ElseIf InStr(inputString, " aout3") > 9 Or InStr(inputString, " analog3") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=23 $len=2"
        ElseIf InStr(inputString, " aout2") > 9 Or InStr(inputString, " analog2") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=19 $len=2"
        ElseIf InStr(inputString, " aout") > 9 Or InStr(inputString, " analog") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=15 $len=2"
        ElseIf InStr(inputString, " input") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=3 $len=2"
        ElseIf InStr(inputString, " pt100 2") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=9 $len=2"
        ElseIf InStr(inputString, " pt100") > 9 Then
            vGrabIOEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=7 $len=2"
        Else
            vGrabIOEXPDirect(1) = inputString
            'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid instruction]"
        End If
    Else
    
        ioString = Replace(Replace(inputString, " ", ""), "#", "$")
        ioString = Mid(ioString, 1, InStr(ioString, "$") - 1)
        LArray = Split(ioString, ".")
        
        If UBound(LArray) + 1 > 1 Then
            
            vGrabIOEXPDirect(1) = "load $var=" & ioName
            
            For L = LBound(LArray) + 1 To UBound(LArray)
                If (LArray(L) = "set" Or LArray(L) = "or") And Grab_value_0(inputString, L) <> "NULL" Then
                    ReDim Preserve vGrabIOEXPDirect(1 To UBound(vGrabIOEXPDirect) + 1)
                    vGrabIOEXPDirect(UBound(vGrabIOEXPDirect)) = "math.or $" & Grab_value_0(inputString, L)
                ElseIf (LArray(L) = "unset" Or LArray(L) = "and") And Grab_value_0(inputString, L) <> "NULL" Then
                    ReDim Preserve vGrabIOEXPDirect(1 To UBound(vGrabIOEXPDirect) + 1)
                    vGrabIOEXPDirect(UBound(vGrabIOEXPDirect)) = "math.and $" & Grab_value_0(inputString, L)
                ElseIf (LArray(L) = "xor") And Grab_value_0(inputString, L) <> "NULL" Then
                    ReDim Preserve vGrabIOEXPDirect(1 To UBound(vGrabIOEXPDirect) + 1)
                    vGrabIOEXPDirect(UBound(vGrabIOEXPDirect)) = "math.xor $" & Grab_value_0(inputString, L)
                End If
            Next L
            
            ReDim Preserve vGrabIOEXPDirect(1 To UBound(vGrabIOEXPDirect) + 1)
            vGrabIOEXPDirect(UBound(vGrabIOEXPDirect)) = "send.mod $dev=" & ioName & " $reg=4 $len=2"
            ReDim Preserve vGrabIOEXPDirect(1 To UBound(vGrabIOEXPDirect) + 1)
            vGrabIOEXPDirect(UBound(vGrabIOEXPDirect)) = "save $var=" & ioName
        Else
            vGrabIOEXPDirect(1) = "nop"
            'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IOEXP instruction]"
        End If
        
    End If
    
    GrabIOEXPDirect = vGrabIOEXPDirect
End Function

'iodexp direct
Private Function isAnIODEXPDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "iodexp") = 1 And (InStr(Replace(LCase(inputString), " ", ""), ".") = 7 Or InStr(Replace(LCase(inputString), " ", ""), ".") = 8) Then
        isAnIODEXPDirect = True
    Else
        isAnIODEXPDirect = False
    End If
End Function

Private Function GrabIODEXPDirect(inputString As String, thisRCU As String, thisLine As String) As Variant
    Dim ioString As String, LArray() As String, ioName As String, L As Integer
    Dim vGrabIODEXPDirect As Variant
    ReDim vGrabIODEXPDirect(1 To 1)
    'detect ioexp
    If Mid(inputString, InStr(inputString, "iodexp") + 6, 1) = "." Then
        ioName = "iodexp1"
    Else
        ioName = "iodexp" & Mid(inputString, InStr(inputString, "iodexp") + 6, 1)
    End If
    
    If InStr(inputString, "read") = (InStr(inputString, ".") + 1) Then
        If InStr(inputString, " aout8") > 9 Or InStr(inputString, " analog8") > 9 Or InStr(inputString, " ch8") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=34 $len=2"
        ElseIf InStr(inputString, " aout7") > 9 Or InStr(inputString, " analog7") > 9 Or InStr(inputString, " ch7") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=30 $len=2"
        ElseIf InStr(inputString, " aout6") > 9 Or InStr(inputString, " analog6") > 9 Or InStr(inputString, " ch6") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=26 $len=2"
        ElseIf InStr(inputString, " aout5") > 9 Or InStr(inputString, " analog5") > 9 Or InStr(inputString, " ch5") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=22 $len=2"
        ElseIf InStr(inputString, " aout4") > 9 Or InStr(inputString, " analog4") > 9 Or InStr(inputString, " ch4") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=18 $len=2"
        ElseIf InStr(inputString, " aout3") > 9 Or InStr(inputString, " analog3") > 9 Or InStr(inputString, " ch3") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=14 $len=2"
        ElseIf InStr(inputString, " aout2") > 9 Or InStr(inputString, " analog2") > 9 Or InStr(inputString, " ch2") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=10 $len=2"
        ElseIf InStr(inputString, " aout") > 9 Or InStr(inputString, " analog") > 9 Or InStr(inputString, " ch") > 9 Then
            vGrabIODEXPDirect(1) = "read.modbus $dev=" & ioName & " $reg=6 $len=2"
        Else
            vGrabIODEXPDirect(1) = "nop"
            'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IODEXP instruction]"
        End If
    Else
        vGrabIODEXPDirect(1) = "nop"
        'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IODEXP instruction]"
    End If
    
    GrabIODEXPDirect = vGrabIODEXPDirect
End Function


'rcu relay direct
Private Function isARCUDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "rcu.") = 1 And _
    (InStr(Replace(LCase(inputString), " ", ""), ".set$relay=") > 0 Or InStr(Replace(LCase(inputString), " ", ""), ".or$relay=") > 0 Or _
    InStr(Replace(LCase(inputString), " ", ""), ".unset$relay=") > 0 Or InStr(Replace(LCase(inputString), " ", ""), ".and$relay=") > 0 Or _
    InStr(Replace(LCase(inputString), " ", ""), ".xor$relay=") > 0) Then
        isARCUDirect = True
    Else
        isARCUDirect = False
    End If
End Function

Private Function GrabRCUDirect(inputString As String) As Variant
    
    Dim ioString As String, LArray() As String, LArray1() As String
    Dim L As Integer
    LArray = Split(Replace(inputString, "#", ""), " $")
    LArray1 = Split(LArray(0), ".")
    
    Dim vGrabRCUDirect As Variant
    ReDim vGrabRCUDirect(1 To 1)
    
    For L = 0 To UBound(LArray1) - 1
        ReDim Preserve vGrabRCUDirect(1 To L + 1)
        vGrabRCUDirect(L + 1) = "rcu." & LArray1(L + 1) & " $" & Grab_value_0(inputString, L + 1)
    Next L
    
    GrabRCUDirect = vGrabRCUDirect
End Function

'math direct
Private Function isAMathDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "math.") = 1 Then
        isAMathDirect = True
    Else
        isAMathDirect = False
    End If
End Function

Private Function GrabMathDirect(inputString As String, thisRCU As String, thisLine As String) As Variant
    
    Dim ioString As String, LArray() As String, LArray1() As String
    Dim L As Integer
    LArray = Split(Replace(inputString, "#", ""), " $")
    LArray1 = Split(LArray(0), ".")
    
    Dim vGrabMathDirect As Variant
    ReDim vGrabMathDirect(1 To 1)
    
    For L = 0 To UBound(LArray1) - 1
        ReDim Preserve vGrabMathDirect(1 To L + 1)
        If Not Grab_value_0(inputString, L + 1) = "NULL" Then
            vGrabMathDirect(L + 1) = "math." & LArray1(L + 1) & " $" & Grab_value_0(inputString, L + 1)
        Else
            'MsgBox "Invalid data type in """ & inputString & """ at """ & LArray1(L + 1) & """" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & " at """ & LArray1(L + 1) & """][Invalid math data type]"
            vGrabMathDirect(L + 1) = "math." & LArray1(L + 1)
        End If
    Next L
    
    GrabMathDirect = vGrabMathDirect
End Function

'save direct
Private Function isASaveDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "save") = 1 Then
        isASaveDirect = True
    Else
        isASaveDirect = False
    End If
End Function

Private Function GrabSaveDirect(inputString As String) As Variant
    Dim ioString As String, LArray() As String, LArray1() As String, ioLen As String
    Dim vGrabSaveDirect As Variant
    ReDim vGrabSaveDirect(1 To 1)
    
    If InStr(inputString, " $var=") > 3 And InStr(inputString, " $len=") > 3 Then
        'added in v5.3 specially for save settings with length
        'can inprove this entire function later
        LArray = Split(inputString, " $") 'filter with space
        
        For i = LBound(LArray) To UBound(LArray)
            If InStr(LArray(i), "var=") = 1 Then
                ioString = Replace(LArray(i), "var=", "")
            End If
            If InStr(LArray(i), "len=") = 1 Then
                ioLen = Replace(LArray(i), "len=", "")
            End If
        Next i
        LArray1 = Split(ioString, ",")
        
        For i = LBound(LArray1) To UBound(LArray1)
            ReDim Preserve vGrabSaveDirect(1 To i + 1)
            vGrabSaveDirect(i + 1) = LArray(0) & " $var=" & LArray1(i) & " $len=" & ioLen
        Next i
        
        GrabSaveDirect = vGrabSaveDirect
    
    ElseIf InStr(inputString, " $var=") > 3 Then
        LArray = Split(inputString, "$")
        
        For i = LBound(LArray) To UBound(LArray)
            If InStr(LArray(i), "var=") = 1 Then
                ioString = Replace(LArray(i), "var=", "")
            End If
        Next i
        LArray1 = Split(ioString, ",")
        
        
        For i = LBound(LArray1) To UBound(LArray1)
            ReDim Preserve vGrabSaveDirect(1 To i + 1)
            vGrabSaveDirect(i + 1) = LArray(0) & "$var=" & LArray1(i)
        Next i
        
        GrabSaveDirect = vGrabSaveDirect
    Else
        vGrabSaveDirect(1) = inputString
        GrabSaveDirect = vGrabSaveDirect
    End If
End Function

'extend modbus direct
Private Function isAModbusDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "send.mod") = 1 Then
        isAModbusDirect = True
    Else
        isAModbusDirect = False
    End If
End Function

Private Function GrabModbusDirect(inputString As String) As Variant
    Dim vGrabModbusDirect As Variant, len_check As String, val_check As String, var_check As String
    
    len_check = Grab_variant(inputString, "len")
    val_check = Grab_value_0(inputString, 1)
    'var_check = Grab_variant(inputString, "var")
    
    If len_check <> "NULL" And val_check <> "NULL" Then ' (val_check <> "NULL" Or var_check <> "NULL") Then
        ReDim vGrabModbusDirect(1 To 2)
        'If var_check <> "NULL" Then
            'vGrabModbusDirect(1) = "load $var=" & var_check
            'vGrabModbusDirect(2) = Replace(inputString, " $var=" & var_check, "")
        'Else
            vGrabModbusDirect(1) = "load $" & val_check
            vGrabModbusDirect(2) = Replace(inputString, " $" & val_check, "")
        'End If
        GrabModbusDirect = vGrabModbusDirect
    ElseIf len_check = "NULL" And val_check = "NULL" Then 'avoid len and val are both NULL 4.7
        ReDim vGrabModbusDirect(1 To 1)
        vGrabModbusDirect(1) = WorksheetFunction.Trim(inputString & " $len=2")
        GrabModbusDirect = vGrabModbusDirect
    Else
        ReDim vGrabModbusDirect(1 To 1)
        vGrabModbusDirect(1) = inputString
        GrabModbusDirect = vGrabModbusDirect
    End If
End Function

'expansion dimmer direct
Private Function isADimmerDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "send.dim") = 1 Then
        isADimmerDirect = True
    Else
        isADimmerDirect = False
    End If
End Function

Private Function GrabDimmerDirect(inputString As String, thisRCU As String, thisLine As String) As Variant
    Dim ioString As String, chString As String, LArray() As String, LArray1() As String, LArray2() As String
    Dim thisIOdev As String, thisChReg As Integer, ioLen As Integer
    
    thisIOdev = "ioexp1": ioString = "": chString = ""
    ioLen = 6
    LArray = Split(inputString, " $")
    
    For i = LBound(LArray) To UBound(LArray)
        If InStr(LArray(i), "dev=") = 1 Then
            ioString = Replace(LArray(i), "dev=", "")
        ElseIf InStr(LArray(i), "ch=") = 1 Then
            chString = Replace(LArray(i), "ch=", "")
        End If
    Next i
    
    Dim vGrabDimmerDirect As Variant
    ReDim vGrabDimmerDirect(1 To 1)
    
    If Len(ioString) > 0 And Len(chString) > 0 Then
        
        LArray1 = Split(ioString, ",")
        LArray2 = Split(X1.CommaExpand(chString, 1, 16), ",") 'added in 5.1
        
        If UBound(LArray1) < UBound(LArray2) Then
            For i = UBound(LArray1) + 1 To UBound(LArray2)
                ReDim Preserve LArray1(0 To i)
                LArray1(i) = LArray1(i - 1)
            Next i
        ElseIf UBound(LArray2) < UBound(LArray1) Then
            For i = UBound(LArray2) + 1 To UBound(LArray1)
                ReDim Preserve LArray2(0 To i)
                LArray2(i) = LArray2(i - 1)
            Next i
        End If
        
        
        For i = LBound(LArray1) To UBound(LArray1)
            If Len(LArray1(i)) > 1 And IsNumeric(Right(LArray1(i), 1)) Then
                thisIOdev = LArray1(i)
            ElseIf Len(LArray1(i)) > 1 And Not IsNumeric(Right(LArray1(i), 1)) Then
                thisIOdev = LArray1(i) & "1"
            ElseIf Len(LArray1(i)) = 1 And IsNumeric(Right(LArray1(i), 1)) Then
                thisIOdev = Mid(thisIOdev, 1, Len(thisIOdev) - 1) & LArray1(i)
            End If
            
            If IsNumeric(LArray2(i)) Then
                If CInt(LArray2(i)) > 0 And CInt(LArray2(i)) < 9 Then
                    thisChReg = CInt(LArray2(i))
                Else
                    thisChReg = 1
                End If
            Else
                thisChReg = 1
            End If
            
            If InStr(thisIOdev, "ioe") = 1 Then
                thisChReg = 12 + 4 * (thisChReg - 1)
            Else
                thisChReg = 3 + 4 * (thisChReg - 1)
            End If
            
            If InStr(inputString, " stop") > 1 Then
                thisChReg = thisChReg + 2
                ioLen = 2
            End If
            
            ReDim Preserve vGrabDimmerDirect(1 To i + 1)
            vGrabDimmerDirect(i + 1) = "send.modbus $dev=" & thisIOdev & " $reg=" & thisChReg & " $len=" & ioLen
        Next i
    Else
        vGrabDimmerDirect(1) = "nop"
        'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid dimmer instruction]"
    End If
    GrabDimmerDirect = vGrabDimmerDirect
End Function

Private Function isADimmerReadDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "read.dim") = 1 Then
        isADimmerReadDirect = True
    Else
        isADimmerReadDirect = False
    End If
End Function

Private Function GrabDimmerReadDirect(inputString As String, thisRCU As String, thisLine As String) As Variant  'validation error and warn added in 5.1
    Dim ioString As String, chString As String, LArray() As String, LArray1() As String, LArray2() As String
    Dim thisIOdev As String, thisChReg As Integer, ioLen As Integer
    
    thisIOdev = "ioexp1"
    ioLen = 2
    LArray = Split(inputString, " $")
    
    For i = LBound(LArray) To UBound(LArray)
        If InStr(LArray(i), "dev=") = 1 Then
            ioString = Replace(LArray(i), "dev=", "")
        ElseIf InStr(LArray(i), "ch=") = 1 Then
            chString = Replace(LArray(i), "ch=", "")
        End If
    Next i
    
    Dim vGrabDimmerReadDirect As Variant
    ReDim vGrabDimmerReadDirect(1 To 1)
    
    If Len(ioString) > 0 And Len(chString) > 0 Then
    
        LArray1 = Split(ioString, ",")
        LArray2 = Split(X1.CommaExpand(chString, 1, 16), ",")
        
        If Len(LArray1(0)) > 1 And IsNumeric(Right(LArray1(0), 1)) Then
            thisIOdev = LArray1(0)
        ElseIf Len(LArray1(0)) > 1 And Not IsNumeric(Right(LArray1(0), 1)) Then
            thisIOdev = LArray1(0) & "1"
        ElseIf Len(LArray1(0)) = 1 And IsNumeric(Right(LArray1(0), 1)) Then
            thisIOdev = Mid(thisIOdev, 1, Len(thisIOdev) - 1) & LArray1(0)
        End If
        
        If IsNumeric(LArray2(0)) Then
            If CInt(LArray2(0)) > 0 And CInt(LArray2(0)) < 9 Then
                thisChReg = CInt(LArray2(0))
            Else
                thisChReg = 1
            End If
        Else
            thisChReg = 1
        End If
        
        If InStr(thisIOdev, "ioe") = 1 Then
            thisChReg = 15 + 4 * (thisChReg - 1)
        Else
            thisChReg = 6 + 4 * (thisChReg - 1)
        End If
        
        vGrabDimmerReadDirect(1) = "read.modbus $dev=" & thisIOdev & " $reg=" & thisChReg & " $len=" & ioLen
        
        If (UBound(LArray1) - LBound(LArray1) + 1) > 1 Then
            X1.addWarn "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Reading from muntiple devices is not practical]"
        End If
        If (UBound(LArray2) - LBound(LArray2) + 1) > 1 Then
            X1.addWarn "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Reading from muntiple channels is not practical]"
        End If
        
    Else
        vGrabDimmerReadDirect(1) = "nop"
        'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid dimmer instruction]"
    End If
    GrabDimmerReadDirect = vGrabDimmerReadDirect
End Function


'other var direct

Private Function isOtherVarDirect(inputString As String, pp As Application, thisRCU As String) As Boolean
    
    If InStr(inputString, ".") > 0 Then
        Dim VarArray() As String, tempS As String
        VarArray = Split(inputString, ".")
        tempS = "load $var=" & VarArray(0)
        If X1.hasValidVar(tempS) Or X1.hasValidC_Var(tempS, pp, thisRCU) Or X1.hasValidM_Var(tempS, pp, thisRCU) Then
            isOtherVarDirect = True
        Else
            isOtherVarDirect = False
        End If
    Else
        isOtherVarDirect = False
    End If
End Function

Private Function GrabOtherVarDirect(inputString As String, pp As Application, thisRCU As String) As Variant
    
    Dim ioString As String, LArray() As String, LArray1() As String
    Dim OtherVar As String, L As Integer, Smode As String
    LArray = Split(Replace(inputString, "#", ""), " $")
    LArray1 = Split(LArray(0), ".")
    OtherVar = LArray1(0)
    Smode = ""
    
    Dim vGrabOtherVarDirect As Variant
    ReDim vGrabOtherVarDirect(1 To 1)
    
    If X1.hasValidM_Var("load $var=" & OtherVar, pp, thisRCU) Then
        Smode = " modbus"
    End If
    
    If LArray1(1) = "set" Then
        'If hasValidVar(inputString) Then
            'vGrabOtherVarDirect(1) = "load $var=" & Grab_variant(inputString, "var")
        'ElseIf hasValidC_Var(inputString, pp, rr) Then
            'vGrabOtherVarDirect(1) = "load $var=" & Grab_variant(inputString, "var")
        'Else
        If X1.hasValidM_Var(inputString, pp, thisRCU) Then
            vGrabOtherVarDirect(1) = "load modbus $var=" & Grab_variant(inputString, "var")
        Else
            vGrabOtherVarDirect(1) = "load $" & Grab_value_0(inputString, 1)
        End If
        ReDim Preserve vGrabOtherVarDirect(1 To 2)
        vGrabOtherVarDirect(2) = "save" & Smode & " $var=" & OtherVar
    ElseIf LArray1(1) = "unset" Then
        vGrabOtherVarDirect(1) = "load $hex=0"
        ReDim Preserve vGrabOtherVarDirect(1 To 2)
        vGrabOtherVarDirect(2) = "save" & Smode & " $var=" & OtherVar
    Else
        vGrabOtherVarDirect(1) = "load" & Smode & " $var=" & OtherVar
        For L = 1 To UBound(LArray1)
            'ReDim Preserve vGrabOtherVarDirect(1 To L + 1)
            If Grab_value_0(inputString, L) <> "NULL" Then
                ReDim Preserve vGrabOtherVarDirect(1 To L + 1)
                vGrabOtherVarDirect(L + 1) = "math." & LArray1(L) & " $" & Grab_value_0(inputString, L)
            End If
            'Else
                'vGrabOtherVarDirect(L + 1) = "math." & LArray1(L) & " $var=" & Grab_variant(inputString, "var")
            'End If
        Next L
        ReDim Preserve vGrabOtherVarDirect(1 To UBound(vGrabOtherVarDirect) + 1)
        vGrabOtherVarDirect(UBound(vGrabOtherVarDirect)) = "save" & Smode & " $var=" & OtherVar
    End If
    
    GrabOtherVarDirect = vGrabOtherVarDirect
End Function

'dali direct
Private Function isADaliDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "dali") = 1 Then
        isADaliDirect = True
    Else
        isADaliDirect = False
    End If
End Function

Private Function GrabDaliDirect(inputString As String, pp As Application, thisRCU As String, thisLine As String) As Variant
    Dim ioString As String, LArray1() As String, ioName As String, L As Integer
    Dim Dalibus As String, ioBlst As String, ioGp As String
    Dim idFactor As Integer, bgLum As String, bgID As String
    
    ioString = Replace(Replace(inputString, " ", ""), "#", "$")
    ioString = Mid(ioString, 1, InStr(ioString, "$") - 1)
    Dalibus = Grab_variant(inputString, "bus")
    ioBlst = Grab_variant(inputString, "ballast")
    ioGp = Grab_variant(inputString, "group")
    idFactor = 0
    bgLum = Grab_value_5(inputString, pp)
    
    If Mid(inputString, InStr(inputString, "dali") + 4, 1) = "." Then
        ioName = "dali1"
    Else
        ioName = "dali" & Mid(inputString, InStr(inputString, "dali") + 4, 1)
    End If
    
    If Dalibus <> "NULL" And IsNumeric(Dalibus) Then
        If CInt(Dalibus) = 2 Then
            Dalibus = "05"
        Else
            Dalibus = "06"
        End If
    Else
        Dalibus = "06"
    End If
    
    If InStr(ioString, ".up") > 0 Or InStr(ioString, ".on") > 0 Then
        idFactor = 1
        bgLum = "05"
    ElseIf InStr(ioString, ".down") > 0 Or InStr(ioString, ".off") > 0 Then
        idFactor = 1
        bgLum = "00"
    ElseIf InStr(ioString, ".stop") > 0 Then
        idFactor = 0
        bgLum = "FF"
    Else
        idFactor = 0
        bgLum = Grab_value_5(inputString, pp)
    End If
    
    Dim vGrabDaliDirect As Variant
    ReDim vGrabDaliDirect(1 To 1)
    
    If ioBlst <> "NULL" Or ioGp <> "NULL" Then
        If ioBlst <> "NULL" Then
            ioBlst = X1.CommaExpand(ioBlst, 0, 63)
            LArray1 = Split(ioBlst, ",")
            For i = LBound(LArray1) To UBound(LArray1)
                If IsNumeric(LArray1(i)) Then
                    If CInt(LArray1(i)) > -1 And CInt(LArray1(i)) < 64 Then
                        ReDim Preserve vGrabDaliDirect(1 To UBound(vGrabDaliDirect) + 1)
                        bgID = pp.WorksheetFunction.Dec2Hex(CInt(LArray1(i)) * 2 + idFactor, 2)
                        vGrabDaliDirect(UBound(vGrabDaliDirect)) = "send.modbus $dev=" & ioName & " $reg=" & Dalibus & " $hex=" & bgID & bgLum
                    End If
                End If
            Next i
        End If
        
        If ioGp <> "NULL" Then
            ioGp = X1.CommaExpand(ioGp, 0, 15)
            LArray1 = Split(ioGp, ",")
            For i = LBound(LArray1) To UBound(LArray1)
                If IsNumeric(LArray1(i)) Then
                    If CInt(LArray1(i)) > -1 And CInt(LArray1(i)) < 16 Then
                        ReDim Preserve vGrabDaliDirect(1 To UBound(vGrabDaliDirect) + 1)
                        bgID = pp.WorksheetFunction.Dec2Hex(CInt(LArray1(i)) * 2 + idFactor + 128, 2)
                        vGrabDaliDirect(UBound(vGrabDaliDirect)) = "send.modbus $dev=" & ioName & " $reg=" & Dalibus & " $hex=" & bgID & bgLum
                    End If
                End If
            Next i
        End If
    Else
        'MsgBox """" & inputString & """ is an invalid DALI instruction" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]"
        vGrabDaliDirect(1) = "nop"
    End If
    
    GrabDaliDirect = vGrabDaliDirect
End Function

'dmx
Private Function isADMXDirect(inputString As String) As Boolean
    If InStr(Replace(LCase(inputString), " ", ""), "dmx.") = 1 Then
        isADMXDirect = True
    Else
        isADMXDirect = False
    End If
End Function

Private Function GrabDMXDirect(inputString As String, pp As Application, thisRCU As String, thisLine As String) As Variant
    Dim ioChString As String, ioChExp As String, LArray1() As String
    
    ioChString = Grab_variant(inputString, "ch")
    
    Dim vGrabDMXDirect As Variant
    ReDim vGrabDMXDirect(1 To 1)
    
    If ioChString <> "NULL" Then
        ioChExp = X1.CommaExpand(ioChString, 0, 63)
        LArray1 = Split(ioChExp, ",")
        
        For i = LBound(LArray1) To UBound(LArray1)
            If IsNumeric(LArray1(i)) Then
                If CInt(LArray1(i)) > -1 And CInt(LArray1(i)) < 64 Then
                    ReDim Preserve vGrabDMXDirect(1 To UBound(vGrabDMXDirect) + 1)
                    vGrabDMXDirect(UBound(vGrabDMXDirect)) = Replace(inputString, "$ch=" & ioChString, "$ch=" & LArray1(i))
                End If
            End If
        Next i
        
    Else
        'MsgBox """" & inputString & """ is an invalid DMX instruction" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DMX instruction]"
        vGrabDMXDirect(1) = "nop"
    End If
    
    GrabDMXDirect = vGrabDMXDirect
End Function


'goto, gotowait, 03 01, 03 02, 03 03, if conditions
Private Function FunctionConvert(inputString As String, inputLine As Integer, pp As Application, rr As String) As String
    Dim LArray() As String, ConditionF As String
    Dim thisFunction As String, cFunctionConvert As String, LineJump As String, damper As String
    Dim withVar As Boolean

    cFunctionConvert = "00 00": withVar = False: damper = "00"
    ConditionF = ConditionSet(inputString, pp, rr)
    withVar = X1.hasValidVar(inputString) Or X1.hasValidC_Var(inputString, pp, rr)
    LArray = Split(inputString, " ")
    
    For i = LBound(LArray) To UBound(LArray)
        If InStr(LArray(i), "#") = 1 Then
            thisFunction = Replace(LArray(i), "#", "")
            LineJump = X1.getLineFromFunctionName(thisFunction, inputLine, pp, rr) 'input line for nearest exit
            cFunctionConvert = Right(LineJump, 2) & " " & Left(LineJump, 2)
        End If
    Next i
    
    
    If cFunctionConvert = "00 00" Then
        'MsgBox """" & inputString & """ Invalid function call, function definition not found" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & inputLine & " | " & inputString & "][Invalid function call]"
        FunctionConvert = "01 01 00 00 00 00 00 00;"
    Else
        If ConditionF = "01 01" Or ConditionF = "03 01" Or ConditionF = "03 02" Then
            FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Grab_Time_1(inputString, pp) & ";"
        ElseIf ConditionF = "03 03" Then
            FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Left(Grab_Time_1(inputString, pp), 8) & " " & Grab_Dev_1(inputString, pp, rr) & ";"
        Else
            If withVar Then
                ConditionF = Left(ConditionF, 2) & " 0A"
                
                If ConditionF = "5c 0A" Then  '5c damper "02" in the end
                    damper = Grab_variant(inputString, "milisec")
                    
                    If damper <> "NULL" And IsNumeric(damper) Then
                        If CInt(damper) > 0 And CInt(damper) < 256 Then
                            damper = pp.WorksheetFunction.Dec2Hex(damper, 2)
                        Else
                            damper = "02"
                        End If
                    Else
                        damper = "02"
                    End If
                End If
                
                If X1.hasValidVar(inputString) And (Not X1.hasValidC_Var(inputString, pp, rr)) Then
                    FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Grab_Var_1(inputString) & " 00 00 " & damper & ";"
                ElseIf (Not X1.hasValidVar(inputString)) And X1.hasValidC_Var(inputString, pp, rr) Then
                    FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Grab_Var_2(inputString, pp, rr) & " 00 00 " & damper & ";"
                End If
            Else
                FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Grab_value_1(inputString, "OR", pp) & ";"
            End If
        End If
    End If
End Function

Private Function ConditionSet(inputString As String, pp As Application, rr As String) As String
    
    If InStr(inputString, "if ") = 1 Then
        If (InStr(inputString, "bit ") > 0 Or InStr(inputString, "bitwise ") > 0) And (InStr(inputString, "true ") > 0 Or InStr(inputString, "check ") > 0) Then
            ConditionSet = "09 01"
        ElseIf (InStr(inputString, "bit ") > 0 Or InStr(inputString, "bitwise ") > 0) And InStr(inputString, "run ") > 0 Then
            ConditionSet = "60 01"
        ElseIf InStr(inputString, "not equal ") > 0 Or InStr(inputString, "not equals to ") > 0 Or InStr(Replace(inputString, " ", ""), "if<>") = 1 Or InStr(Replace(inputString, " ", ""), "if!=") = 1 Then
            ConditionSet = "05 01"
        ElseIf InStr(inputString, "equal ") > 0 Or InStr(inputString, "equals to ") > 0 Or InStr(Replace(inputString, " ", ""), "if=") = 1 Then
            ConditionSet = "08 01"
        ElseIf InStr(inputString, "greater ") > 0 Or InStr(inputString, "greater than ") > 0 Or InStr(Replace(inputString, " ", ""), "if>") = 1 Then
            ConditionSet = "06 01"
        ElseIf InStr(inputString, "lesser ") > 0 Or InStr(inputString, "less than ") > 0 Or InStr(Replace(inputString, " ", ""), "if<") = 1 Then
            ConditionSet = "07 01"
        Else
            ConditionSet = "01 01"
        End If
    ElseIf InStr(inputString, "dimflow.") = 1 Then
        If InStr(inputString, ".oncheck ") > 0 Or InStr(inputString, ".checkon ") > 0 Then
            ConditionSet = "5b 0A"
        ElseIf InStr(inputString, ".offcheck ") > 0 Or InStr(inputString, ".checkoff ") > 0 Then
            ConditionSet = "5c 0A"
        Else
            ConditionSet = "01 01"
        End If
    ElseIf InStr(inputString, "run ") = 1 Or InStr(inputString, "reuse ") = 1 Or InStr(inputString, "re-run ") = 1 Or InStr(inputString, "rerun ") = 1 Then
        If InStr(inputString, "reuse ") = 1 Or InStr(inputString, "re-run ") = 1 Or InStr(inputString, "rerun ") = 1 Then
            If pp.WorksheetFunction.Hex2Dec(Grab_Dev_1(inputString, pp, rr)) > 0 Then
                ConditionSet = "03 03"
            Else
                ConditionSet = "03 02"
            End If
        Else
            ConditionSet = "03 01"
        End If
    Else
        ConditionSet = "01 01"
    End If
End Function

'wait (no goto)
Private Function WaitConvert(inputString As String, pp As Application) As String
    WaitConvert = "01 01 00 00 " & Grab_Time_1(inputString, pp) & ";"
End Function

'wait abs

Private Function WaitABSConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    If X1.hasValidVar(inputString) Then
        WaitABSConvert = "0A 0A 00 00 " & Grab_Var_1(inputString) & " 00 00 00;"
    ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
        WaitABSConvert = "0A 0A 00 00 " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
    ElseIf hasTime(inputString) Then
        WaitABSConvert = "0A 01 00 00 " & Grab_Time_1(inputString, pp) & ";"
    Else
        X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid Wait ABS instruction]"
        WaitABSConvert = "01 01 00 00 00 00 00 00;"
    End If
End Function

'exit handling
Private Function isAnExit(inputString As String) As Boolean
    If InStr(inputString, "exit") = 1 Or InStr(inputString, "quit") = 1 Or InStr(inputString, "end") = 1 Then
        isAnExit = True
    Else
        isAnExit = False
    End If
End Function

Private Function ExitConvert() As String
    ExitConvert = "04 01 00 00 00 00 00 00;"
End Function

'debug handling
Private Function isDebug(inputString As String) As Boolean
    If InStr(inputString, "debug") = 1 Then
        isDebug = True
    Else
        isDebug = False
    End If
End Function

Private Function DebugConvert() As String
    DebugConvert = "40 01 01 00 00 00 00 00;"
End Function

'version handling
Private Function isVersion(inputString As String) As Boolean
    If InStr(inputString, "version") = 1 Then
        isVersion = True
    Else
        isVersion = False
    End If
End Function

Private Function VersionConvert(inputString As String) As String
    Dim buildVersion As Integer, vData() As String
    vData = Split(Mid(inputString, InStr(inputString, "[") + 1, InStr(inputString, "]") - InStr(inputString, "[") - 1), ",")
    buildVersion = CInt(vData(1))
    VersionConvert = "00 00 00 00 " & WorksheetFunction.Dec2Hex(CDbl(X1.getThisXCODE_version) * 10, 2) & " " & X1.buildDate_hex & " " & WorksheetFunction.Dec2Hex(buildVersion, 2) & ";"
End Function

'math handling
Private Function isAMath(inputString As String) As Boolean
    If InStr(inputString, "math.") = 1 Then
        isAMath = True
    Else
        isAMath = False
    End If
End Function

Private Function MathConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    Dim operator As String, Smode As String
    operator = "80": Smode = "00"
    
    If InStr(inputString, "shift ") > 1 Then
        Smode = "01"
    End If
    
    If InStr(inputString, "math.add ") = 1 Or InStr(inputString, "math.+ ") = 1 Then
        operator = "01"
    ElseIf InStr(inputString, "math.sub ") = 1 Or InStr(inputString, "math.- ") = 1 Then
        operator = "02"
    ElseIf InStr(inputString, "math.mul ") = 1 Or InStr(inputString, "math.* ") = 1 Or InStr(inputString, "math.x ") = 1 Then
        operator = "03"
    ElseIf InStr(inputString, "math.or ") = 1 Or InStr(inputString, "math.|| ") = 1 Then
        operator = "04"
    ElseIf InStr(inputString, "math.xor ") = 1 Or InStr(inputString, "math.x| ") = 1 Then
        operator = "05"
    ElseIf InStr(inputString, "math.and ") = 1 Or InStr(inputString, "math.&& ") = 1 Then
        operator = "06"
    ElseIf InStr(inputString, "math.shiftl ") = 1 Or InStr(inputString, "math.<< ") = 1 Then
        operator = "07"
    ElseIf InStr(inputString, "math.shiftr ") = 1 Or InStr(inputString, "math.>> ") = 1 Then
        operator = "08"
    ElseIf InStr(inputString, "math.div ") = 1 Or InStr(inputString, "math./ ") = 1 Then
        operator = "09"
    ElseIf InStr(inputString, "math.mod ") = 1 Or InStr(inputString, "math.% ") = 1 Then
        operator = "0A"
    End If
    
    If operator = "80" Then
        'not math function
        'MsgBox """" & inputString & """ is an invalid math function" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid math instruction]"
        MathConvert = "01 01 00 00 00 00 00 00;"
    Else
        If X1.hasValidVar(inputString) Then
            MathConvert = "10 " & operator & " 0A " & Smode & " " & Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
            MathConvert = "10 " & operator & " 0A " & Smode & " " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
        Else
            If InStr(inputString, "$var=") > 4 Then
                'MsgBox "Variable """ & Grab_variant(inputString, "var") & """ not found" 'replaced with errorlog ver 5.0
                X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & Grab_variant(inputString, "var") & "]"
                MathConvert = "01 01 00 00 00 00 00 00;"
            Else
                If operator = "06" Then
                    MathConvert = "10 " & operator & " 01 " & Smode & " " & Grab_value_1(inputString, "AND", pp) & ";"
                Else
                    MathConvert = "10 " & operator & " 01 " & Smode & " " & Grab_value_1(inputString, "OR", pp) & ";"
                End If
            End If
        End If
    End If
End Function

'RCU handling
Private Function isARCU(inputString As String) As Boolean
    If InStr(inputString, "rcu.") = 1 Then
        isARCU = True
    Else
        isARCU = False
    End If
End Function

Private Function RCUConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    Dim operator As String, Smode As String
    operator = "09": Smode = "01"
    
    If InStr(inputString, " acc") > 1 Then
        Smode = "00"
    End If
    
    If InStr(inputString, "rcu.read input") = 1 Then 'removed in 4.3 <Or InStr(inputString, "rcu.input") = 1>
        operator = "08"
    ElseIf InStr(inputString, "rcu.read mosfet") = 1 Then
        operator = "0C"
    ElseIf InStr(inputString, "rcu.read batt") = 1 Then
        operator = "0E"
    ElseIf InStr(inputString, "rcu.read es") = 1 Then
        operator = "H1"
    ElseIf InStr(inputString, "rcu.read aout2") = 1 Or InStr(inputString, "rcu.read analog2") = 1 Then
        operator = "H3"
    ElseIf InStr(inputString, "rcu.read aout") = 1 Or InStr(inputString, "rcu.read analog") = 1 Then
        operator = "H2"
    ElseIf InStr(inputString, "rcu.read") = 1 Then
        operator = "03"
    ElseIf InStr(inputString, "rcu.xor ") = 1 Then
        operator = "0B"
    ElseIf InStr(inputString, "rcu.set batt") = 1 Then
        operator = "0D"
    ElseIf InStr(inputString, "rcu.unset batt") = 1 Then
        operator = "0D"
    ElseIf InStr(inputString, "rcu.set ") = 1 And InStr(inputString, " all") > 0 Then
        operator = "04"
    ElseIf (InStr(inputString, "rcu.set ") = 1 Or InStr(inputString, "rcu.unset ") = 1) And InStr(inputString, " es") > 0 Then
        operator = "05"
    ElseIf (InStr(inputString, "rcu.set ") = 1 Or InStr(inputString, "rcu.unset ") = 1) And (InStr(inputString, " analog2") > 0 Or InStr(inputString, " aout2") > 0) Then
        operator = "07"
    ElseIf (InStr(inputString, "rcu.set ") = 1 Or InStr(inputString, "rcu.unset ") = 1) And (InStr(inputString, " analog") > 0 Or InStr(inputString, " aout") > 0) Then
        operator = "06"
    ElseIf InStr(inputString, "rcu.set ") = 1 Or InStr(inputString, "rcu.or ") = 1 Then
        operator = "09"
    ElseIf InStr(inputString, "rcu.unset ") = 1 Or InStr(inputString, "rcu.and ") = 1 Then
        operator = "0A"
    ElseIf InStr(inputString, "rcu.pwm ") = 1 Then
        operator = "00"
    ElseIf InStr(inputString, "rcu.dim es") = 1 Then
        operator = "G0"
    ElseIf InStr(inputString, "rcu.dim aout2") = 1 Or InStr(inputString, "rcu.dim analog2") = 1 Then
        operator = "G2"
    ElseIf InStr(inputString, "rcu.dim aout") = 1 Or InStr(inputString, "rcu.dim analog") = 1 Then
        operator = "G1"
    ElseIf InStr(inputString, "rcu.stop es") = 1 Then
        operator = "I0"
    ElseIf InStr(inputString, "rcu.stop aout2") = 1 Or InStr(inputString, "rcu.stop analog2") = 1 Then
        operator = "I2"
    ElseIf InStr(inputString, "rcu.stop aout") = 1 Or InStr(inputString, "rcu.stop analog") = 1 Then
        operator = "I1"
    End If
    
    If operator = "03" Or operator = "08" Or operator = "0C" Or operator = "0E" Then
        RCUConvert = "02 " & operator & " 00 00 00 00 00 00;"
    ElseIf operator = "H1" Or operator = "H2" Or operator = "H3" Then
        RCUConvert = "53 0" & Right(operator, 1) & " 00 00 00 00 00 00;"
    ElseIf X1.hasValidVar(inputString) Then
        If Left(operator, 1) = "G" Or Left(operator, 1) = "I" Then
            RCUConvert = "5" & Right(operator, 1) & " 0A 00 00 " & Grab_Var_1(inputString) & " 00 00 00;"
        Else
            RCUConvert = "02 " & operator & " 0A 00 " & Grab_Var_1(inputString) & " 00 00 00;"
        End If
    ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
        If Left(operator, 1) = "G" Or Left(operator, 1) = "I" Then
            RCUConvert = "5" & Right(operator, 1) & " 0A 00 00 " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
        Else
            RCUConvert = "02 " & operator & " 0A 00 " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
        End If
    Else
        If Smode = "00" Then
            If Left(operator, 1) = "G" Or Left(operator, 1) = "I" Then
                RCUConvert = "5" & Right(operator, 1) & " 00 00 00 00 00 00 00;"
            Else
                RCUConvert = "02 " & operator & " 00 00 00 00 00 00;"
            End If
        Else
            If InStr(inputString, "$var=") > 4 Then
                'MsgBox "Variable """ & Grab_variant(inputString, "var") & """ not found" 'replaced with errorlog ver 5.0
                X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & Grab_variant(inputString, "var") & "]"
                RCUConvert = "01 01 00 00 00 00 00 00;"
            Else
                If operator = "05" Or operator = "06" Or operator = "07" Then
                    If InStr(inputString, "rcu.unset ") = 1 Then
                        RCUConvert = "02 " & operator & " 01 00 00 00 00 00;"
                    Else
                        RCUConvert = "02 " & operator & " 01 00 " & Grab_value_4(inputString, pp) & " 00 00 00;"
                    End If
                ElseIf operator = "0D" Then
                    If InStr(inputString, "rcu.unset") = 1 Then
                        RCUConvert = "02 " & operator & " 01 00 00 00 00 00;"
                    ElseIf InStr(inputString, "rcu.set") = 1 And InStr(inputString, " $") < 1 Then
                        RCUConvert = "02 " & operator & " 01 00 01 00 00 00;"
                    Else
                        RCUConvert = "02 " & operator & " 01 00 " & Grab_value_4(inputString, pp) & " 00 00 00;"
                    End If
                ElseIf Left(operator, 1) = "G" Or Left(operator, 1) = "I" Then
                    RCUConvert = "5" & Right(operator, 1) & " 01 00 00 " & Grab_Time_3(inputString, pp) & " " & Grab_value_4(inputString, pp) & ";"
                Else
                    RCUConvert = "02 " & operator & " 01 00 " & Grab_value_1(inputString, "OR", pp) & ";"
                End If
            End If
        End If
    End If
End Function

'send handling

Private Function isASend(inputString As String) As Boolean
    If InStr(inputString, "send.") = 1 Then
        isASend = True
    Else
        isASend = False
    End If
End Function

Private Function SendConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    If (InStr(inputString, "send.mod ") = 1 Or InStr(inputString, "send.modbus ") = 1) And (InStr(inputString, " $dev=") > 5 Or InStr(inputString, " $type=") > 5) And InStr(inputString, " $reg=") > 5 Then
        'If pp.WorksheetFunction.Hex2Dec(Grab_Dev_1(inputString, pp, rr)) = 255 Then //XCODE 4.7
            ''20 & 22
            'If InStr(inputString, "$len=") Then
                'SendConvert = "22 " & Grab_Dev_1(inputString, pp, rr) & " " & Grab_Reg_1(inputString) & " " & Grab_value_1(inputString, "OR", pp) & ";"
            'Else
                'SendConvert = "20 " & Grab_Dev_1(inputString, pp, rr) & " " & Grab_Reg_1(inputString) & " " & Grab_value_3(inputString, "OR", pp) & " 00 00;"
            'End If
        'Else/
        If pp.WorksheetFunction.Hex2Dec(Grab_Dev_1(inputString, pp, rr)) > 0 Then 'And Grab_Reg_1(inputString) <> "00 00" Then
        '20 & 22
            If InStr(inputString, "$len=") Then
                SendConvert = "22 " & Grab_Dev_1(inputString, pp, rr) & " " & Grab_Reg_1(inputString) & " " & Grab_value_1(inputString, "OR", pp) & ";"
            Else
                SendConvert = "20 " & Grab_Dev_1(inputString, pp, rr) & " " & Grab_Reg_1(inputString) & " " & Grab_value_3(inputString, "OR", pp) & " 00 00;"
            End If
        ElseIf pp.WorksheetFunction.Hex2Dec(Grab_Dev_Type_1(inputString, pp)) > 0 Then 'And Grab_Reg_1(inputString) <> "00 00" Then
        '21 & 23
            If InStr(inputString, "$len=") Then
                SendConvert = "23 " & Grab_Dev_Type_1(inputString, pp) & " " & Grab_Reg_1(inputString) & " " & Grab_value_1(inputString, "OR", pp) & ";"
            Else
                SendConvert = "21 " & Grab_Dev_Type_1(inputString, pp) & " " & Grab_Reg_1(inputString) & " " & Grab_value_3(inputString, "OR", pp) & " 00 00;"
            End If
        Else
            SendConvert = "01 01 00 00 00 00 00 00;"
        End If
    Else
        'MsgBox """" & inputString & """ Invalid modbus instruction" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid modbus instruction]"
        SendConvert = "01 01 00 00 00 00 00 00;"
    End If
End Function

'read handling

Private Function isARead(inputString As String) As Boolean
    If InStr(inputString, "read.") = 1 Then
        isARead = True
    Else
        isARead = False
    End If
End Function

Private Function ReadConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    If InStr(inputString, "read.mod ") = 1 Or InStr(inputString, "read.modbus ") = 1 Then
        If pp.WorksheetFunction.Hex2Dec(Grab_Dev_1(inputString, pp, rr)) > 0 Then ' And Grab_Reg_1(inputString) <> "00 00" Then //reg 0 allowed
        '20 & 22
            If InStr(inputString, "$len=") Then
                ReadConvert = "26 " & Grab_Dev_1(inputString, pp, rr) & " " & Grab_Reg_1(inputString) & " " & Grab_value_1(inputString, "OR", pp) & ";"
            Else
                ReadConvert = "26 " & Grab_Dev_1(inputString, pp, rr) & " " & Grab_Reg_1(inputString) & " 02 00 00 00;"
            End If
        Else
            If pp.WorksheetFunction.Hex2Dec(Grab_Dev_1(inputString, pp, rr)) = 0 Then
                'MsgBox """" & inputString & """ Invalid device" 'replaced with errorlog ver 5.0
                X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid device]"
            End If
            
            ReadConvert = "01 01 00 00 00 00 00 00;"
        End If
    Else
        'MsgBox """" & inputString & """ Invalid modbus instruction" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid modbus instruction]"
        ReadConvert = "01 01 00 00 00 00 00 00;"
    End If
End Function

' load handling
Private Function isALoad(inputString As String) As Boolean
    If InStr(inputString, "load ") = 1 Then
        isALoad = True
    Else
        isALoad = False
    End If
End Function

Private Function LoadConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    Dim Smode As String
    Smode = "00"
    If InStr(inputString, "shift ") > 1 Then
        Smode = "01"
    End If
    
    If InStr(inputString, " set") > 1 Then
        Dim thisLen As String
        'added in v5.3, checking first to avoid duplicate detection in variable space
        thisLen = Grab_variant(inputString, "len")
        
        If Not thisLen = "NULL" Then
            If X1.isValidHexString(thisLen) Then
                thisLen = Right("00" & pp.WorksheetFunction.Hex2Dec(thisLen), 2)
            Else
                X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Incompatible length | " & Grab_variant(inputString, "len") & "]"
                thisLen = "01"
            End If
        Else
            thisLen = "01"
        End If
        
        If thisLen = "04" Then
            LoadConvert = "10 80 18 " & Smode & " " & Grab_Var_4(inputString) & ";"
        ElseIf thisLen = "02" Then
            LoadConvert = "10 80 17 " & Smode & " " & Grab_Var_4(inputString) & ";"
        Else
            LoadConvert = "10 80 16 " & Smode & " " & Grab_Var_4(inputString) & ";"
        End If
    ElseIf X1.hasValidVar(inputString) Then
        LoadConvert = "10 80 0A " & Smode & " " & Grab_Var_1(inputString) & " 00 00 00;"
    ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
        LoadConvert = "10 80 0A " & Smode & " " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
    ElseIf X1.hasValidM_Var(inputString, pp, rr) And InStr(inputString, " mod") > 1 And InStr(inputString, " flo") > 1 Then 'added in ver 5.3
        LoadConvert = "10 80 15 " & Smode & " " & Grab_Var_3(inputString, pp, rr) & " 00 00 00;"
    ElseIf X1.hasValidM_Var(inputString, pp, rr) And InStr(inputString, " mod") > 1 Then
        LoadConvert = "10 80 12 " & Smode & " " & Grab_Var_3(inputString, pp, rr) & " 00 00 00;"
    ElseIf InStr(inputString, " cbs") > 1 Then
        LoadConvert = "10 85 01 " & Smode & " " & Grab_element(inputString) & ";"
    ElseIf InStr(inputString, " enum") > 1 Then
        LoadConvert = "10 80 10 " & Smode & " 00 00 00 00;"
    ElseIf InStr(inputString, " time") > 1 Then
        If InStr(inputString, " unix") > 1 Or InStr(inputString, " utc") > 1 Then
            LoadConvert = "10 84 00 " & Smode & " 00 00 00 00;"
        ElseIf InStr(inputString, " now") > 1 Then
            LoadConvert = "10 80 01 " & Smode & " " & Give_Time_1 & ";"
        Else
            LoadConvert = "10 82 00 " & Smode & " 00 00 00 00;"
        End If
    Else
        If InStr(inputString, "$var=") > 4 Then
            'MsgBox "Variable """ & Grab_variant(inputString, "var") & """ not found" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & Grab_variant(inputString, "var") & "]"
            LoadConvert = "01 01 00 00 00 00 00 00;"
        Else
            LoadConvert = "10 80 01 " & Smode & " " & Grab_value_1(inputString, "OR", pp) & ";"
        End If
    End If
End Function

Private Function isADimmerLoad(inputString As String) As Boolean
    If InStr(inputString, "load dimmer ") = 1 Then
        isADimmerLoad = True
    Else
        isADimmerLoad = False
    End If
End Function

Private Function LoadDimmerConvert(inputString As String, pp As Application) As String
    LoadDimmerConvert = "10 80 01 01 " & Grab_value_2(inputString, pp) & " " & Grab_Time_2(inputString, pp) & ";"
End Function

'save handling
Private Function isASave(inputString As String) As Boolean
    If InStr(inputString, "save ") = 1 Then
        isASave = True
    Else
        isASave = False
    End If
End Function

Private Function SaveConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    Dim Smode As String
    Smode = "00"
    If InStr(inputString, "shift ") > 1 Then
        Smode = "01"
    End If
    
    If InStr(inputString, " set") > 1 Then
        Dim thisLen As String
        'added in v5.3, checking first to avoid duplicate detection in variable space
        thisLen = Grab_variant(inputString, "len")
        
        If Not thisLen = "NULL" Then
            If X1.isValidHexString(thisLen) Then
                thisLen = Right("00" & pp.WorksheetFunction.Hex2Dec(thisLen), 2)
            Else
                X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Incompatible length | " & Grab_variant(inputString, "len") & "]"
                thisLen = "01"
            End If
        Else
            thisLen = "01"
        End If
        
        If thisLen = "04" Then
            SaveConvert = "10 81 18 " & Smode & " " & Grab_Var_4(inputString) & ";"
        ElseIf thisLen = "02" Then
            SaveConvert = "10 81 17 " & Smode & " " & Grab_Var_4(inputString) & ";"
        Else
            SaveConvert = "10 81 16 " & Smode & " " & Grab_Var_4(inputString) & ";"
        End If
    ElseIf X1.hasValidVar(inputString) Then
        SaveConvert = "10 81 0A " & Smode & " " & Grab_Var_1(inputString) & " 00 00 00;"
    ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
        SaveConvert = "10 81 0A " & Smode & " " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
    ElseIf X1.hasValidM_Var(inputString, pp, rr) And InStr(inputString, " mod") > 1 And InStr(inputString, " flo") > 1 Then 'added in ver 5.3
        SaveConvert = "10 81 15 " & Smode & " " & Grab_Var_3(inputString, pp, rr) & " 00 00 00;"
    ElseIf X1.hasValidM_Var(inputString, pp, rr) And InStr(inputString, " mod") > 1 Then
        SaveConvert = "10 81 12 " & Smode & " " & Grab_Var_3(inputString, pp, rr) & " 00 00 00;"
    ElseIf InStr(inputString, " time") > 1 Then
        SaveConvert = "10 83 00 00 00 00 00 00;"
    Else
        If InStr(inputString, "$var") > 0 Then
            'MsgBox "Variable """ & Grab_variant(inputString, "var") & """ not found" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & Grab_variant(inputString, "var") & "]"
        Else
            'MsgBox "Invalid instruction """ & inputString & """" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid instruction]"
        End If
        SaveConvert = "01 01 00 00 00 00 00 00;"
    End If
End Function

'dimflow handling

Private Function isADimFlow(inputString As String) As Boolean
    If InStr(inputString, "dimflow.") = 1 Then
        isADimFlow = True
    Else
        isADimFlow = False
    End If
End Function

Private Function DimFlowConvert(inputString As String, pp As Application, rr As String, thisLine As String) As String
    Dim Dmode As String
    Dmode = "01 01"
    
    If InStr(inputString, ".start ") > 0 Or InStr(inputString, ".onstart ") > 0 Or InStr(inputString, ".starton ") > 0 Then
        Dmode = "5a 0A"
    ElseIf InStr(inputString, ".getdir") > 0 Or InStr(inputString, ".dirget ") > 0 Then
        Dmode = "5e 0A"
    ElseIf InStr(inputString, ".setdir") > 0 Or InStr(inputString, ".dirset ") > 0 Then
        Dmode = "5d 0A"
    End If
    
    If Dmode <> "01 01" Then
        If X1.hasValidVar(inputString) Then
            DimFlowConvert = Dmode & " 00 00 " & Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
            DimFlowConvert = Dmode & " 00 00 " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
        Else
            'MsgBox "Variable """ & Grab_variant(inputString, "var") & """ not found" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & Grab_variant(inputString, "var") & "]"
            DimFlowConvert = "01 01 00 00 00 00 00 00;"
        End If
    Else
        'MsgBox inputString & " is an invalid instruction" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & rr & " | XCODE" & rr & " | Program" & rr & " | line " & thisLine & " | " & inputString & "][Invalid instruction]"
        DimFlowConvert = "01 01 00 00 00 00 00 00;"
    End If
End Function

'DMX handling
Private Function isADMX(inputString As String) As Boolean
    If InStr(inputString, "dmx.") = 1 Then
        isADMX = True
    Else
        isADMX = False
    End If
End Function

Private Function DMXConvert(inputString As String, pp As Application, rr As String) As String
    Dim operator As String, Smode As String, ioCh As String
    operator = "01": Smode = "01"
    
    If InStr(inputString, " acc") > 1 Then
        Smode = "00"
    End If
    
    ioCh = Grab_variant(inputString, "ch")
    
    If InStr(inputString, "dmx.send ") = 1 Or InStr(inputString, "dmx.ramp ") = 1 Or InStr(inputString, "dmx.set ") = 1 Then
        operator = "01"
    ElseIf InStr(inputString, "dmx.stop ") = 1 Or InStr(inputString, "dmx.hold ") = 1 Then
        operator = "02"
    ElseIf InStr(inputString, "dmx.read ") = 1 Then
        operator = "03"
    End If
    
    If ioCh <> "NULL" Then
        ioCh = pp.WorksheetFunction.Dec2Hex(CInt(ioCh), 2)
        
        If operator = "03" Then
            DMXConvert = "61 02 00 " & ioCh & " 00 00 00 00;"
        ElseIf operator = "02" Then
            DMXConvert = "61 01 01 " & ioCh & " 00 00 00 00;"
        Else
            If Smode = "00" Then
                DMXConvert = "61 01 00 " & ioCh & " 00 00 00 00;"
            ElseIf X1.hasValidVar(inputString) Then
                DMXConvert = "61 01 0A " & ioCh & " " & Grab_Var_1(inputString) & " 00 00 00;"
            ElseIf X1.hasValidC_Var(inputString, pp, rr) Then
                DMXConvert = "61 01 0A " & ioCh & " " & Grab_Var_2(inputString, pp, rr) & " 00 00 00;"
            Else
                DMXConvert = "61 01 01 " & ioCh & " " & Grab_Time_3(inputString, pp) & " " & Grab_value_6(inputString, pp) & ";"
            End If
        End If
    Else
        DMXConvert = "01 01 00 00 00 00 00 00;"
    End If
End Function

'main grabber functions
'variable handling
'4.9 added room_empty_temp and room_empty_thermo. old room_empty var kept also for backword compatibility
Private Function Grab_Var_1(inputString As String) As String 'RCU var
    Dim V As String, LArray_T1() As String, P As String
    Dim RcuVariableList() As String, RcuNameList() As String
    RcuVariableList = Split("04,05,06,07,08,09,0A,0B,0C,0D,0E,0F,10,11,12,13,14,14,33,15,16,17,18,19,1A,1B,1C,1D,1E,1F,20,21,22,23,24,25,26,27", ",")
    RcuNameList = Split("anti_ice,checkin,thermo1,thermo2,thermo3,thermo4,thermo5,thermo6,temp1,temp2,temp3,temp4,temp5,temp6,season,intervention,room_empty,room_empty_temp,room_empty_thermo,user,workflow,gs1,gs2,gs3,gs4,gs5,gs6,gs7,gs8,gs9,gs10,ioexp1,ioexp2,ioexp3,ioexp4,ioexp5,ioexp6,tag_minmax", ",")
    V = "": P = ""
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
        If InStr(Replace(LArray_T1(i), " ", ""), "var=") Then
            V = Replace(LArray_T1(i), "var=", "")
            Exit For
        End If
    Next i
    For i = LBound(RcuVariableList) To UBound(RcuVariableList)
        If V = RcuVariableList(i) Or V = RcuNameList(i) Then
            P = RcuVariableList(i)
            Exit For
        End If
    Next i
    Grab_Var_1 = P
End Function

Private Function Grab_Var_2(inputString As String, pp As Application, rr As String) As String 'Custom var
    Dim V As String, LArray_T2() As String, P As String
    Dim rngVar As Range
    V = "": P = ""
    LArray_T2 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T2) To UBound(LArray_T2)
        If InStr(Replace(LArray_T2(i), " ", ""), "var=") Then
            V = Replace(LArray_T2(i), "var=", "")
            Exit For
        End If
    Next i
    For Each rngVar In pp.Range(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0), pp.Range("CustomVar" & rr & "!B3").End(xlDown))
         If V = pp.WorksheetFunction.Trim(LCase(rngVar.Text)) Or V = pp.WorksheetFunction.Trim(LCase(rngVar.Offset(0, 1).Text)) Then
            P = rngVar.Text
            Exit For
         End If
    Next rngVar
    Grab_Var_2 = Right("00" & P, 2)
End Function

Private Function Grab_Var_3(inputString As String, pp As Application, rr As String) As String 'Modbus var
    Dim V As String, LArray_T2() As String, P As String
    Dim rngVar As Range
    V = "": P = ""
    LArray_T2 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T2) To UBound(LArray_T2)
        If InStr(Replace(LArray_T2(i), " ", ""), "var=") Then
            V = Replace(LArray_T2(i), "var=", "")
            Exit For
        End If
    Next i
    For Each rngVar In pp.Range(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0), pp.Range("ModbusVar" & rr & "!B3").End(xlDown))
         If V = pp.WorksheetFunction.Trim(LCase(rngVar.Text)) Or V = pp.WorksheetFunction.Trim(LCase(rngVar.Offset(0, 1).Text)) Then
            P = rngVar.Text
            Exit For
         End If
    Next rngVar
    Grab_Var_3 = Right("00" & P, 2)
End Function


'added in v5.3 read write settings
Private Function Grab_Var_4(inputString As String) As String 'settings
    Dim V As String, LArray_T1() As String
    V = "00000000"
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    
    For i = LBound(LArray_T1) To UBound(LArray_T1)
        If InStr(Replace(LArray_T1(i), " ", ""), "var=") Then
            V = Replace(LArray_T1(i), "var=", "")
            Exit For
        End If
    Next i
    
    If X1.isValidHexString(V) Then
        If WorksheetFunction.Hex2Dec(V) < 65536 Then
            V = Right("00000000" & V, 8)
        End If
    End If
    
    Grab_Var_4 = Right(V, 2) & " " & Mid(V, 5, 2) & " " & Mid(V, 3, 2) & " " & Left(V, 2)
End Function


'Grab_Time_1 > return time Ticks
'Grab_Value_1 > retrun value input to functions
'Grab_Dev_1 > return device for 03 03

'time for run wait
Private Function Grab_Time_1(inputString As String, pp As Application) As String
    Dim P As Double, Q As Double, R As Double, t As Double, s As String
    Dim LArray_T1() As String
    P = 0: Q = 0: R = 0
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
        If InStr(Replace(LArray_T1(i), " ", ""), "hour=") Then
            P = X1.GrabPara_Dbl(LArray_T1(i), "hour")
        ElseIf InStr(Replace(LArray_T1(i), " ", ""), "min=") Then
            Q = X1.GrabPara_Dbl(LArray_T1(i), "min")
        ElseIf InStr(Replace(LArray_T1(i), " ", ""), "sec=") Then
            R = X1.GrabPara_Dbl(LArray_T1(i), "sec")
        End If
    Next i
    t = ((P * 3600) + (Q * 60) + R) * 10000
    s = pp.WorksheetFunction.Dec2Hex(t, 8)
    Grab_Time_1 = Right(s, 2) & " " & Mid(s, 5, 2) & " " & Mid(s, 3, 2) & " " & Left(s, 2)
End Function

'time for dimmer load
Private Function Grab_Time_2(inputString As String, pp As Application) As String
    Dim P As Double, Q As Double, R As Double, t As Double, s As String
    Dim LArray_T1() As String
    P = 0: Q = 0: R = 0
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
        If InStr(Replace(LArray_T1(i), " ", ""), "hour=") Then
            P = X1.GrabPara_Dbl(LArray_T1(i), "hour")
        ElseIf InStr(Replace(LArray_T1(i), " ", ""), "min=") Then
            Q = X1.GrabPara_Dbl(LArray_T1(i), "min")
        ElseIf InStr(Replace(LArray_T1(i), " ", ""), "sec=") Then
            R = X1.GrabPara_Dbl(LArray_T1(i), "sec")
        End If
    Next i
    t = ((P * 3600) + (Q * 60) + R) * 100
    s = pp.WorksheetFunction.Dec2Hex(t, 8)
    Grab_Time_2 = Right(s, 2) & " " & Mid(s, 5, 2)
End Function

'time for RCU dimmer
Private Function Grab_Time_3(inputString As String, pp As Application) As String
    Dim P As Double, Q As Double, R As Double, t As Double, s As String
    Dim LArray_T1() As String
    P = 0: Q = 0: R = 0
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
        If InStr(Replace(LArray_T1(i), " ", ""), "hour=") Then
            P = X1.GrabPara_Dbl(LArray_T1(i), "hour")
        ElseIf InStr(Replace(LArray_T1(i), " ", ""), "min=") Then
            Q = X1.GrabPara_Dbl(LArray_T1(i), "min")
        ElseIf InStr(Replace(LArray_T1(i), " ", ""), "sec=") Then
            R = X1.GrabPara_Dbl(LArray_T1(i), "sec")
        End If
    Next i
    t = ((P * 3600) + (Q * 60) + R) * 10000
    s = pp.WorksheetFunction.Dec2Hex(t, 8)
    Grab_Time_3 = Right(s, 2) & " " & Mid(s, 5, 2) & " " & Mid(s, 3, 2)
End Function

'send time for load time to use <save time> instrauction , added in XCODE 5.2
Private Function Give_Time_1() As String
    Dim YY As Integer, MM As Integer, DD As Integer, h As Integer, m As Integer, s As Integer
    YY = Year(Now)
    MM = Month(Now)
    DD = Day(Now)
    h = Hour(Format(Now, "MM/dd/yyyy HH:mm:ss"))
    m = Minute(Now)
    s = Second(Now)
    ' sec 60(64)6 : min 60(64)6 : hr 24(32)5 : day 31(32)5 : month 12(16)4 : year 64(64)6
    Give_Time_1 = WorksheetFunction.Dec2Bin(Right(YY, 2), 6) & WorksheetFunction.Dec2Bin(MM, 4) & WorksheetFunction.Dec2Bin(DD, 5) & WorksheetFunction.Dec2Bin(h, 5) & WorksheetFunction.Dec2Bin(m, 6) & WorksheetFunction.Dec2Bin(s, 6)
    Give_Time_1 = WorksheetFunction.Bin2Hex(Right(Give_Time_1, 8), 2) & " " & WorksheetFunction.Bin2Hex(Mid(Give_Time_1, 17, 8), 2) & " " & WorksheetFunction.Bin2Hex(Mid(Give_Time_1, 9, 8), 2) & " " & WorksheetFunction.Bin2Hex(Left(Give_Time_1, 8), 2)
End Function
 
Private Function hasTime(inputString As String) As Boolean
    If InStr(LCase(inputString), " $hour") Or InStr(LCase(inputString), " $min") Or InStr(LCase(inputString), " $sec") Then
        hasTime = True
    Else
        hasTime = False
    End If
End Function

Private Function isWaitABS(inputString As String) As Boolean   'XCODE 5.1 wait abs instruction check
    If InStr(inputString, "wait abs ") = 1 Then
        isWaitABS = True
    Else
        isWaitABS = False
    End If
End Function

'inteligent Device Grabber
'main function
Private Function Grab_Dev_1(inputString As String, pp As Application, rr As String) As String
    Dim P As Integer
    Dim LArray_T1() As String
    Dim LArray_GP() As String
    P = 0
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
           
        If InStr(Replace(LArray_T1(i), " ", ""), "dev=") Then
            LArray_GP = Split(Replace(LArray_T1(i), " ", ""), "=")
            If IsNumeric(LArray_GP(1)) And LArray_GP(0) = "dev" Then
                If CInt(LArray_GP(1)) > 0 And CInt(LArray_GP(1)) < (X1.getDevLimit(pp, rr) - 2) Then ' 21 Then flexible dev
                    P = CInt(LArray_GP(1))
                End If
            ElseIf (LArray_GP(1) = "server" Or LArray_GP(1) = "db" Or LArray_GP(1) = "ff") And LArray_GP(0) = "dev" Then
                P = 255
            ElseIf (LArray_GP(1) = "last" Or LArray_GP(1) = "previous" Or LArray_GP(1) = "fe") And LArray_GP(0) = "dev" Then
                P = 254
            ElseIf (Not IsNumeric(LArray_GP(1))) And LArray_GP(0) = "dev" Then
                P = Grab_ID_from_Dev_Name(LArray_GP(1), pp, rr)
            End If
        End If
    Next i
     
    Grab_Dev_1 = pp.WorksheetFunction.Dec2Hex(P, 2)
End Function

Private Function Grab_ID_from_Dev_Name(inputString As String, pp As Application, rr As String) As Integer
    If X1.checkSheet("Device" & rr, pp) Then
        Grab_ID_from_Dev_Name = 0
        'checking device validity
        Dim RcuDevList() As String, RcuDevNameList() As String, RCUDevType As String, DevEnumID As String
        RcuDevList = Split("01,02,03,04,04,05,05,06,06,07,07,08,0a", ",") 'should be in lcase for comparison
        RcuDevNameList = Split("idpg,tig,tag,ioexp,ioe,iodexp,iod,gsw,gs,ioexp,ioe,bsp,dali", ",")
        RCUDevType = "00"
        DevEnumID = "01"
           
        For i = LBound(RcuDevList) To UBound(RcuDevList)
            If InStr(inputString, RcuDevNameList(i)) Then
                RCUDevType = RcuDevList(i)
                If IsNumeric(Replace(inputString, RcuDevNameList(i), "")) Then
                    DevEnumID = pp.WorksheetFunction.Dec2Hex(CInt(Replace(inputString, RcuDevNameList(i), "") - 1), 1) & "1"
                End If
                Exit For
            End If
        Next i
           
        If RCUDevType <> "00" Then
            Dim ce As Range
            For Each ce In pp.Range("Device" & rr & "!C4:C" & X1.getDevLimit(pp, rr))
                If RCUDevType = "04" Then
                    If (LCase(ce.Text) = "04" Or LCase(ce.Text) = "07") And Replace(ce.Offset(0, -1).Text, " ", "") = DevEnumID Then
                        Grab_ID_from_Dev_Name = ce.Row - 3
                        Exit Function
                    End If
                Else
                    If LCase(ce.Text) = RCUDevType And Replace(ce.Offset(0, -1).Text, " ", "") = DevEnumID Then
                        Grab_ID_from_Dev_Name = ce.Row - 3
                        Exit Function
                    End If
                End If
            Next ce
            'MsgBox """" & inputString & """ not found, check enum" 'replaced with errorlog ver 5.0
            X1.addError "[RCU" & rr & " | Device" & rr & " | " & inputString & "][Device not found, check enum]"
        Else
            X1.addError "[RCU" & rr & " | Device" & rr & " | " & inputString & "][Device type not identified]"
        End If
    
    Else
        Grab_ID_from_Dev_Name = 0
        'MsgBox "Device" & rr & " Sheet not Found" 'replaced with errorlog ver 5.0
        X1.addError "[RCU" & rr & " | Device" & rr & "][Device Sheet not found]"
    End If
End Function

Private Function Grab_Dev_Type_1(inputString As String, pp As Application) As String
    Dim P As Integer
    Dim LArray_T1() As String
    Dim LArray_GP() As String
    P = 0
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
           
        If InStr(Replace(LArray_T1(i), " ", ""), "devtype=") Or InStr(Replace(LArray_T1(i), " ", ""), "type=") Then
            LArray_GP = Split(Replace(LArray_T1(i), " ", ""), "=")
            If IsNumeric(LArray_GP(1)) And (LArray_GP(0) = "devtype" Or LArray_GP(0) = "type") Then
                If CInt(LArray_GP(1)) > 0 And CInt(LArray_GP(1)) < 11 Then
                    P = CInt(LArray_GP(1))
                End If
            ElseIf (LArray_GP(1) = "server" Or LArray_GP(1) = "db" Or LArray_GP(1) = "ff") And (LArray_GP(0) = "devtype" Or LArray_GP(0) = "type") Then
                P = 255
            ElseIf (Not IsNumeric(LArray_GP(1))) And (LArray_GP(0) = "devtype" Or LArray_GP(0) = "type") Then
                P = Grab_Type_from_Name(LArray_GP(1))
            End If
        End If
    Next i
     
    Grab_Dev_Type_1 = pp.WorksheetFunction.Dec2Hex(P, 2)
End Function

Private Function Grab_Type_from_Name(inputString As String) As Integer
    
    Grab_Type_from_Name = 0
    Dim RcuDevList() As String, RcuDevNameList() As String
    RcuDevList = Split("01,02,03,04,04,05,05,06,06,07,07,08,10", ",")
    RcuDevNameList = Split("idpg,tig,tag,ioexp,ioe,iodexp,iod,gs,gsw,ioexp,ioe,bsp,dali", ",")
       
    For i = LBound(RcuDevList) To UBound(RcuDevList)
        If InStr(inputString, RcuDevNameList(i)) Then
            Grab_Type_from_Name = CInt(RcuDevList(i))
            Exit Function
        End If
    Next i
              
End Function


'------end of inteligent Device grabber


' inteligent register grabber

Private Function Grab_Reg_1(inputString As String) As String
    Dim P As Integer
    Dim LArray_T1() As String
    Dim LArray_GP() As String
    P = 0
    LArray_T1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_T1) To UBound(LArray_T1)
           
        If InStr(Replace(LArray_T1(i), " ", ""), "reg=") Then
            LArray_GP = Split(Replace(LArray_T1(i), " ", ""), "=")
            If IsNumeric(LArray_GP(1)) And LArray_GP(0) = "reg" Then
                If CInt(LArray_GP(1)) > -1 And CInt(LArray_GP(1)) < 65536 Then
                    P = CInt(LArray_GP(1))
                End If
            'ElseIf (Not IsNumeric(LArray_GP(1))) And LArray_GP(0) = "reg" Then
                'P = Grab_ID_from_Dev_Name(LArray_GP(1))
            End If
        End If
    Next i
     
    Grab_Reg_1 = Left(WorksheetFunction.Dec2Hex(P, 4), 2) & " " & Right(WorksheetFunction.Dec2Hex(P, 4), 2)
End Function



'------end of inteligent register grabber

'value for OR, relay wise
Private Function Grab_value_1(inputString As String, Vmode As String, pp As Application) As String
    Dim P As String, Q As Long, R As String, t As Integer
    Dim LArray_v1() As String
    P = "": Q = 0: R = "": t = 0
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(Replace(LArray_v1(i), " ", ""), "bin=") Then
            P = X1.GrabPara_BIN(LArray_v1(i))
            R = pp.WorksheetFunction.Bin2Hex(Left(P, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 9, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "relay=") And (Vmode = "OR") Then
            P = X1.GrabPara_BitPOS_OR(LArray_v1(i), "relay")
            R = pp.WorksheetFunction.Bin2Hex(Left(P, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 9, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "relay=") And (Vmode = "AND") Then
            P = X1.GrabPara_BitPOS_AND(LArray_v1(i), "relay")
            R = pp.WorksheetFunction.Bin2Hex(Left(P, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 9, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "bit=") And (Vmode = "OR") Then
            P = X1.GrabPara_BitPOS_OR(LArray_v1(i), "bit")
            R = pp.WorksheetFunction.Bin2Hex(Left(P, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 9, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "bit=") And (Vmode = "AND") Then
            P = X1.GrabPara_BitPOS_AND(LArray_v1(i), "bit")
            R = pp.WorksheetFunction.Bin2Hex(Left(P, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 9, 8), 2) & pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "dec=") Then
            Q = X1.GrabPara_Lng(LArray_v1(i), "dec")
            R = pp.WorksheetFunction.Dec2Hex(Q, 8)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "hex=") Then
            R = X1.GrabPara_HEX(LArray_v1(i))
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "fan=") Then
            R = X1.GrabPara_FAN(LArray_v1(i))
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "display=") Then
            Grab_value_1 = X1.GrabPara_Display(LArray_v1(i))
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "len=") Then
            t = X1.GrabPara_Int(LArray_v1(i), "len")
            R = pp.WorksheetFunction.Dec2Hex(t, 8)
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "clock=") Then
            R = X1.GrabPara_Clock(LArray_v1(i))
            Grab_value_1 = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
            Exit Function
        End If
    Next i
    Grab_value_1 = "00 00 00 00"
End Function

Private Function Grab_value_2(inputString As String, pp As Application) As String
    Dim P As String, Q As Double, R As String
    Dim LArray_v1() As String
    P = "": Q = 0: R = ""
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
            P = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "lum")
            R = pp.WorksheetFunction.Dec2Hex(P * 100, 8)
            Grab_value_2 = Right(R, 2) & " " & Mid(R, 5, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
            Q = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "dec")
            R = pp.WorksheetFunction.Dec2Hex(Q * 100, 8)
            Grab_value_2 = Right(R, 2) & " " & Mid(R, 5, 2)
            Exit Function
        End If
    Next i
    Grab_value_2 = "00 00"
End Function

Private Function Grab_value_3(inputString As String, Vmode As String, pp As Application) As String
    Dim P As String, Q As Long, R As String
    Dim LArray_v1() As String
    P = "": Q = 0: R = ""
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(Replace(LArray_v1(i), " ", ""), "bin=") Then
            P = X1.GrabPara_BIN(LArray_v1(i))
            R = pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_3 = Left(R, 2) & " " & Mid(R, 3, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "relay=") And (Vmode = "OR") Then
            P = X1.GrabPara_BitPOS_OR(LArray_v1(i), "relay")
            R = pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_3 = Left(R, 2) & " " & Mid(R, 3, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "relay=") And (Vmode = "AND") Then
            P = X1.GrabPara_BitPOS_AND(LArray_v1(i), "relay")
            R = pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_3 = Left(R, 2) & " " & Mid(R, 3, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "bit=") And (Vmode = "OR") Then
            P = X1.GrabPara_BitPOS_OR(LArray_v1(i), "bit")
            R = pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_3 = Left(R, 2) & " " & Mid(R, 3, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "bit=") And (Vmode = "AND") Then
            P = X1.GrabPara_BitPOS_AND(LArray_v1(i), "bit")
            R = pp.WorksheetFunction.Bin2Hex(Mid(P, 17, 8), 2) & pp.WorksheetFunction.Bin2Hex(Right(P, 8), 2)
            Grab_value_3 = Left(R, 2) & " " & Mid(R, 3, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "dec=") Then
            Q = X1.GrabPara_Lng(LArray_v1(i), "dec")
            R = pp.WorksheetFunction.Dec2Hex(Q, 4)
            Grab_value_3 = Left(R, 2) & " " & Mid(R, 3, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "hex=") Then
            R = X1.GrabPara_HEX(LArray_v1(i))
            Grab_value_3 = Mid(R, 5, 2) & " " & Right(R, 2)
            Exit Function
        End If
    Next i
    Grab_value_3 = "00 00"
End Function

Private Function Grab_value_4(inputString As String, pp As Application) As String
    Dim P As String, Q As Double, R As String
    Dim LArray_v1() As String
    P = "": Q = 0: R = ""
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
            P = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "lum")
            R = pp.WorksheetFunction.Dec2Hex(P, 8)
            Grab_value_4 = Right(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
            Q = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "dec")
            R = pp.WorksheetFunction.Dec2Hex(Q, 8)
            Grab_value_4 = Right(R, 2)
            Exit Function
        End If
    Next i
    Grab_value_4 = "00"
End Function

Private Function Grab_value_5(inputString As String, pp As Application) As String 'for dali
    Dim P As Double
    Dim LArray_v1() As String
    P = 0
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
            P = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "lum")
            R = P * 254 / 100
            Grab_value_5 = pp.WorksheetFunction.Dec2Hex(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
            P = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "dec")
            R = P * 254 / 100
            Grab_value_5 = pp.WorksheetFunction.Dec2Hex(R, 2)
            Exit Function
        End If
    Next i
    Grab_value_5 = "00"
End Function

Private Function Grab_value_6(inputString As String, pp As Application) As String 'for dmx
    Dim P As Double
    Dim LArray_v1() As String
    P = 0
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
            P = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "lum")
            R = P * 255 / 100
            Grab_value_6 = pp.WorksheetFunction.Dec2Hex(R, 2)
            Exit Function
        ElseIf InStr(Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
            P = X1.GrabPara_Dbl(Replace(LArray_v1(i), "%", ""), "dec")
            R = P * 255 / 100
            Grab_value_6 = pp.WorksheetFunction.Dec2Hex(R, 2)
            Exit Function
        End If
    Next i
    Grab_value_6 = "00"
End Function

Private Function Grab_value_0(inputString As String, index As Integer) As String 'return indexed value otherwise null
    Dim LArray_v1() As String, tempS As String, indexcount As Integer
    LArray_v1 = Split(Replace(inputString, "#", " $"), " $")
    indexcount = 0
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        tempS = Replace(LArray_v1(i), " ", "")
        If InStr(tempS, "bin=") = 1 Or InStr(tempS, "dec=") = 1 Or InStr(tempS, "hex=") = 1 _
        Or InStr(tempS, "relay=") = 1 Or InStr(tempS, "bit=") = 1 Or InStr(tempS, "bin=") = 1 _
        Or InStr(tempS, "fan=") = 1 Or InStr(tempS, "display=") = 1 Or InStr(tempS, "var=") Then
            Grab_value_0 = tempS
            indexcount = indexcount + 1
            If indexcount = index Then
                Exit Function
            End If
        End If
    Next i
    Grab_value_0 = "NULL"
End Function

Private Function Grab_element(inputString As String)
    Dim element As String
    element = Grab_variant(inputString, "element")
    If element = "NULL" Then
        Grab_element = "00 00 00 00"
    Else
        element = Right("00000000" & element, 8)
        Grab_element = Right(element, 2) & " " & Mid(element, 5, 2) & " " & Mid(element, 3, 2) & " " & Left(element, 2)
    End If
End Function

'variant
Private Function Grab_variant(inputString As String, dataType As String) As String
    Dim LArray_v1() As String
    LArray_v1 = Split(Replace(Replace(LCase(inputString), "#", "$"), " ", ""), "$")
    
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(LArray_v1(i), dataType & "=") = 1 Then
            Grab_variant = Replace(LArray_v1(i), dataType & "=", "")
            If Len(Grab_variant) < 1 Then
                Grab_variant = "NULL"
            End If
            Exit Function
        End If
    Next i
    Grab_variant = "NULL"
End Function

' ip address

Private Function Grab_IP_addr(inputString As String) As String
    Dim LArray_ip() As String, R As String
    LArray_ip = Split(Replace(inputString, " ", ""), ".")
    R = ""
    
    For k = LBound(LArray_ip) To UBound(LArray_ip)
        R = R & WorksheetFunction.Dec2Hex(LArray_ip(k), 2)
    Next k
    Grab_IP_addr = Right(R, 2) & " " & Mid(R, 5, 2) & " " & Mid(R, 3, 2) & " " & Left(R, 2)
End Function

'toolbar functions

Public Sub XCodeQuickIns_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        XCodeQuickIns.setPP pp
        XCodeQuickIns.Show vbModal
    End If
End Sub

'comment set and unset

Public Sub setAsComment_Function(pp As Application)
    For Each ce2 In pp.Selection
        If Len(pp.WorksheetFunction.Trim(ce2.Text)) > 0 And Not Left(pp.WorksheetFunction.Trim(ce2.Text), 2) = "//" Then '4.9 adding comment only to len>0 lines
            ce2.FormulaR1C1 = "//" & pp.WorksheetFunction.Trim(ce2.Text)
        End If
    Next ce2
End Sub

Public Sub unsetAsComment_Function(pp As Application)
    For Each ce2 In pp.Selection
        If Left(pp.WorksheetFunction.Trim(ce2.Text), 2) = "//" Then
            ce2.FormulaR1C1 = Mid(pp.WorksheetFunction.Trim(ce2.Text), 3, Len(pp.WorksheetFunction.Trim(ce2.Text)))
        End If
    Next ce2
End Sub

'insert or remove XCODE lines

Public Sub insertLineAbove_Function(pp As Application)
    Dim thisRCU As String, Rstart As Integer, Rheight As Integer
    thisRCU = X1.getRCUnumber(pp)
    If pp.ActiveSheet.Name = "XCODE" & thisRCU Then
        Rstart = pp.Selection.Row
        Rheight = pp.Selection.Rows.Count
        If X1.pointToXCode2(pp) Then
            pp.Range("C" & Rstart).Select
            pp.ScreenUpdating = False
            For i = 1 To Rheight
                pp.CutCopyMode = False
                pp.ActiveCell.Rows("1:1").EntireRow.Select
                pp.Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                pp.ActiveCell.Offset(-1, 1).Range("A1").Select
                pp.Selection.AutoFill Destination:=pp.ActiveCell.Range("A1:A3"), Type:=xlFillDefault
                pp.ActiveCell.Offset(1, 1).Range("A1").Select
            Next i
            pp.ScreenUpdating = True
        End If
    Else
        MsgBox "Out of reach of code", vbExclamation, "XCODE | Formatter"
    End If
End Sub

Public Sub insertLineBelow_Function(pp As Application)
    Dim thisRCU As String, Rstart As Integer, Rheight As Integer
    thisRCU = X1.getRCUnumber(pp)
    If pp.ActiveSheet.Name = "XCODE" & thisRCU Then
        Rstart = pp.Selection.Row
        Rheight = pp.Selection.Rows.Count
        If X1.pointToXCode2(pp) Then
            If Rheight > 1 Then
                pp.Range("C" & Rstart - 1).Select
            Else
                pp.Range("C" & Rstart).Select
            End If
            pp.ScreenUpdating = False
            For i = 1 To Rheight
                pp.CutCopyMode = False
                pp.ActiveCell.Offset(1, 1).Range("A1").Select
                pp.ActiveCell.Rows("1:1").EntireRow.Select
                pp.Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                pp.ActiveCell.Offset(-1, 1).Range("A1").Select
                pp.Selection.AutoFill Destination:=pp.ActiveCell.Range("A1:A3"), Type:=xlFillDefault
                pp.ActiveCell.Offset(1, 1).Range("A1").Select
            Next i
            If Rheight > 1 Then
                pp.Range("C" & Rstart).Select
            End If
            pp.ScreenUpdating = True
        End If
    Else
        MsgBox "Out of reach of code", vbExclamation, "XCODE | Formatter"
    End If
End Sub

Public Sub removeThisLine_Function(pp As Application)
    Dim thisRCU As String, Rstart As Integer, Rheight As Integer
    thisRCU = X1.getRCUnumber(pp)
    If pp.ActiveSheet.Name = "XCODE" & thisRCU Then
        Rstart = pp.Selection.Row
        Rheight = pp.Selection.Rows.Count
        If X1.pointToXCode2(pp) Then
            pp.Range("C" & Rstart).Select
            pp.ScreenUpdating = False
            For i = 1 To Rheight
                pp.ActiveCell.Rows("1:1").EntireRow.Select
                pp.Selection.Delete Shift:=xlUp
                pp.ActiveCell.Offset(-1, 1).Range("A1").Select
                pp.Selection.AutoFill Destination:=pp.ActiveCell.Range("A1:A2"), Type:=xlFillDefault
                pp.ActiveCell.Offset(1, 1).Range("A1").Select
            Next i
            pp.ScreenUpdating = True
        End If
    Else
        MsgBox "Out of reach of code", vbExclamation, "XCODE | Formatter"
    End If
End Sub

'variable handling

'rcu variables
Public Sub RCUVariableX_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        RcuVariablesX2.setPP pp
        RcuVariablesX2.Show vbModal
    End If
End Sub

'custom variable
Public Sub CustomVariableX_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        If X1.checkSheet("CustomVar" & thisRCU, pp) Then
            'If X1.checkCustomVarList(pp, thisRCU) Then ' removed in ver 5.0
            If X1.C_VarListNotEmpty(pp, thisRCU) Then
                CustomVariableX2.setPP pp, thisRCU
                CustomVariableX2.Show vbModal
            Else
                MsgBox "Custom Variable Table is empty." & vbNewLine & "Please add variables", vbExclamation, "XCODE | Formatter"
                X1.gotoEndofList pp, thisRCU
            End If
        Else
            X1.InsertCustomVarSheet pp, thisRCU
        End If
    End If
End Sub

'modbus register handling
Public Sub ModbusVariableX_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        If X1.checkSheet("ModbusVar" & thisRCU, pp) Then
            'If X1.checkModbusVarList(pp, thisRCU) Then ' removed in ver 5.0
            If X1.M_VarListNotEmpty(pp, thisRCU) Then
                ModbusVariableX2.setPP pp, thisRCU
                ModbusVariableX2.Show vbModal
            Else
                MsgBox "Modbus Variable Table is empty." & vbNewLine & "Please add variables", vbExclamation, "XCODE | Formatter"
                X1.gotoEndofMList pp, thisRCU
            End If
        Else
            X1.InsertModbusVarSheet pp, thisRCU
        End If
    End If
End Sub

'device function list

Public Sub XCodeDevFunc_Function(pp As Application)
    If X1.pointToXCode3(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeDevFunc.setPP pp, thisRCU
        XCodeDevFunc.Show vbModal
        'XCodeDevFunc.Show_C
    End If
End Sub

Public Sub XCodeRunCondition_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeRunCondition.setPP pp, thisRCU
        XCodeRunCondition.Show vbModal
    End If
End Sub

Public Sub XCodeRunReuse_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeRunReuse.setPP pp, thisRCU
        XCodeRunReuse.Show vbModal
    End If
End Sub

'function refactor
Public Sub XCodeFuncRefac_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String, thisRCUList As Variant
        thisRCU = X1.getRCUnumber(pp)
        thisRCUList = X1.RCUList(pp)
        XCodeFuncRefac.setPP pp, thisRCU, thisRCUList
        XCodeFuncRefac.Show vbModal
    End If
End Sub

'commissioning mode
Public Sub XCodeVarChecker_Function(pp As Application)
    'If X1.pointToXCode2(pp) Then ' removed in 4.9
        XCodeVarChecker.setPP pp
        XCodeVarChecker.Show vbModal
    'End If
End Sub

'FCU config
Public Sub XCodeFCUMode_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeFCUmode.setPP pp, thisRCU
        XCodeFCUmode.Show vbModal
    End If
End Sub

'Connectivity and integration Config
'lock wifi tcp modbus integration configs

Public Sub XCodeXCOMConfig_function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeXCOM.setPP pp, thisRCU
        XCodeXCOM.Show vbModal
    End If
End Sub


'go find
Public Sub gotoThisLineX(pp As Application)
    If X1.pointJumpX(pp) Then
        'code for the rcu memmap navigation
        
        Dim Linetext As String, thisRCU As String, lineNum As Integer
        thisRCU = X1.getRCUnumber(pp)
        If pp.ActiveSheet.Name = "Device" & thisRCU Then
        
            If Not pp.Intersect(pp.ActiveCell, pp.Range("A4:G" & X1.getDevLimit(pp, thisRCU))) Is Nothing Then
                LineNumberHex = Replace(Replace(pp.ActiveCell.Text, " ", ""), ";", "")
                LineNumberHex = Right(LineNumberHex, 1) & Left(LineNumberHex, 2)
                LineNumberHex = pp.WorksheetFunction.Hex2Dec(LineNumberHex)
                
                If CLng(LineNumberHex) > 0 Then
                    pp.ActiveSheet.Range("J" & CStr(CLng(LineNumberHex) + 4)).Select
                Else
                    MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                End If
            Else
                Linetext = LCase(pp.Range("Device" & thisRCU & "!N" & pp.ActiveCell.Row).Text)
                If X1.pointToXCode(pp) Then
                    If Len(Linetext) > 0 Then
                        If Not gotoFLocation("goto #" & Linetext, pp, 2, thisRCU) Then
                            MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                        End If
                    Else
                        MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                    End If
                End If
            End If
            
        ElseIf pp.ActiveSheet.Name = "XCODE" & thisRCU Then
        
            If pp.ActiveCell.Row = 2 Then
                XCodeStartup_Jump.setPP pp, thisRCU
                XCodeStartup_Jump.Label1.Caption = X1.getfilename(pp) & " Stratup Functions"
                XCodeStartup_Jump.Show vbModal
            ElseIf pp.ActiveCell.Row > 2 Then
                Dim otherRCU As String
                Linetext = LCase(pp.WorksheetFunction.Trim(pp.ActiveCell.Text))
                otherRCU = X1.getOtherRCU(Linetext, pp)
                
                
                If otherRCU <> "NULL" Then
                    If X1.checkSheet("XCODE" & otherRCU, pp) Then
                        pp.ActiveWorkbook.Sheets("XCODE" & otherRCU).Select
                        If hasAFunction(Linetext) Then
                            If Left(Linetext, 2) = "//" Then
                                MSG1 = MsgBox("This is a commented line, Do you still want to check the usage?", vbYesNo + vbQuestion, "Find Usage")
                                If MSG1 = vbYes Then
                                    If Not gotoFLocation(Linetext, pp, 2, otherRCU) Then
                                        MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                                    End If
                                End If
                            ElseIf Not gotoFLocation(Linetext, pp, 2, otherRCU) Then
                                MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                            End If
                        Else
                            MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                        End If
                    Else
                        MsgBox "XCODE" & otherRCU & " not found", vbExclamation, "XCODE | Error_Checker"
                    End If
                Else
                    If hasAFunction(Linetext) Then
                        If Left(Linetext, 2) = "//" Then
                            MSG1 = MsgBox("This is a commented line, Do you still want to check the usage?", vbYesNo + vbQuestion, "Find Usage")
                            If MSG1 = vbYes Then
                                If Not gotoFLocation(Linetext, pp, pp.ActiveCell.Row, thisRCU) Then
                                    MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                                End If
                            End If
                        ElseIf Not gotoFLocation(Linetext, pp, pp.ActiveCell.Row, thisRCU) Then
                            MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                        End If
                    Else
                        MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                    End If
                End If
            Else
                MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
            End If
            
        ElseIf pp.ActiveSheet.Name = "Program" & thisRCU Then
            
            'only works when active cell was in Program sheet
            
            If pp.ActiveCell.Row = 3 Then
                XCodeStartup_Jump.setPP pp, thisRCU, False
                XCodeStartup_Jump.Label1.Caption = X1.getfilename(pp) & " Stratup Functions"
                XCodeStartup_Jump.Show vbModal
            ElseIf pp.ActiveCell.Row > 3 Then
            
                Linetext = Replace(pp.ActiveCell.Text, " ", "")
                
                If X1.Jumptester(LCase(CStr(Linetext))) Then
                
                    Linetext = Mid(Linetext, 7, 2) & Mid(Linetext, 5, 2)
                    lineNum = pp.WorksheetFunction.Hex2Dec(Linetext)
                    
                    If lineNum > 0 Then
                        pp.ActiveSheet.Range("E" & CStr(lineNum + 3)).Select
                    Else
                        MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
                    End If
                    
                Else
                    MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
                End If
            Else
                MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
            End If
            
        End If
        
    ElseIf X1.isEOSmemmap(pp) Then
          
        
        If pp.ActiveSheet.Name = "Device" Then
            
        ElseIf pp.ActiveSheet.Name = "Dispatch" Then
            
            Linetext = LCase(pp.Range("Dispatch!P" & pp.ActiveCell.Row).Text)
            
            If X1.pointToEOSXCode(pp) Then
                If Len(Linetext) > 0 Then
                    If Not gotoFLocation("goto #" & Linetext, pp, 2, "") Then
                        MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                    End If
                Else
                    MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                End If
            End If
            
        ElseIf pp.ActiveSheet.Name = "XCODE" Then
        
            pp.Range("C" & pp.ActiveCell.Row).Select
        
            If pp.ActiveCell.Row = 2 Then
                                
                XCodeStartup_Jump.setPP pp, thisRCU
                XCodeStartup_Jump.Label1.Caption = X1.getfilename(pp) & " Stratup Functions"
                XCodeStartup_Jump.Show vbModal
            ElseIf pp.ActiveCell.Row > 2 Then
                Linetext = LCase(pp.WorksheetFunction.Trim(pp.ActiveCell.Text))
                
                If hasAFunction(Linetext) Then
                    If Left(Linetext, 2) = "//" Then
                        MSG1 = MsgBox("This is a commented line, Do you still want to check the usage?", vbYesNo + vbQuestion, "Find Usage")
                        If MSG1 = vbYes Then
                            If Not gotoFLocation(Linetext, pp, pp.ActiveCell.Row, thisRCU) Then
                                MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                            End If
                        End If
                    ElseIf Not gotoFLocation(Linetext, pp, pp.ActiveCell.Row, thisRCU) Then
                        MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                    End If
                Else
                    MsgBox "No location found for jumping from here", vbInformation, "XCODE | GOTO"
                End If
            
            End If
            
        ElseIf pp.ActiveSheet.Name = "Script" Then
            
            pp.Range("E" & pp.ActiveCell.Row).Select
            
            If pp.ActiveCell.Row = 3 Then
                XCodeStartup_Jump.setPP pp, thisRCU, False
                XCodeStartup_Jump.Label1.Caption = X1.getfilename(pp) & " Stratup Functions"
                XCodeStartup_Jump.Show vbModal
            ElseIf pp.ActiveCell.Row > 3 Then
            
                Linetext = Replace(pp.ActiveCell.Text, " ", "")
                
                If X1.Jumptester(LCase(CStr(Linetext))) Then
                
                    Linetext = Mid(Linetext, 7, 2) & Mid(Linetext, 5, 2)
                    lineNum = pp.WorksheetFunction.Hex2Dec(Linetext)
                    
                    If lineNum > 0 Then
                        pp.ActiveSheet.Range("E" & CStr(lineNum + 3)).Select
                    Else
                        MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
                    End If
                    
                Else
                    MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
                End If
            Else
                MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
            End If
            
        Else
            MsgBox "Out of reach of code", vbCritical, "XCODE | GOTO"
        End If
    Else
        MsgBox "Memmap definition fails to identify", vbCritical, "XCODE | GOTO"
    End If
    
End Sub


Private Function gotoFLocation(inString As String, pp As Application, fromLine As Integer, rr As String) As Boolean
    gotoFLocation = False
    
    Dim Fname As String, FArray() As String
    Fname = ""
    
    FArray() = Split(inString, " ")
    For k = LBound(FArray) To UBound(FArray)
        If InStr(FArray(k), "#") = 1 Then
            Fname = Replace(FArray(k), "#", "")
            Exit For
        End If
    Next k
    
    If Len(Fname) > 0 Then
        If Not IsNumeric(Fname) Then
            Dim ce1 As Range
    
            For Each ce1 In pp.Range(pp.Range("XCODE" & rr & "!B2"), pp.Range("XCODE" & rr & "!B2").End(xlDown))
                If ce1.Offset(0, 1).Text <> "" Then
                    
                    If InStr(ce1.Offset(0, 1).Text, "//") > 2 Then 'filter if there is a comment in the same line to jump
                        Linetext = Left(ce1.Offset(0, 1).Text, InStr(ce1.Offset(0, 1).Text, "//") - 1)
                        Linetext = LCase(pp.WorksheetFunction.Trim(Linetext))
                    Else
                        Linetext = LCase(pp.WorksheetFunction.Trim(ce1.Offset(0, 1).Text))
                    End If
                    
                    If Linetext = "function " & Fname Then
                        pp.Range("XCODE" & rr & "!C" & ce1.Row).Select
                        gotoFLocation = True
                        Exit Function
                    ElseIf Fname = "exit" And Linetext = "exit" And ce1.Row > fromLine Then
                        pp.Range("XCODE" & rr & "!C" & ce1.Row).Select
                        gotoFLocation = True
                        Exit Function
                    End If
                End If
            Next ce1
            
        Else
            MsgBox "Jumping specific number of lines not predictable in this format", vbExclamation, "XCODE | GOTO"
        End If
    End If
End Function


'-----------[Group TAG]----------

Public Sub XCodeTAGDisplayX2_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeTagDisplayX2.setPP pp, thisRCU
        XCodeTagDisplayX2.Show vbModal
    End If
End Sub

Public Sub XCodeTAGX_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        XCodeTAG.setPP pp, thisRCU
        XCodeTAG.Show vbModal
    End If
End Sub

'-----------[Group GSW]----------

Public Sub XCodeGSW12X_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        
        MSG1 = MsgBox(X1.getfilename(pp) & vbNewLine & "Haven't Saved in a while." & vbNewLine & "Would you like to save it now?", vbYesNo + vbQuestion, "Save Your XCODE")
        If MSG1 = vbYes Then
            pp.ActiveWorkbook.Save
        End If
        
        XCodeGSW12X.setPP pp, thisRCU
        XCodeGSW12X.Show vbModal
    End If
End Sub

Public Sub XCodeGSW6X_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        
        MSG1 = MsgBox(X1.getfilename(pp) & vbNewLine & "Haven't Saved in a while." & vbNewLine & "Would you like to save it now?", vbYesNo + vbQuestion, "Save Your XCODE")
        If MSG1 = vbYes Then
            pp.ActiveWorkbook.Save
        End If
        
        XCodeGSW6X.setPP pp, thisRCU
        XCodeGSW6X.Show vbModal
    End If
End Sub

Public Sub XCodeGSW9X_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        
        MSG1 = MsgBox(X1.getfilename(pp) & vbNewLine & "Haven't Saved in a while." & vbNewLine & "Would you like to save it now?", vbYesNo + vbQuestion, "Save Your XCODE")
        If MSG1 = vbYes Then
            pp.ActiveWorkbook.Save
        End If
        
        XCodeGSW9X.setPP pp, thisRCU
        XCodeGSW9X.Show vbModal
    End If
End Sub

Public Sub XCodeIDPG_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim thisRCU As String
        thisRCU = X1.getRCUnumber(pp)
        
        MSG1 = MsgBox(X1.getfilename(pp) & vbNewLine & "Haven't Saved in a while." & vbNewLine & "Would you like to save it now?", vbYesNo + vbQuestion, "Save Your XCODE")
        If MSG1 = vbYes Then
            pp.ActiveWorkbook.Save
        End If
        
        XCodeIDPG.setPP pp, thisRCU
        XCodeIDPG.Show vbModal
    End If
End Sub

Public Sub SaveAllXcode_Function(pp As Application)
    X1.SaveAllXcode pp
End Sub

Public Sub SendAllEmailXcode_Function(pp As Application)
    X1.SendAllEmailXcode pp
End Sub

Public Function evalOpenFilePath_Function(pp As Application) As String
    Dim evalPath As String
    evalPath = X1.getFilePathFormat(pp)
    
    If evalPath > 0 Then
        evalOpenFilePath_Function = X1.getfilepath(pp)
    Else
        evalOpenFilePath_Function = "NULL"
    End If
End Function

Public Sub show_info(pp As Application)
    pp.ScreenUpdating = False
    If X1.pointToXCode4(pp) Then
        If X1.checkSheet("XLog", pp) Then
            pp.Sheets("XLog").Select
            pp.Range("XLog!B3").Select
        Else
            Dim vLine As String, vCompile As Integer, vhexDate As String, vYear As String, vMonth As String, vDay As String, vbuild As String
            vLine = Replace(Replace(pp.Range("Program" & X1.getRCUnumber(pp) & "!E4").Text, " ", ""), ";", "")
            
            If Left(vLine, 4) = "0000" Then
                vCompile = CInt(pp.WorksheetFunction.Hex2Dec(Mid(vLine, 9, 2)))
                vhexDate = Mid(vLine, 11, 4)
                vhexDate = pp.WorksheetFunction.Hex2Bin(Left(vhexDate, 2), 8) & pp.WorksheetFunction.Hex2Bin(Right(vhexDate, 2), 8)
                vDay = pp.WorksheetFunction.Bin2Dec(Right(vhexDate, 5))
                vMonth = pp.WorksheetFunction.Bin2Dec(Mid(vhexDate, 8, 4))
                vYear = pp.WorksheetFunction.Bin2Dec(Left(vhexDate, 7))
                'vbuild = CInt(pp.WorksheetFunction.Hex2Dec(Right(vLine, 2)))
                vbuild = Right(vLine, 2) ' ver 5.0 show version in hex
                
                MsgBox "Logic buid version" & vbTab & vbTab & ": " & vbuild & vbNewLine & vbNewLine & "Last built date" & vbTab & vbTab & ": " & vDay & "/" & vMonth & "/" & (2000 + vYear) & vbNewLine & vbNewLine & "XCODE version" & vbTab & vbTab & ": " & (vCompile / 10), vbInformation, "XCODE | Info"
            Else
                MsgBox "No version data found", vbExclamation, "XCODE | Error_Checker"
            End If
        End If
    End If
    pp.ScreenUpdating = True
End Sub

Public Sub ChangeMyPrevXLogComment_Function(pp As Application)
    'pp.ScreenUpdating = False
    
    If X1.pointToXCode2(pp) Then
        If X1.checkSheet("XLog", pp) Then
            Dim PreFileName As String, PreVersion As String, PreUser As String, thisVersion As String
            Dim thisRCUList As Variant, firstRCU As String, tempVLine As String, ErrMsg As String
            
            ErrMsg = ""
            thisRCUList = X1.RCUList(pp)
            firstRCU = thisRCUList(LBound(thisRCUList))
            tempVLine = Replace(pp.Range("Program" & firstRCU & "!E4").Text, ";", "")
            thisVersion = Right(tempVLine, 2)
            
            PreFileName = pp.Range("XLog!B3").Text
            PreVersion = pp.Range("XLog!C3").Text
            PreUser = pp.Range("XLog!D3").Text
            
            If PreFileName = X1.getfilename(pp) And PreVersion = thisVersion And PreUser = pp.UserName Then
                X1.XLogUpdater pp, firstRCU, pp.WorksheetFunction.Hex2Dec(thisVersion), pp.Range("XLog!F3").Text
            Else
                If PreFileName <> X1.getfilename(pp) Then
                    ErrMsg = ErrMsg & "Current File [" & X1.getfilename(pp) & "] Name does not match with Last Built File Name [" & PreFileName & "]." & vbNewLine
                End If
                If PreVersion <> thisVersion Then
                    ErrMsg = ErrMsg & "This Built version [" & thisVersion & "] does not match with Last Built version " & PreVersion & "]." & vbNewLine
                End If
                If PreUser <> pp.UserName Then
                    ErrMsg = ErrMsg & "You are not the one who compiled this logic last time." & vbNewLine
                End If
                
                ErrMsg = "Not allowed to modify last XLog due to following reasons, " & vbNewLine & vbNewLine & ErrMsg
                MsgBox ErrMsg, vbExclamation, "XCODE | XLog | Error_Checker"
            End If
            
        Else
            MsgBox "No XLog data found", vbExclamation, "XCODE | Error_Checker"
        End If
    End If
    
    'pp.ScreenUpdating = True
End Sub
