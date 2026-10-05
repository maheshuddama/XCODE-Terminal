Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX2
    Dim xlDeviceSheet As Excel.Worksheet
    Dim xlXCODESheet As Excel.Worksheet
    Dim xlProgramSheet As Excel.Worksheet
    Dim xlSettingsSheet As Excel.Worksheet
    Dim xlCustomVarSheet As Excel.Worksheet
    Dim xlModbusVarSheet As Excel.Worksheet
    Dim xlFUNCTIONSheet As Excel.Worksheet


    Dim RCUname As String
    Dim CustomVarDictionary As Dictionary(Of String, String)
    Dim ModbusVarDictionary As Dictionary(Of String, String)
    Dim FunctionDictionary As Dictionary(Of String, String)

    Dim X0 As XCX0

    Sub XCODE_Core1_Core2(MemmapName As String, ByRef thisWorkBook As Excel.Workbook, ByRef xlSheetList As List(Of String), ByRef thisRCUList As List(Of String),
                          ByRef ErrorWarnLog As String(), unAttended As Boolean, ByRef RCUNameDictionary As Dictionary(Of String, String), ByRef IPDictionary As Dictionary(Of String, String))


        X0 = New XCX0(unAttended)

        For Each thisRCU As String In thisRCUList

            If thisRCU = "" Then
                RCUname = ""
            Else
                RCUname = "RCU" & thisRCU & " > "
            End If

            X0.ConsoleMsg("XC Progress:> " & RCUname & "XCODE > Core - 1 > Compiling.......")

            xlDeviceSheet = thisWorkBook.Sheets("Device" & thisRCU)
            xlXCODESheet = thisWorkBook.Sheets("XCODE" & thisRCU)
            xlProgramSheet = thisWorkBook.Sheets("Program" & thisRCU)
            xlSettingsSheet = thisWorkBook.Sheets("Settings" & thisRCU)

            'X0.ClearFunctionSeparaters(xlProgramSheet) 'optional
            'adding function sheet
            If Not X0.IsSheetExist_From_List(xlSheetList, "FUNCTION" & thisRCU) Then
                thisWorkBook.Sheets.Add(After:=thisWorkBook.Sheets("Settings" & thisRCU)).Name = "FUNCTION" & thisRCU
                'xlSheetList.Add("FUNCTION" & thisRCU) 'optional not required to add
            End If
            xlFUNCTIONSheet = thisWorkBook.Sheets("FUNCTION" & thisRCU)
            FunctionDictionary = New Dictionary(Of String, String)

            'assign CustomVar and ModbusVar sheets
            'CustomVar sheet exist since RCUList checks
            xlCustomVarSheet = thisWorkBook.Sheets("CustomVar" & thisRCU)
            CustomVarDictionary = New Dictionary(Of String, String)
            Populate_C_Var_Dictionary()

            'check and populate ModbusVar dictionary
            If X0.IsSheetExist_From_List(xlSheetList, "ModbusVar" & thisRCU) Then
                xlModbusVarSheet = thisWorkBook.Sheets("ModbusVar" & thisRCU)
                ModbusVarDictionary = New Dictionary(Of String, String)
                Populate_M_Var_Dictionary()
            Else
                xlModbusVarSheet = Nothing
                ModbusVarDictionary = Nothing
            End If

            'clear XCODE
            X0.ClearProgFUN(xlProgramSheet, xlFUNCTIONSheet)

            X0.Print_XCODE_build(xlXCODESheet)
            X0.Check_XCODE_format(MemmapName, xlXCODESheet, thisRCU)

            Dim lineNum As Integer, Fcount As Integer, pcount As Integer, pfullcount As Integer
            Dim ThisTempLine As String, ThisHexCode As String, thisString As String, thisListOfString As List(Of String)
            lineNum = 0 : Fcount = 0 : pcount = 0 : pfullcount = 0

            'XCodeProgress.Show()
            'XCodeProgress.ProgressBar_val(pcount)
            pfullcount = xlXCODESheet.Range(xlXCODESheet.Range("B2"), xlXCODESheet.Range("B2").End(Excel.XlDirection.xlDown)).Count '1524

            'optimization
            Dim XCODEArray As New List(Of String())

            '------------------------------
            Dim ScripttargetRange As Excel.Range = xlXCODESheet.Range("C2:C" & (pfullcount + 1))
            Dim ScriptValues As Object(,) = DirectCast(ScripttargetRange.Value, Object(,))
            Dim rowCount As Integer = ScriptValues.GetLength(0)
            '------------------------------

            For r As Integer = 1 To rowCount
                'For Each ce1 In xlXCODESheet.Range(xlXCODESheet.Range("B2"), xlXCODESheet.Range("B2").End(Excel.XlDirection.xlDown))

                ThisTempLine = If(ScriptValues(r, 1)?.ToString(), String.Empty).Trim().ToLower()

                If ThisTempLine <> "" Then
                    'ThisTempLine = Strings.LCase(Strings.Trim(ce1.Offset(0, 1).Text))

                    If Strings.InStr(ThisTempLine, "//") > 2 AndAlso Not Strings.Left(ThisTempLine, 2) = "0x" Then 'comment detection and filter except direct hex
                        ThisTempLine = Strings.Trim(Strings.Left(ThisTempLine, Strings.InStr(ThisTempLine, "//") - 1))
                    End If

                    If Strings.Left(ThisTempLine, 2) = "//" Then
                        'do nothing this is a comment
                    ElseIf ThisTempLine.StartsWith("0x") Then
                        Dim thisHexString As String = X0.GrabHexString(ThisTempLine)
                        If thisHexString <> "FALSE" Then
                            'insertXCODE(lineNum, X0.GrabComment(ThisTempLine))
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, X0.GrabComment(ThisTempLine), thisHexString))
                            'insertHEXCODE(lineNum, X0.GrabHexString(ThisTempLine))
                            lineNum += 1
                        End If
                    ElseIf X0.IsOtherRCUcall(ThisTempLine) Then
                        'insertXCODE(lineNum, "//" & ThisTempLine)
                        XCODEArray.Add(GetXCODEArrayObject(lineNum, "//" & ThisTempLine))
                        'insertHEXCODE(lineNum, "01 01 00 00 00 00 00 00;")
                        lineNum += 1
                    ElseIf X0.IsAFunction(ThisTempLine) Then
                        Dim thisFunctionName As String = X0.GrabFunctionName(ThisTempLine)
                        If IsNotDuplicateFunction(thisFunctionName) Then
                            InsertFunction(Fcount, lineNum, thisFunctionName)
                            Fcount += 1
                        Else
                            ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & pcount & " | " & thisFunctionName & "][Duplicate Function]")
                        End If
                    ElseIf X0.IsAnIOEXPDirect(ThisTempLine) Then
                        thisListOfString = GrabIOEXPDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsAnIODEXPDirect(ThisTempLine) Then
                        thisListOfString = GrabIODEXPDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsARCUDirect(ThisTempLine) Then
                        thisListOfString = GrabRCUDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsAMathDirect(ThisTempLine) Then
                        thisListOfString = GrabMathDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf IsOtherVarDirect(ThisTempLine) Then
                        thisListOfString = GrabOtherVarDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsASaveDirect(ThisTempLine) Then
                        thisListOfString = GrabSaveDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsAModbusDirect(ThisTempLine) Then
                        thisListOfString = GrabModbusDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsADimmerDirect(ThisTempLine) Then
                        thisListOfString = GrabDimmerDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsADimmerReadDirect(ThisTempLine) Then
                        thisListOfString = GrabDimmerReadDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsADimmerLoad(ThisTempLine) Then
                        'insertXCODE(lineNum, "load $hex=1")
                        XCODEArray.Add(GetXCODEArrayObject(lineNum, "load $hex=1"))
                        lineNum += 1
                        'insertXCODE(lineNum, "//" & ThisTempLine)
                        XCODEArray.Add(GetXCODEArrayObject(lineNum, "//" & ThisTempLine, X0.LoadDimmerConvert(ThisTempLine)))
                        'insertHEXCODE(lineNum, X0.LoadDimmerConvert(ThisTempLine))
                        lineNum += 1
                    ElseIf X0.IsADaliDirect(ThisTempLine) Then
                        thisListOfString = GrabDaliDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsADMXDirect(ThisTempLine) Then
                        thisListOfString = GrabDMXDirect(ThisTempLine, thisRCU, pcount, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            'insertXCODE(lineNum, thisString)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, thisString))
                            lineNum += 1
                        Next
                    ElseIf X0.IsADefinition(ThisTempLine) Then
                        If Strings.InStr(ThisTempLine, " rcu") > 6 AndAlso Strings.InStr(ThisTempLine, "$ip=") > 6 Then
                            'ip definition example
                            'define rcu2 $ip=192.168.1.224 
                            Dim RCUIP As String = X0.Grab_variant(ThisTempLine, "ip")
                            Dim RCUNum As String = X0.Grab_RCU_Num(ThisTempLine)

                            If RCUNum <> "NULL" AndAlso RCUIP <> "NULL" Then
                                If X0.IsAValidIP(RCUIP) Then
                                    If Not IPDictionary.ContainsKey("XCODE" & RCUNum) Then
                                        IPDictionary.Add("XCODE" & RCUNum, RCUIP)
                                    End If
                                Else
                                    ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Definition of RCU" & RCUNum & " $ip=" & RCUIP & "][Invalid IP address]")
                                End If
                            End If

                        ElseIf Strings.InStr(ThisTempLine, " rcu ") > 6 AndAlso Strings.InStr(ThisTempLine, "$name=") > 6 Then
                            'rcu name definition example
                            'define rcu $name=king-room
                            Dim getDefinedName As String = X0.Grab_name(ThisTempLine)

                            If getDefinedName <> "NULL" Then
                                RCUNameDictionary.Add("XCODE" & thisRCU, getDefinedName)
                            End If
                        End If
                    Else
                        If ThisTempLine.Replace(" "c, "").Length < 2 Then
                            'ignore multiple space empty lines and lines less then 2 charactors
                        Else
                            'insertXCODE(lineNum, ThisTempLine)
                            XCODEArray.Add(GetXCODEArrayObject(lineNum, ThisTempLine))
                            lineNum += 1
                        End If
                    End If
                End If

                pcount += 1
                X0.ConsoleProgress(RCUname & "XCODE > Core - 1 > Compiling", (pcount * 50 / pfullcount))

            Next


            '--------------optimization-------------
            Dim XCODEArraynumRows As Integer = XCODEArray.Count
            Dim XCODEArraynumCols As Integer = XCODEArray(0).Length
            Dim resultArray(XCODEArraynumRows - 1, XCODEArraynumCols - 1) As String
            For i As Integer = 0 To XCODEArraynumRows - 1
                For j As Integer = 0 To XCODEArraynumCols - 1
                    resultArray(i, j) = XCODEArray(i)(j)
                Next j
            Next i
            xlProgramSheet.Range("D3").Resize(XCODEArraynumRows, XCODEArraynumCols).Value = resultArray
            '---------------------------


            InsertHEXCODE(-2, ";")
            InsertHEXCODE(-1, ";code")
            InsertStartupHEXCODE(0, xlProgramSheet.Range("F3").Text, thisRCU)
            lineNum = 1 : pcount = 0
            X0.ConsoleProgress(RCUname & "XCODE > Core -1 > Compiling", 50)
            pfullcount = xlProgramSheet.Range(xlProgramSheet.Range("F4"), xlProgramSheet.Range("F4").End(Excel.XlDirection.xlDown)).Count
            X0.ConsoleMsg("XC Progress:> " & RCUname & "XCODE > Core - 2 > Compiling.......")

            '------------------------------optimization------
            Dim HEXArray(pfullcount - 1, 0) As String
            '------------------------------------------------

            ScripttargetRange = xlProgramSheet.Range("F4:F" & (pfullcount + 3))
            Dim HextargetRange As Excel.Range = xlProgramSheet.Range("E4:E" & (pfullcount + 3)) 'should be identical height
            ScriptValues = DirectCast(ScripttargetRange.Value, Object(,))
            Dim HexValues As Object(,) = DirectCast(HextargetRange.Value, Object(,))
            rowCount = ScriptValues.GetLength(0)
            Dim thisRow As Integer, thisHexRow As String

            'For Each ce2 In xlProgramSheet.Range(xlProgramSheet.Range("F4"), xlProgramSheet.Range("F4").End(Excel.XlDirection.xlDown))
            For r As Integer = 1 To rowCount

                ThisTempLine = If(ScriptValues(r, 1)?.ToString(), String.Empty).Trim() 'already in lcase
                thisRow = r + 3 'excel row
                thisHexRow = r.ToString("X4")
                ThisHexCode = "01 01 00 00 00 00 00 00;"


                If ThisTempLine.StartsWith("//") Then
                    ThisHexCode = If(HexValues(r, 1)?.ToString(), String.Empty).Trim()
                    InsertXCODE(lineNum, Strings.Trim(Strings.Replace(ThisTempLine, "//", "")))
                ElseIf X0.HasAFunction(ThisTempLine) Then
                    ThisHexCode = FunctionConvert(ThisTempLine, thisRow, thisWorkBook, thisRCU, ErrorWarnLog) 'ce2 for nearest exit
                ElseIf X0.IsARCU(ThisTempLine) Then
                    ThisHexCode = RCUConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.IsADMX(ThisTempLine) Then
                    ThisHexCode = DMXConvert(ThisTempLine)
                ElseIf X0.IsWaitABS(ThisTempLine) Then
                    ThisHexCode = WaitABSConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.HasTime(ThisTempLine) Then
                    ThisHexCode = X0.WaitConvert(ThisTempLine)
                ElseIf X0.IsAnExit(ThisTempLine) Then
                    ThisHexCode = X0.ExitConvert
                ElseIf X0.IsDebug(ThisTempLine) Then
                    ThisHexCode = X0.DebugConvert
                ElseIf X0.IsVersion(ThisTempLine) Then
                    ThisHexCode = X0.VersionConvert(ThisTempLine)
                ElseIf X0.IsALoad(ThisTempLine) Then
                    ThisHexCode = LoadConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.IsASend(ThisTempLine) Then
                    ThisHexCode = SendConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.IsARead(ThisTempLine) Then
                    ThisHexCode = ReadConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.IsASave(ThisTempLine) Then
                    ThisHexCode = SaveConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.IsAMath(ThisTempLine) Then
                    ThisHexCode = MathConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                ElseIf X0.IsADimFlow(ThisTempLine) Then
                    ThisHexCode = DimFlowConvert(ThisTempLine, thisHexRow, thisRCU, ErrorWarnLog)
                Else
                    'else part goes here, catching poorly written instructions
                    If Strings.InStr(ThisTempLine, ".set") > 1 OrElse Strings.InStr(ThisTempLine, ".unset") > 1 OrElse Strings.InStr(ThisTempLine, ".or") > 1 OrElse Strings.InStr(ThisTempLine, ".and") > 1 OrElse Strings.InStr(ThisTempLine, ".xor") > 1 Then  '[dropped from 5.2] Or Strings.instr(ThisTempLine, "$") > 2 Then
                        ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisHexRow & " | " & ThisTempLine & "][Undefined variable | " & Strings.Left(ThisTempLine, Strings.InStr(ThisTempLine, ".") - 1) & "]")
                    ElseIf (Strings.InStr(ThisTempLine, "wait ") > 0) AndAlso (Not (Strings.InStr(ThisTempLine, "$hour") > 5 OrElse Strings.InStr(ThisTempLine, "$min") > 5 OrElse Strings.InStr(ThisTempLine, "$sec") > 5)) Then 'added from 5.2
                        ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisHexRow & " | " & ThisTempLine & "][Invalid parameter | " & Strings.Replace(ThisTempLine, "wait ", "") & "]")
                    End If
                End If

                'insertHEXCODE(lineNum, ThisHexCode)
                HEXArray(lineNum - 1, 0) = ThisHexCode
                lineNum += 1

                pcount += 1
                X0.ConsoleProgress(RCUname & "XCODE > Core - 2 > Compiling", 50 + (pcount * 50 / pfullcount))
            Next


            InsertHEXArray(HEXArray, pfullcount)
            'X0.setFunctionSeparaters(xlProgramSheet) 'optional
            FunctionCalibrate(thisRCU)

            'devlist size fix
            Dim devListSize As Integer = GetDevLimit() - 3 'remove offset
            If xlSettingsSheet.Range("B29").Text <> (CStr(devListSize) & ";") Then
                xlSettingsSheet.Range("B29").FormulaR1C1 = devListSize.ToString("X2") & ";"
            End If

        Next


        ReleaseObject(xlDeviceSheet)
        ReleaseObject(xlXCODESheet)
        ReleaseObject(xlProgramSheet)
        ReleaseObject(xlSettingsSheet)
        ReleaseObject(xlCustomVarSheet)
        ReleaseObject(xlModbusVarSheet)
        ReleaseObject(xlFUNCTIONSheet)
        ReleaseObject(X0)
    End Sub

    'inserters xcode
    Private Sub InsertXCODE(cLineNum As Integer, cXCODE As String, Optional HEXCODE As String = "01 01 00 00 00 00 00 00;")
        'xlProgramSheet.Range("D" & (cLineNum + 3)).FormulaR1C1 = xlWorkFunc.Dec2Hex(cLineNum, 4)
        'xlProgramSheet.Range("F" & (cLineNum + 3)).FormulaR1C1 = cXCODE
        xlProgramSheet.Range("D" & (cLineNum + 3)).Resize(1, 3).Value = New Object(,) {{cLineNum.ToString("X4"), HEXCODE, cXCODE}}
    End Sub

    Private Function GetXCODEArrayObject(cLineNum As Integer, cXCODE As String, Optional HEXCODE As String = "01 01 00 00 00 00 00 00;") As String()
        GetXCODEArrayObject = New String() {cLineNum.ToString("X4"), HEXCODE, cXCODE}
    End Function

    Private Sub InsertHEXCODE(cLineNum As Integer, cXCODE As String)
        xlProgramSheet.Range("E" & (cLineNum + 3)).FormulaR1C1 = cXCODE
    End Sub


    Private Sub InsertHEXArray(cArray As String(,), cHeight As Integer)
        xlProgramSheet.Range("E4").Resize(cHeight, 1).Value = cArray
    End Sub

    Private Sub InsertFunction(cFcount As Integer, cLineNum As Integer, Fname As String)
        'xlFUNCTIONSheet.Range("D" & (cFcount + 3)).FormulaR1C1 = cFcount
        'xlFUNCTIONSheet.Range("E" & (cFcount + 3)).FormulaR1C1 = Fname
        'xlFUNCTIONSheet.Range("F" & (cFcount + 3)).FormulaR1C1 = xlWorkFunc.Dec2Hex(cLineNum, 4)

        xlFUNCTIONSheet.Range("D" & (cFcount + 3)).Resize(1, 3).Value = New Object(,) {{cFcount, Fname, cLineNum.ToString("X4")}}
        FunctionDictionary.Add(Fname, cLineNum.ToString("X4"))
    End Sub

    Private Sub InsertStartupHEXCODE(cLineNum As Integer, cXCODE As String, thisRCU As String)
        Dim LArray1() As String, thisFunc() As String, LineJ() As String
        Dim i As Integer, j As Integer, cFC As String

        thisFunc = "NULL,NULL,NULL,NULL".Split(","c)
        LineJ = "0000,0000,0000,0000".Split(","c)
        LArray1 = cXCODE.Split(" "c)

        j = LBound(LineJ)
        For i = LBound(LArray1) To UBound(LArray1)
            If LArray1(i).Contains("#"c) AndAlso j < 4 Then
                thisFunc(j) = LArray1(i).Replace("#"c, "")
                LineJ(j) = GetLineFromFunctionName(thisFunc(j), 3, thisRCU)
                j += 1
            End If
        Next i

        cFC = Strings.Right(LineJ(0), 2) & " " & Strings.Left(LineJ(0), 2) & " " & Strings.Right(LineJ(1), 2) & " " & Strings.Left(LineJ(1), 2) & " " & Strings.Right(LineJ(2), 2) & " " & Strings.Left(LineJ(2), 2) & " " & Strings.Right(LineJ(3), 2) & " " & Strings.Left(LineJ(3), 2) & ";"
        xlProgramSheet.Range("E" & (cLineNum + 3)).FormulaR1C1 = cFC
    End Sub

    ' devices function calibration
    'Private Sub FunctionCalibrate(thisRCU As String)
    '    Dim ce3 As Excel.Range, pcount As Double, pfullcount As Double
    '    pcount = 0 : pfullcount = 0
    '    pfullcount = xlDeviceSheet.Range(xlDeviceSheet.Range("J4"), xlDeviceSheet.Range("J4").End(Excel.XlDirection.xlDown)).Count

    '    For Each ce3 In xlDeviceSheet.Range(xlDeviceSheet.Range("J4"), xlDeviceSheet.Range("J4").End(Excel.XlDirection.xlDown))
    '        If ce3.Offset(0, 4).Text <> "" Then
    '            Dim thisFunc As String, tempJumpOld As String, tempJump As String

    '            thisFunc = Strings.LCase(Strings.Trim(ce3.Offset(0, 4).Text))
    '            If Not isNotDuplicateFunction(thisFunc) Then
    '                tempJumpOld = Strings.Mid(Strings.Trim(Strings.Replace(ce3.Text, ";", "")), 4, 1)
    '                tempJump = tempJumpOld & Strings.Right(getLineFromFunctionName(thisFunc, 3, thisRCU), 3)
    '                ce3.FormulaR1C1 = Strings.Right(tempJump, 2) & " " & Strings.Left(tempJump, 2) & ";"
    '                setFunctionNode(ce3.Row)
    '            Else
    '                If ce3.Offset(0, 4).MergeCells Then
    '                    ce3.Offset(0, 4).MergeArea.Interior.ColorIndex = 0
    '                    ce3.Offset(0, 4).MergeArea.ClearContents()
    '                    ce3.Offset(0, 4).MergeArea.UnMerge()
    '                End If

    '                ce3.Offset(0, 4).Interior.ColorIndex = 0
    '                ce3.Offset(0, 4).ClearContents()

    '                unsetFunctionNode(ce3.Row)
    '            End If
    '        Else
    '            unsetFunctionNode(ce3.Row)
    '        End If

    '        pcount = pcount + 1

    '        X0.ConsoleProgress("Formatter > Device" & thisRCU & " inputs Calibration", (pcount * 100 / pfullcount))

    '    Next ce3

    'End Sub

    Private Sub FunctionCalibrate(thisRCU As String)
        Dim startCell As Excel.Range
        Dim dataRange As Excel.Range
        Dim ce3 As Excel.Range
        Dim pcount As Double, pfullcount As Double
        Dim offsetCell As Excel.Range
        Dim tempJumpOld As String, tempJump As String, newLine As String

        pcount = 0
        startCell = xlDeviceSheet.Range("J4")
        dataRange = xlDeviceSheet.Range(startCell, startCell.End(Excel.XlDirection.xlDown))
        pfullcount = dataRange.Count

        For Each ce3 In dataRange

            offsetCell = ce3.Offset(0, 4)

            If Trim(offsetCell.Text) <> "" Then
                Dim thisFunc As String
                thisFunc = LCase(Trim(offsetCell.Text))

                If Not IsNotDuplicateFunction(thisFunc) Then

                    tempJumpOld = Mid(Replace(Trim(ce3.Text), ";"c, ""), 4, 1)
                    newLine = GetLineFromFunctionName(thisFunc, 3, thisRCU)

                    If Left(newLine, 1) = "0" Then
                        tempJump = tempJumpOld & Right(newLine, 3)
                        ce3.FormulaR1C1 = Right(tempJump, 2) & " " & Left(tempJump, 2) & ";"
                        SetFunctionNode(CStr(ce3.Row))
                    Else
                        If offsetCell.MergeCells Then
                            With offsetCell.MergeArea
                                .Interior.ColorIndex = 0
                                '.ClearContents()
                                .UnMerge()
                            End With
                        End If

                        With offsetCell
                            .Interior.ColorIndex = 0
                            '.ClearContents()
                        End With

                        UnsetFunctionNode(CStr(ce3.Row))
                    End If
                Else
                    If offsetCell.MergeCells Then
                        With offsetCell.MergeArea
                            .Interior.ColorIndex = 0
                            .ClearContents()
                            .UnMerge()
                        End With
                    End If

                    With offsetCell
                        .Interior.ColorIndex = 0
                        .ClearContents()
                    End With

                    UnsetFunctionNode(CStr(ce3.Row))
                End If
            Else
                UnsetFunctionNode(CStr(ce3.Row))
            End If

            pcount += 1
            X0.ConsoleProgress("Formatter > Device" & thisRCU & " inputs Calibration", (pcount * 100.0# / pfullcount))
        Next ce3

        startCell = Nothing
        dataRange = Nothing
        ce3 = Nothing
        offsetCell = Nothing
    End Sub

    'Private Sub setFunctionNode(inputRange As String)
    '    xlDeviceSheet.Range("J" & inputRange).Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
    '    xlDeviceSheet.Range("J" & inputRange).Interior.Color = 65535
    '    xlDeviceSheet.Range("J" & inputRange).Interior.TintAndShade = 0
    'End Sub

    'Private Sub unsetFunctionNode(inputRange As String)
    '    xlDeviceSheet.Range("J" & inputRange).Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
    '    xlDeviceSheet.Range("J" & inputRange).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4
    '    xlDeviceSheet.Range("J" & inputRange).Interior.TintAndShade = 0.799981688894314
    '    xlDeviceSheet.Range("J" & inputRange).FormulaR1C1 = "00 00;"
    'End Sub

    Private Sub SetFunctionNode(inputRange As String)
        Dim cell As Object
        cell = xlDeviceSheet.Range("J" & inputRange)

        With cell.Interior
            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
    End Sub

    Private Sub UnsetFunctionNode(inputRange As String)
        Dim cell As Object
        cell = xlDeviceSheet.Range("J" & inputRange)

        With cell.Interior
            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
        End With

        cell.FormulaR1C1 = "00 00;"
    End Sub

    'Functions
    Private Function IsNotDuplicateFunction(inputString As String) As Boolean
        'If Strings.Len(xlFUNCTIONSheet.Range("E3").Text) = 0 Then
        '    isNotDuplicateFunction = True
        'ElseIf Strings.Len(xlFUNCTIONSheet.Range("E4").Text) = 0 Then
        '    If xlFUNCTIONSheet.Range("E3").Text = inputString Then
        '        isNotDuplicateFunction = False
        '    Else
        '        isNotDuplicateFunction = True
        '    End If
        'Else
        '    For Each ce As Excel.Range In xlFUNCTIONSheet.Range(xlFUNCTIONSheet.Range("E3"), xlFUNCTIONSheet.Range("E3").End(Excel.XlDirection.xlDown))
        '        If ce.Text = inputString Then
        '            isNotDuplicateFunction = False
        '            Exit Function
        '        End If
        '    Next ce
        '    isNotDuplicateFunction = True
        'End If

        If FunctionDictionary.ContainsKey(inputString) Then
            Return False
        End If

        Return True
    End Function


    'convertors

    'goto, gotowait, 03 01, 03 02, 03 03, if conditions
    Private Function FunctionConvert(inputString As String, inputLine As Integer, ByRef thisWorkbook As Excel.Workbook, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim LArray() As String, ConditionF As String, i As Integer
        Dim thisFunction As String, cFunctionConvert As String, LineJump As String, damper As String
        Dim withVar As Boolean

        cFunctionConvert = "00 00" : withVar = False : damper = "00"
        ConditionF = ConditionSet(inputString, thisRCU, ErrorWarnLog)
        withVar = X0.HasValidVar(inputString) OrElse HasValidC_Var(inputString)
        LArray = inputString.Split(" "c)

        For i = LBound(LArray) To UBound(LArray)
            If LArray(i).StartsWith("#") Then
                thisFunction = LArray(i).Replace("#"c, "")
                LineJump = GetLineFromFunctionName(thisFunction, inputLine, thisRCU) 'input line for nearest exit
                cFunctionConvert = Strings.Right(LineJump, 2) & " " & Strings.Left(LineJump, 2)
                Exit For 'skip next items in for loop
            End If
        Next i

        If cFunctionConvert = "00 00" Then
            If ConditionF = "01 01" Then
                'thread redirect pointer not found, therefore it should not allow to pass down using nop. it should be an exit.
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & inputLine & " | " & inputString & "][Invalid function call]")
                FunctionConvert = "04 01 00 00 00 00 00 00;"
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & inputLine & " | " & inputString & "][Invalid function call]")
                FunctionConvert = "01 01 00 00 00 00 00 00;"
            End If
        Else
            If ConditionF = "01 01" OrElse ConditionF = "03 01" OrElse ConditionF = "03 02" Then
                FunctionConvert = ConditionF & " " & cFunctionConvert & " " & X0.Grab_Time(inputString) & ";"
            ElseIf ConditionF = "03 03" Then
                FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Strings.Left(X0.Grab_Time(inputString), 8) & " " & Grab_Dev_1(inputString, thisRCU, ErrorWarnLog) & ";"
            Else
                If withVar Then
                    ConditionF = Strings.Left(ConditionF, 2) & " 0A"

                    If ConditionF = "5c 0A" Then  '5c damper "02" in the end
                        damper = X0.Grab_variant(inputString, "milisec")

                        If damper <> "NULL" AndAlso IsNumeric(damper) Then
                            If CInt(damper) > 0 AndAlso CInt(damper) < 256 Then
                                damper = CInt(damper).ToString("X2")
                            Else
                                damper = "02"
                            End If
                        Else
                            damper = "02"
                        End If
                    End If

                    If X0.HasValidVar(inputString) AndAlso (Not HasValidC_Var(inputString)) Then
                        FunctionConvert = ConditionF & " " & cFunctionConvert & " " & X0.Grab_Var_1(inputString) & " 00 00 " & damper & ";"
                    ElseIf (Not X0.HasValidVar(inputString)) AndAlso HasValidC_Var(inputString) Then
                        FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Grab_Var_2(inputString) & " 00 00 " & damper & ";"
                    Else
                        ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & inputLine & " | " & inputString & "][Invalid function call]")
                        FunctionConvert = "01 01 00 00 00 00 00 00;"
                    End If
                Else
                    FunctionConvert = ConditionF & " " & cFunctionConvert & " " & X0.Grab_value_1(inputString, "OR") & ";"
                End If
            End If
        End If

    End Function

    Private Function ConditionSet(inputString As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String

        If inputString.StartsWith("if ") Then
            If (Strings.InStr(inputString, "bit ") > 0 OrElse Strings.InStr(inputString, "bitwise ") > 0) AndAlso (Strings.InStr(inputString, "true ") > 0 OrElse Strings.InStr(inputString, "check ") > 0) Then
                ConditionSet = "09 01"
            ElseIf (Strings.InStr(inputString, "bit ") > 0 OrElse Strings.InStr(inputString, "bitwise ") > 0) AndAlso Strings.InStr(inputString, "run ") > 0 Then
                ConditionSet = "60 01"
            ElseIf Strings.InStr(inputString, "not equal ") > 0 OrElse Strings.InStr(inputString, "not equals to ") > 0 OrElse Strings.InStr(Strings.Replace(inputString, " "c, ""), "if<>") = 1 OrElse Strings.InStr(Strings.Replace(inputString, " "c, ""), "if!=") = 1 Then
                ConditionSet = "05 01"
            ElseIf Strings.InStr(inputString, "equal ") > 0 OrElse Strings.InStr(inputString, "equals to ") > 0 OrElse Strings.InStr(Strings.Replace(inputString, " "c, ""), "if=") = 1 Then
                ConditionSet = "08 01"
            ElseIf Strings.InStr(inputString, "greater ") > 0 OrElse Strings.InStr(inputString, "greater than ") > 0 OrElse Strings.InStr(Strings.Replace(inputString, " "c, ""), "if>") = 1 Then
                ConditionSet = "06 01"
            ElseIf Strings.InStr(inputString, "lesser ") > 0 OrElse Strings.InStr(inputString, "less than ") > 0 OrElse Strings.InStr(Strings.Replace(inputString, " "c, ""), "if<") = 1 Then
                ConditionSet = "07 01"
            Else
                ConditionSet = "01 01"
            End If
        ElseIf inputString.StartsWith("dimflow.") Then
            If Strings.InStr(inputString, ".oncheck ") > 0 OrElse Strings.InStr(inputString, ".checkon ") > 0 Then
                ConditionSet = "5b 0A"
            ElseIf Strings.InStr(inputString, ".offcheck ") > 0 OrElse Strings.InStr(inputString, ".checkoff ") > 0 Then
                ConditionSet = "5c 0A"
            Else
                ConditionSet = "01 01"
            End If
        ElseIf Strings.InStr(inputString, "run ") = 1 OrElse Strings.InStr(inputString, "reuse ") = 1 OrElse Strings.InStr(inputString, "re-run ") = 1 OrElse Strings.InStr(inputString, "rerun ") = 1 Then
            If Strings.InStr(inputString, "reuse ") = 1 OrElse Strings.InStr(inputString, "re-run ") = 1 OrElse Strings.InStr(inputString, "rerun ") = 1 Then
                If Convert.ToByte(Grab_Dev_1(inputString, thisRCU, ErrorWarnLog), 16) > 0 Then
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

    Private Function GetLineFromFunctionName(inputString As String, inputLine As Integer, thisRCU As String) As String
        Dim ce As Excel.Range
        GetLineFromFunctionName = "0000"

        If inputString = "exit" OrElse inputString = "quit" OrElse inputString = "end" Then
            If Strings.Len(xlProgramSheet.Range("F" & (inputLine + 1)).Text) = 0 Then
                'to remove this condition from else
            ElseIf Strings.Len(xlProgramSheet.Range("F" & (inputLine + 2)).Text) = 0 Then
                If xlProgramSheet.Range("F" & (inputLine + 1)).Text = inputString Then
                    GetLineFromFunctionName = (inputLine - 3).ToString("X4")
                End If
            Else
                For Each ce In xlProgramSheet.Range(xlProgramSheet.Range("F" & (inputLine + 1)), xlProgramSheet.Range("F" & (inputLine + 1)).End(Excel.XlDirection.xlDown))
                    If ce.Text = inputString Then
                        GetLineFromFunctionName = (ce.Row - 3).ToString("X4")
                        Exit Function
                    End If
                Next ce
            End If
        ElseIf IsNumeric(inputString) Then
            If (inputLine + CInt(inputString) - 3) > 0 Then
                GetLineFromFunctionName = (inputLine + CInt(inputString) - 3).ToString("X4")
            End If
        Else
            'If Strings.Len(xlFUNCTIONSheet.Range("E3").Text) = 0 Then
            '    'to remove this condition from else
            'ElseIf Strings.Len(xlFUNCTIONSheet.Range("E4").Text) = 0 Then
            '    If xlFUNCTIONSheet.Range("E3").Text = inputString Then
            '        getLineFromFunctionName = xlFUNCTIONSheet.Range("F3").Text
            '    End If
            'Else
            '    For Each ce In xlFUNCTIONSheet.Range(xlFUNCTIONSheet.Range("E3"), xlFUNCTIONSheet.Range("E3").End(Excel.XlDirection.xlDown))
            '        If ce.Text = inputString Then
            '            getLineFromFunctionName = ce.Offset(0, 1).Text
            '            Exit Function
            '        End If
            '    Next ce
            'End If
            If FunctionDictionary.ContainsKey(inputString) Then
                GetLineFromFunctionName = FunctionDictionary(inputString)
            End If
        End If
    End Function



    ' rcu convert
    Private Function RCUConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim Opt As String, Smode As String
        Opt = "09" : Smode = "01"

        If Strings.InStr(inputString, " acc") > 1 Then
            Smode = "00"
        End If

        If Strings.InStr(inputString, "rcu.read input") = 1 Then 'removed in 4.3 <Or Strings.instr(inputString, "rcu.input") = 1>
            Opt = "08"
        ElseIf Strings.InStr(inputString, "rcu.read mosfet") = 1 Then
            Opt = "0C"
        ElseIf Strings.InStr(inputString, "rcu.read batt") = 1 Then
            Opt = "0E"
        ElseIf Strings.InStr(inputString, "rcu.read es") = 1 Then
            Opt = "H1"
        ElseIf Strings.InStr(inputString, "rcu.read aout2") = 1 OrElse Strings.InStr(inputString, "rcu.read analog2") = 1 Then
            Opt = "H3"
        ElseIf Strings.InStr(inputString, "rcu.read aout") = 1 OrElse Strings.InStr(inputString, "rcu.read analog") = 1 Then
            Opt = "H2"
        ElseIf Strings.InStr(inputString, "rcu.read") = 1 Then
            Opt = "03"
        ElseIf Strings.InStr(inputString, "rcu.xor ") = 1 Then
            Opt = "0B"
        ElseIf Strings.InStr(inputString, "rcu.set batt") = 1 Then
            Opt = "0D"
        ElseIf Strings.InStr(inputString, "rcu.unset batt") = 1 Then
            Opt = "0D"
        ElseIf Strings.InStr(inputString, "rcu.set ") = 1 AndAlso Strings.InStr(inputString, " all") > 0 Then
            Opt = "04"
        ElseIf (Strings.InStr(inputString, "rcu.set ") = 1 OrElse Strings.InStr(inputString, "rcu.unset ") = 1) AndAlso Strings.InStr(inputString, " es") > 0 Then
            Opt = "05"
        ElseIf (Strings.InStr(inputString, "rcu.set ") = 1 OrElse Strings.InStr(inputString, "rcu.unset ") = 1) AndAlso (Strings.InStr(inputString, " analog2") > 0 OrElse Strings.InStr(inputString, " aout2") > 0) Then
            Opt = "07"
        ElseIf (Strings.InStr(inputString, "rcu.set ") = 1 OrElse Strings.InStr(inputString, "rcu.unset ") = 1) AndAlso (Strings.InStr(inputString, " analog") > 0 OrElse Strings.InStr(inputString, " aout") > 0) Then
            Opt = "06"
        ElseIf Strings.InStr(inputString, "rcu.set ") = 1 OrElse Strings.InStr(inputString, "rcu.or ") = 1 Then
            Opt = "09"
        ElseIf Strings.InStr(inputString, "rcu.unset ") = 1 OrElse Strings.InStr(inputString, "rcu.and ") = 1 Then
            Opt = "0A"
        ElseIf Strings.InStr(inputString, "rcu.pwm ") = 1 Then
            Opt = "00"
        ElseIf Strings.InStr(inputString, "rcu.dim es") = 1 Then
            Opt = "G0"
        ElseIf Strings.InStr(inputString, "rcu.dim aout2") = 1 OrElse Strings.InStr(inputString, "rcu.dim analog2") = 1 Then
            Opt = "G2"
        ElseIf Strings.InStr(inputString, "rcu.dim aout") = 1 OrElse Strings.InStr(inputString, "rcu.dim analog") = 1 Then
            Opt = "G1"
        ElseIf Strings.InStr(inputString, "rcu.stop es") = 1 Then
            Opt = "I0"
        ElseIf Strings.InStr(inputString, "rcu.stop aout2") = 1 OrElse Strings.InStr(inputString, "rcu.stop analog2") = 1 Then
            Opt = "I2"
        ElseIf Strings.InStr(inputString, "rcu.stop aout") = 1 OrElse Strings.InStr(inputString, "rcu.stop analog") = 1 Then
            Opt = "I1"
        End If

        If Opt = "03" OrElse Opt = "08" OrElse Opt = "0C" OrElse Opt = "0E" Then
            RCUConvert = "02 " & Opt & " 00 00 00 00 00 00;"
        ElseIf Opt = "H1" OrElse Opt = "H2" OrElse Opt = "H3" Then
            RCUConvert = "53 0" & Strings.Right(Opt, 1) & " 00 00 00 00 00 00;"
        ElseIf X0.HasValidVar(inputString) Then
            If Strings.Left(Opt, 1) = "G" OrElse Strings.Left(Opt, 1) = "I" Then
                RCUConvert = "5" & Strings.Right(Opt, 1) & " 0A 00 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            Else
                RCUConvert = "02 " & Opt & " 0A 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            End If
        ElseIf HasValidC_Var(inputString) Then
            If Strings.Left(Opt, 1) = "G" OrElse Strings.Left(Opt, 1) = "I" Then
                RCUConvert = "5" & Strings.Right(Opt, 1) & " 0A 00 00 " & Grab_Var_2(inputString) & " 00 00 00;"
            Else
                RCUConvert = "02 " & Opt & " 0A 00 " & Grab_Var_2(inputString) & " 00 00 00;"
            End If
        Else
            If Smode = "00" Then
                If Strings.Left(Opt, 1) = "G" OrElse Strings.Left(Opt, 1) = "I" Then
                    RCUConvert = "5" & Strings.Right(Opt, 1) & " 00 00 00 00 00 00 00;"
                Else
                    RCUConvert = "02 " & Opt & " 00 00 00 00 00 00;"
                End If
            Else
                If Strings.InStr(inputString, "$var=") > 4 Then
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                    RCUConvert = "01 01 00 00 00 00 00 00;"
                Else
                    If Opt = "05" OrElse Opt = "06" OrElse Opt = "07" Then
                        If Strings.InStr(inputString, "rcu.unset ") = 1 Then
                            RCUConvert = "02 " & Opt & " 01 00 00 00 00 00;"
                        Else
                            RCUConvert = "02 " & Opt & " 01 00 " & X0.Grab_value_4(inputString) & " 00 00 00;"
                        End If
                    ElseIf Opt = "0D" Then
                        If Strings.InStr(inputString, "rcu.unset") = 1 Then
                            RCUConvert = "02 " & Opt & " 01 00 00 00 00 00;"
                        ElseIf Strings.InStr(inputString, "rcu.set") = 1 AndAlso Strings.InStr(inputString, " $") < 1 Then
                            RCUConvert = "02 " & Opt & " 01 00 01 00 00 00;"
                        Else
                            RCUConvert = "02 " & Opt & " 01 00 " & X0.Grab_value_4(inputString) & " 00 00 00;"
                        End If
                    ElseIf Strings.Left(Opt, 1) = "G" OrElse Strings.Left(Opt, 1) = "I" Then
                        RCUConvert = "5" & Strings.Right(Opt, 1) & " 01 00 00 " & X0.Grab_Time(inputString, 3) & " " & X0.Grab_value_4(inputString) & ";"
                    Else
                        RCUConvert = "02 " & Opt & " 01 00 " & X0.Grab_value_1(inputString, "OR") & ";"
                    End If
                End If
            End If
        End If
    End Function

    ' dmx convert
    Private Function DMXConvert(inputString As String) As String
        Dim Opt As String, Smode As String, ioCh As String
        Opt = "01" : Smode = "01"

        If Strings.InStr(inputString, " acc") > 1 Then
            Smode = "00"
        End If

        ioCh = X0.Grab_variant(inputString, "ch")

        If Strings.InStr(inputString, "dmx.send ") = 1 OrElse Strings.InStr(inputString, "dmx.ramp ") = 1 OrElse Strings.InStr(inputString, "dmx.set ") = 1 Then
            Opt = "01"
        ElseIf Strings.InStr(inputString, "dmx.stop ") = 1 OrElse Strings.InStr(inputString, "dmx.hold ") = 1 Then
            Opt = "02"
        ElseIf Strings.InStr(inputString, "dmx.read ") = 1 Then
            Opt = "03"
        End If

        If ioCh <> "NULL" Then
            ioCh = CInt(ioCh).ToString("X2")

            If Opt = "03" Then
                Return "61 02 00 " & ioCh & " 00 00 00 00;"
            ElseIf Opt = "02" Then
                Return "61 01 01 " & ioCh & " 00 00 00 00;"
            Else
                If Smode = "00" Then
                    Return "61 01 00 " & ioCh & " 00 00 00 00;"
                ElseIf X0.HasValidVar(inputString) Then
                    Return "61 01 0A " & ioCh & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
                ElseIf HasValidC_Var(inputString) Then
                    Return "61 01 0A " & ioCh & " " & Grab_Var_2(inputString) & " 00 00 00;"
                Else
                    Return "61 01 01 " & ioCh & " " & X0.Grab_Time(inputString, 3) & " " & X0.Grab_value_6(inputString) & ";"
                End If
            End If
        Else
            Return "01 01 00 00 00 00 00 00;"
        End If
    End Function

    'abs convert
    Private Function WaitABSConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        If X0.HasValidVar(inputString) Then
            Return "0A 0A 00 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf HasValidC_Var(inputString) Then
            Return "0A 0A 00 00 " & Grab_Var_2(inputString) & " 00 00 00;"
        ElseIf X0.HasTime(inputString) Then
            Return "0A 01 00 00 " & X0.Grab_Time(inputString) & ";"
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid Wait ABS instruction]")
            Return "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' load convert
    Private Function LoadConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String

        Dim Smode As String
        Smode = "00"
        If Strings.InStr(inputString, "shift ") > 1 Then
            Smode = "01"
        End If

        If Strings.InStr(inputString, " set") > 1 Then
            Dim thisLen As String
            'added in v5.3, checking first to avoid duplicate detection in variable space
            thisLen = X0.Grab_variant(inputString, "len")

            If Not thisLen = "NULL" Then
                If X0.isValidHexString(thisLen) Then
                    thisLen = Convert.ToInt32(thisLen, 16).ToString("D2")
                Else
                    ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Incompatible length | " & X0.Grab_variant(inputString, "len") & "]")
                    thisLen = "01"
                End If
            Else
                thisLen = "01"
            End If

            If thisLen = "04" Then
                LoadConvert = "10 80 18 " & Smode & " " & X0.Grab_Var_4(inputString) & ";"
            ElseIf thisLen = "02" Then
                LoadConvert = "10 80 17 " & Smode & " " & X0.Grab_Var_4(inputString) & ";"
            Else
                LoadConvert = "10 80 16 " & Smode & " " & X0.Grab_Var_4(inputString) & ";"
            End If

        ElseIf X0.HasValidVar(inputString) Then
            LoadConvert = "10 80 0A " & Smode & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf HasValidC_Var(inputString) Then
            LoadConvert = "10 80 0A " & Smode & " " & Grab_Var_2(inputString) & " 00 00 00;"
        ElseIf HasValidM_Var(inputString) AndAlso Strings.InStr(inputString, " mod") > 1 AndAlso Strings.InStr(inputString, " flo") > 1 Then 'added in ver 5.3
            LoadConvert = "10 80 15 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf HasValidM_Var(inputString) AndAlso Strings.InStr(inputString, " mod") > 1 Then
            LoadConvert = "10 80 12 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf Strings.InStr(inputString, " cbs") > 1 Then
            LoadConvert = "10 85 01 " & Smode & " " & X0.Grab_element(inputString) & ";"
        ElseIf Strings.InStr(inputString, " enum") > 1 Then
            LoadConvert = "10 80 10 " & Smode & " 00 00 00 00;"
        ElseIf Strings.InStr(inputString, " time") > 1 Then
            If Strings.InStr(inputString, " unix") > 1 OrElse Strings.InStr(inputString, " utc") > 1 Then
                LoadConvert = "10 84 00 " & Smode & " 00 00 00 00;"
            ElseIf Strings.InStr(inputString, " now") > 1 Then
                LoadConvert = "10 80 01 " & Smode & " " & X0.Give_Time_1() & ";"
            Else
                LoadConvert = "10 82 00 " & Smode & " 00 00 00 00;"
            End If
        Else
            If Strings.InStr(inputString, "$var=") > 4 Then
                ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                LoadConvert = "01 01 00 00 00 00 00 00;"
            Else
                LoadConvert = "10 80 01 " & Smode & " " & X0.Grab_value_1(inputString, "OR") & ";"
            End If
        End If
    End Function

    ' send convert
    Private Function SendConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        If (Strings.InStr(inputString, "send.mod ") = 1 OrElse Strings.InStr(inputString, "send.modbus ") = 1) AndAlso (Strings.InStr(inputString, " $dev=") > 5 OrElse Strings.InStr(inputString, " $type=") > 5) AndAlso Strings.InStr(inputString, " $reg=") > 5 Then
            Dim thisDev As String = Grab_Dev_1(inputString, thisRCU, ErrorWarnLog)
            If Convert.ToByte(thisDev, 16) > 0 Then
                '20 & 22
                If inputString.Contains("$len=") Then
                    Return "22 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_1(inputString, "OR") & ";"
                Else
                    Return "20 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_3(inputString, "OR") & " 00 00;"
                End If
            ElseIf Convert.ToByte(X0.Grab_Dev_Type_1(inputString), 16) > 0 Then
                '21 & 23
                If inputString.Contains("$len=") Then
                    Return "23 " & X0.Grab_Dev_Type_1(inputString) & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_1(inputString, "OR") & ";"
                Else
                    Return "21 " & X0.Grab_Dev_Type_1(inputString) & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_3(inputString, "OR") & " 00 00;"
                End If
            Else
                Return "01 01 00 00 00 00 00 00;"
            End If
        Else

            ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid modbus instruction]")
            Return "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' read convert
    Private Function ReadConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        If Strings.InStr(inputString, "read.mod ") = 1 OrElse Strings.InStr(inputString, "read.modbus ") = 1 Then
            Dim thisDev As String = Grab_Dev_1(inputString, thisRCU, ErrorWarnLog)
            If Convert.ToByte(thisDev, 16) > 0 Then
                '20 & 22
                If inputString.Contains("$len=") Then
                    Return "26 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_1(inputString, "OR") & ";"
                Else
                    Return "26 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " 02 00 00 00;"
                End If
            Else
                If Convert.ToByte(thisDev, 16) = 0 Then
                    ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid device]")
                End If
                Return "01 01 00 00 00 00 00 00;"
            End If
        Else
            ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid modbus instruction]")
            Return "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' save convert
    Private Function SaveConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim Smode As String
        Smode = "00"

        If Strings.InStr(inputString, "shift ") > 1 Then
            Smode = "01"
        End If

        If Strings.InStr(inputString, " set") > 1 Then
            Dim thisLen As String

            thisLen = X0.Grab_variant(inputString, "len")

            If Not thisLen = "NULL" Then
                If X0.isValidHexString(thisLen) Then
                    thisLen = Convert.ToInt32(thisLen, 16).ToString("D2")
                Else
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Incompatible length | " & X0.Grab_variant(inputString, "len") & "]")
                    thisLen = "01"
                End If
            Else
                thisLen = "01"
            End If

            If thisLen = "04" Then
                SaveConvert = "10 81 18 " & Smode & " " & X0.Grab_Var_4(inputString) & ";"
            ElseIf thisLen = "02" Then
                SaveConvert = "10 81 17 " & Smode & " " & X0.Grab_Var_4(inputString) & ";"
            Else
                SaveConvert = "10 81 16 " & Smode & " " & X0.Grab_Var_4(inputString) & ";"
            End If
        ElseIf X0.HasValidVar(inputString) Then
            SaveConvert = "10 81 0A " & Smode & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf HasValidC_Var(inputString) Then
            SaveConvert = "10 81 0A " & Smode & " " & Grab_Var_2(inputString) & " 00 00 00;"
        ElseIf HasValidM_Var(inputString) AndAlso Strings.InStr(inputString, " mod") > 1 AndAlso Strings.InStr(inputString, " flo") > 1 Then 'added in ver 5.3
            SaveConvert = "10 81 15 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf HasValidM_Var(inputString) AndAlso Strings.InStr(inputString, " mod") > 1 Then
            SaveConvert = "10 81 12 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf Strings.InStr(inputString, " time") > 1 Then
            SaveConvert = "10 83 00 00 00 00 00 00;"
        Else
            If Strings.InStr(inputString, "$var") > 0 Then
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid instruction]")
            End If
            SaveConvert = "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' math convert
    Private Function MathConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim Opt As String, Smode As String
        Opt = "80" : Smode = "00"

        If Strings.InStr(inputString, "shift ") > 1 Then
            Smode = "01"
        End If

        If Strings.InStr(inputString, "math.add ") = 1 OrElse Strings.InStr(inputString, "math.+ ") = 1 Then
            Opt = "01"
        ElseIf Strings.InStr(inputString, "math.sub ") = 1 OrElse Strings.InStr(inputString, "math.- ") = 1 Then
            Opt = "02"
        ElseIf Strings.InStr(inputString, "math.mul ") = 1 OrElse Strings.InStr(inputString, "math.* ") = 1 OrElse Strings.InStr(inputString, "math.x ") = 1 Then
            Opt = "03"
        ElseIf Strings.InStr(inputString, "math.or ") = 1 OrElse Strings.InStr(inputString, "math.|| ") = 1 Then
            Opt = "04"
        ElseIf Strings.InStr(inputString, "math.xor ") = 1 OrElse Strings.InStr(inputString, "math.x| ") = 1 Then
            Opt = "05"
        ElseIf Strings.InStr(inputString, "math.and ") = 1 OrElse Strings.InStr(inputString, "math.&& ") = 1 Then
            Opt = "06"
        ElseIf Strings.InStr(inputString, "math.shiftl ") = 1 OrElse Strings.InStr(inputString, "math.<< ") = 1 Then
            Opt = "07"
        ElseIf Strings.InStr(inputString, "math.shiftr ") = 1 OrElse Strings.InStr(inputString, "math.>> ") = 1 Then
            Opt = "08"
        ElseIf Strings.InStr(inputString, "math.div ") = 1 OrElse Strings.InStr(inputString, "math./ ") = 1 Then
            Opt = "09"
        ElseIf Strings.InStr(inputString, "math.mod ") = 1 OrElse Strings.InStr(inputString, "math.% ") = 1 Then
            Opt = "0A"
        End If

        If Opt = "80" Then
            'not math function
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid math instruction]")
            MathConvert = "01 01 00 00 00 00 00 00;"
        Else
            If X0.HasValidVar(inputString) Then
                MathConvert = "10 " & Opt & " 0A " & Smode & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            ElseIf HasValidC_Var(inputString) Then
                MathConvert = "10 " & Opt & " 0A " & Smode & " " & Grab_Var_2(inputString) & " 00 00 00;"
            Else
                If Strings.InStr(inputString, "$var=") > 4 Then
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                    MathConvert = "01 01 00 00 00 00 00 00;"
                Else
                    If Opt = "06" Then
                        MathConvert = "10 " & Opt & " 01 " & Smode & " " & X0.Grab_value_1(inputString, "AND") & ";"
                    Else
                        MathConvert = "10 " & Opt & " 01 " & Smode & " " & X0.Grab_value_1(inputString, "OR") & ";"
                    End If
                End If
            End If
        End If
    End Function

    ' dim flow convert
    Private Function DimFlowConvert(inputString As String, thisLine As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim Dmode As String
        Dmode = "01 01"

        If Strings.InStr(inputString, ".start ") > 0 OrElse Strings.InStr(inputString, ".onstart ") > 0 OrElse Strings.InStr(inputString, ".starton ") > 0 Then
            Dmode = "5a 0A"
        ElseIf Strings.InStr(inputString, ".getdir") > 0 OrElse Strings.InStr(inputString, ".dirget ") > 0 Then
            Dmode = "5e 0A"
        ElseIf Strings.InStr(inputString, ".setdir") > 0 OrElse Strings.InStr(inputString, ".dirset ") > 0 Then
            Dmode = "5d 0A"
        End If

        If Dmode <> "01 01" Then
            If X0.HasValidVar(inputString) Then
                Return Dmode & " 00 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            ElseIf HasValidC_Var(inputString) Then
                Return Dmode & " 00 00 " & Grab_Var_2(inputString) & " 00 00 00;"
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                Return "01 01 00 00 00 00 00 00;"
            End If
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid instruction]")
            Return "01 01 00 00 00 00 00 00;"
        End If
    End Function

    '------------------------------------------------------





    ' is other var direct
    Private Function IsOtherVarDirect(inputString As String) As Boolean
        If inputString.Contains("."c) Then
            Dim VarArray() As String, tempS As String
            VarArray = inputString.Split("."c)
            tempS = "load $var=" & VarArray(0) 'create a simple load instruction with var

            If X0.HasValidVar(tempS) OrElse HasValidC_Var(tempS) OrElse HasValidM_Var(tempS) Then
                Return True
            End If

        End If

        Return False
    End Function

    'custom var
    Private Function HasValidC_Var(inputString As String) As Boolean 'Custom var, moified in ver 5.0

        If C_VarListNotEmpty() Then
            Dim LArray_T2() As String, V As String, i As Integer
            V = ""
            LArray_T2 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If LArray_T2(i).Replace(" "c, "").Contains("var=") Then
                    V = LArray_T2(i).Replace("var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
                'Dim rngVar As Excel.Range
                'For Each rngVar In xlCustomVarSheet.Range(xlCustomVarSheet.Range("B3").Offset(1, 0), xlCustomVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
                '    If V = Strings.Trim(Strings.LCase(rngVar.Text)) orelse V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
                '        hasValidC_Var = True
                '        Exit Function
                '    End If
                'Next rngVar
                If CustomVarDictionary.ContainsKey(V) OrElse CustomVarDictionary.ContainsValue(V) Then
                    Return True
                End If
            End If
        End If

        Return False
    End Function

    Private Function C_VarListNotEmpty() As Boolean

        If Not IsNothing(xlCustomVarSheet) Then
            'If Strings.Len(xlCustomVarSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
            '    C_VarListNotEmpty = True
            'End If
            If CustomVarDictionary.Count > 0 Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Function Grab_Var_2(inputString As String) As String 'Custom var
        Dim V As String, LArray_T2() As String, P As String, i As Integer
        'Dim rngVar As Excel.Range
        V = "" : P = ""
        LArray_T2 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = 0 To LArray_T2.Length - 1
            If LArray_T2(i).Replace(" "c, "").StartsWith("var=") Then
                V = LArray_T2(i).Replace("var=", "")
                Exit For
            End If
        Next i

        'For Each rngVar In xlCustomVarSheet.Range(xlCustomVarSheet.Range("B3").Offset(1, 0), xlCustomVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
        '    If V = Strings.Trim(Strings.LCase(rngVar.Text)) orelse V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
        '        P = rngVar.Text
        '        Exit For
        '    End If
        'Next rngVar

        If CustomVarDictionary.ContainsKey(V) Then
            P = CustomVarDictionary(V)
        ElseIf CustomVarDictionary.ContainsValue(V) Then
            P = V
        End If

        Grab_Var_2 = Strings.Right("00" & P, 2)
    End Function

    Private Sub Populate_C_Var_Dictionary()
        If Strings.Len(xlCustomVarSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
            Dim rngVar As Excel.Range
            For Each rngVar In xlCustomVarSheet.Range(xlCustomVarSheet.Range("B3").Offset(1, 0), xlCustomVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
                CustomVarDictionary(Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text))) = Strings.Trim(Strings.LCase(rngVar.Text))
            Next rngVar
        End If
    End Sub

    ' modbus var
    Private Function HasValidM_Var(inputString As String) As Boolean 'Modbus var, moified in ver 5.0

        If M_VarListNotEmpty() Then
            Dim LArray_T2() As String, V As String, i As Integer
            V = ""
            LArray_T2 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
            For i = 0 To LArray_T2.Length - 1
                If LArray_T2(i).Replace(" "c, "").Contains("var=") Then
                    V = LArray_T2(i).Replace("var=", "")
                    Exit For
                End If
            Next i

            If V <> "" Then
                'Dim rngVar As Excel.Range
                'For Each rngVar In xlModbusVarSheet.Range(xlModbusVarSheet.Range("B3").Offset(1, 0), xlModbusVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
                '    If V = Strings.Trim(Strings.LCase(rngVar.Text)) orelse V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
                '        hasValidM_Var = True
                '        Exit Function
                '    End If
                'Next rngVar
                If ModbusVarDictionary.ContainsKey(V) OrElse ModbusVarDictionary.ContainsValue(V) Then
                    Return True
                End If
            End If
        End If

        Return False
    End Function

    Private Function M_VarListNotEmpty() As Boolean

        If Not IsNothing(xlModbusVarSheet) Then

            If ModbusVarDictionary.Count > 0 Then
                Return True
            End If
        End If

        Return False
    End Function

    Private Function Grab_Var_3(inputString As String) As String 'Modbus var
        Dim V As String, LArray_T2() As String, P As String
        'Dim rngVar As Excel.Range

        V = "" : P = ""
        LArray_T2 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)

        For i = LBound(LArray_T2) To UBound(LArray_T2)
            If Strings.InStr(Strings.Replace(LArray_T2(i), " "c, ""), "var=") Then
                V = Strings.Replace(LArray_T2(i), "var=", "")
                Exit For
            End If
        Next i

        'For Each rngVar In xlModbusVarSheet.Range(xlModbusVarSheet.Range("B3").Offset(1, 0), xlModbusVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
        '    If V = Strings.Trim(Strings.LCase(rngVar.Text)) orelse V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
        '        P = rngVar.Text
        '        Exit For
        '    End If
        'Next rngVar

        If ModbusVarDictionary.ContainsKey(V) Then
            P = ModbusVarDictionary(V)
        ElseIf ModbusVarDictionary.ContainsValue(V) Then
            P = V
        End If

        Grab_Var_3 = Strings.Right("00" & P, 2)
    End Function

    Private Sub Populate_M_Var_Dictionary()
        If Strings.Len(xlModbusVarSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
            Dim rngVar As Excel.Range
            For Each rngVar In xlModbusVarSheet.Range(xlModbusVarSheet.Range("B3").Offset(1, 0), xlModbusVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
                ModbusVarDictionary(Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text))) = Strings.Trim(Strings.LCase(rngVar.Text))
            Next rngVar
        End If
    End Sub

    ' grabbers core 1
    Private Function GrabIOEXPDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, LArray() As String, ioName As String, L As Integer
        Dim vGrabIOEXPDirect As New List(Of String)

        'detect ioexp
        If Strings.Mid(inputString, Strings.InStr(inputString, "ioexp") + 5, 1) = "." Then
            ioName = "ioexp1"
        ElseIf Strings.Mid(inputString, Strings.InStr(inputString, "ioexp") + 6, 1) = "." Then
            ioName = "ioexp" & Strings.Mid(inputString, Strings.InStr(inputString, "ioexp") + 5, 1)
        Else
            ioName = "ioexp" & Strings.Mid(inputString, Strings.InStr(inputString, "ioexp") + 5, 2)
        End If

        If Strings.InStr(inputString, "read") = (Strings.InStr(inputString, ".") + 1) Then
            If Strings.InStr(inputString, " relay") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=4 $len=2")
            ElseIf Strings.InStr(inputString, " aout8") > 9 OrElse Strings.InStr(inputString, " analog8") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=43 $len=2")
            ElseIf Strings.InStr(inputString, " aout7") > 9 OrElse Strings.InStr(inputString, " analog7") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=39 $len=2")
            ElseIf Strings.InStr(inputString, " aout6") > 9 OrElse Strings.InStr(inputString, " analog6") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=35 $len=2")
            ElseIf Strings.InStr(inputString, " aout5") > 9 OrElse Strings.InStr(inputString, " analog5") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=31 $len=2")
            ElseIf Strings.InStr(inputString, " aout4") > 9 OrElse Strings.InStr(inputString, " analog4") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=27 $len=2")
            ElseIf Strings.InStr(inputString, " aout3") > 9 OrElse Strings.InStr(inputString, " analog3") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=23 $len=2")
            ElseIf Strings.InStr(inputString, " aout2") > 9 OrElse Strings.InStr(inputString, " analog2") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=19 $len=2")
            ElseIf Strings.InStr(inputString, " aout") > 9 OrElse Strings.InStr(inputString, " analog") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=15 $len=2")
            ElseIf Strings.InStr(inputString, " input") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=3 $len=2")
            ElseIf Strings.InStr(inputString, " pt100 2") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=9 $len=2")
            ElseIf Strings.InStr(inputString, " pt100") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=7 $len=2")
            Else
                vGrabIOEXPDirect.Add(inputString)

                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid instruction]")
            End If
        Else

            ioString = inputString.Replace(" "c, "").Replace("#"c, "$")
            ioString = Strings.Mid(ioString, 1, Strings.InStr(ioString, "$"c) - 1)
            LArray = ioString.Split("."c)

            If UBound(LArray) > 0 Then

                vGrabIOEXPDirect.Add("load $var=" & ioName)

                For L = LBound(LArray) + 1 To UBound(LArray)
                    If (LArray(L) = "set" OrElse LArray(L) = "or") AndAlso X0.Grab_value_0(inputString, L) <> "NULL" Then
                        vGrabIOEXPDirect.Add("math.or $" & X0.Grab_value_0(inputString, L))
                    ElseIf (LArray(L) = "unset" OrElse LArray(L) = "and") AndAlso X0.Grab_value_0(inputString, L) <> "NULL" Then
                        vGrabIOEXPDirect.Add("math.and $" & X0.Grab_value_0(inputString, L))
                    ElseIf (LArray(L) = "xor") AndAlso X0.Grab_value_0(inputString, L) <> "NULL" Then
                        vGrabIOEXPDirect.Add("math.xor $" & X0.Grab_value_0(inputString, L))
                    End If
                Next L

                vGrabIOEXPDirect.Add("send.mod $dev=" & ioName & " $reg=4 $len=2")
                vGrabIOEXPDirect.Add("save $var=" & ioName)
            Else
                vGrabIOEXPDirect.Add("nop")
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IOEXP instruction]")
            End If

        End If

        Return vGrabIOEXPDirect
    End Function

    Private Function GrabIODEXPDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioName As String
        Dim vGrabIODEXPDirect As New List(Of String)

        'detect ioexp
        If Strings.Mid(inputString, Strings.InStr(inputString, "iodexp") + 6, 1) = "." Then
            ioName = "iodexp1"
        Else
            ioName = "iodexp" & Strings.Mid(inputString, Strings.InStr(inputString, "iodexp") + 6, 1)
        End If

        If Strings.InStr(inputString, "read") = (Strings.InStr(inputString, ".") + 1) Then
            If Strings.InStr(inputString, " aout8") > 9 OrElse Strings.InStr(inputString, " analog8") > 9 OrElse Strings.InStr(inputString, " ch8") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=34 $len=2")
            ElseIf Strings.InStr(inputString, " aout7") > 9 OrElse Strings.InStr(inputString, " analog7") > 9 OrElse Strings.InStr(inputString, " ch7") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=30 $len=2")
            ElseIf Strings.InStr(inputString, " aout6") > 9 OrElse Strings.InStr(inputString, " analog6") > 9 OrElse Strings.InStr(inputString, " ch6") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=26 $len=2")
            ElseIf Strings.InStr(inputString, " aout5") > 9 OrElse Strings.InStr(inputString, " analog5") > 9 OrElse Strings.InStr(inputString, " ch5") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=22 $len=2")
            ElseIf Strings.InStr(inputString, " aout4") > 9 OrElse Strings.InStr(inputString, " analog4") > 9 OrElse Strings.InStr(inputString, " ch4") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=18 $len=2")
            ElseIf Strings.InStr(inputString, " aout3") > 9 OrElse Strings.InStr(inputString, " analog3") > 9 OrElse Strings.InStr(inputString, " ch3") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=14 $len=2")
            ElseIf Strings.InStr(inputString, " aout2") > 9 OrElse Strings.InStr(inputString, " analog2") > 9 OrElse Strings.InStr(inputString, " ch2") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=10 $len=2")
            ElseIf Strings.InStr(inputString, " aout") > 9 OrElse Strings.InStr(inputString, " analog") > 9 OrElse Strings.InStr(inputString, " ch") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=6 $len=2")
            Else
                vGrabIODEXPDirect.Add("nop")
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IODEXP instruction]")
            End If
        Else
            vGrabIODEXPDirect(1) = "nop"
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IODEXP instruction]")
        End If

        Return vGrabIODEXPDirect
    End Function

    Private Function GrabRCUDirect(inputString As String) As List(Of String)
        Dim LArray() As String, LArray1() As String
        Dim L As Integer
        LArray = inputString.Replace("#"c, "").Split(New String() {" $"}, StringSplitOptions.None)
        LArray1 = LArray(0).Split("."c)

        Dim vGrabRCUDirect As New List(Of String)

        For L = 0 To UBound(LArray1) - 1
            vGrabRCUDirect.Add("rcu." & LArray1(L + 1) & " $" & X0.Grab_value_0(inputString, L + 1))
        Next L

        Return vGrabRCUDirect
    End Function

    Private Function GrabMathDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)

        Dim LArray() As String, LArray1() As String
        Dim L As Integer
        LArray = inputString.Replace("#"c, "").Split(New String() {" $"}, StringSplitOptions.None)
        LArray1 = LArray(0).Split("."c)

        Dim vGrabMathDirect As New List(Of String)

        For L = 0 To UBound(LArray1) - 1
            If Not X0.Grab_value_0(inputString, L + 1) = "NULL" Then
                vGrabMathDirect.Add("math." & LArray1(L + 1) & " $" & X0.Grab_value_0(inputString, L + 1))
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & " at """ & LArray1(L + 1) & """][Invalid math data type]")
                vGrabMathDirect.Add("math." & LArray1(L + 1))
            End If
        Next L

        Return vGrabMathDirect
    End Function

    Private Function GrabSaveDirect(inputString As String) As List(Of String)
        Dim ioString As String, LArray() As String, LArray1() As String, ioLen As String
        Dim vGrabSaveDirect As New List(Of String)

        ioString = "" : ioLen = "2"

        If Strings.InStr(inputString, " $var=") > 3 AndAlso Strings.InStr(inputString, " $len=") > 3 Then
            'added in v5.3 specially for save settings with length
            'can inprove this entire function later
            LArray = inputString.Split(New String() {" $"}, StringSplitOptions.None) 'filter with space

            For i = LBound(LArray) To UBound(LArray)
                If Strings.InStr(LArray(i), "var=") = 1 Then
                    ioString = Strings.Replace(LArray(i), "var=", "")
                End If
                If Strings.InStr(LArray(i), "len=") = 1 Then
                    ioLen = Strings.Replace(LArray(i), "len=", "")
                End If
            Next i
            LArray1 = ioString.Split(","c)

            For i = LBound(LArray1) To UBound(LArray1)
                vGrabSaveDirect.Add(LArray(0) & " $var=" & LArray1(i) & " $len=" & ioLen)
            Next i
        ElseIf Strings.InStr(inputString, " $var=") > 3 Then
            LArray = inputString.Split("$"c)

            For i = LBound(LArray) To UBound(LArray)
                If Strings.InStr(LArray(i), "var=") = 1 Then
                    ioString = Strings.Replace(LArray(i), "var=", "")
                End If
            Next i
            LArray1 = ioString.Split(","c)

            For i = LBound(LArray1) To UBound(LArray1)
                vGrabSaveDirect.Add(LArray(0) & "$var=" & LArray1(i))
            Next i
        Else
            vGrabSaveDirect.Add(inputString)
        End If

        Return vGrabSaveDirect
    End Function

    Private Function GrabOtherVarDirect(inputString As String) As List(Of String)
        Dim LArray() As String, LArray1() As String
        Dim OtherVar As String, L As Integer, Smode As String
        Dim vGrabOtherVarDirect As New List(Of String)

        LArray = inputString.Replace("#"c, "").Split(New String() {" $"}, StringSplitOptions.None)
        LArray1 = LArray(0).Split("."c)
        OtherVar = LArray1(0)
        Smode = ""

        If HasValidM_Var("load $var=" & OtherVar) Then
            Smode = " modbus"
        End If

        If LArray1(1) = "set" Then
            If HasValidM_Var(inputString) Then
                vGrabOtherVarDirect.Add("load modbus $var=" & X0.Grab_variant(inputString, "var"))
            Else
                vGrabOtherVarDirect.Add("load $" & X0.Grab_value_0(inputString, 1))
            End If
            vGrabOtherVarDirect.Add("save" & Smode & " $var=" & OtherVar)
        ElseIf LArray1(1) = "unset" Then
            vGrabOtherVarDirect.Add("load $hex=0")
            vGrabOtherVarDirect.Add("save" & Smode & " $var=" & OtherVar)
        Else
            vGrabOtherVarDirect.Add("load" & Smode & " $var=" & OtherVar)
            For L = 1 To UBound(LArray1)
                If X0.Grab_value_0(inputString, L) <> "NULL" Then
                    vGrabOtherVarDirect.Add("math." & LArray1(L) & " $" & X0.Grab_value_0(inputString, L))
                End If
            Next L
            vGrabOtherVarDirect.Add("save" & Smode & " $var=" & OtherVar)
        End If

        Return vGrabOtherVarDirect
    End Function

    Private Function GrabModbusDirect(inputString As String) As List(Of String)
        Dim len_check As String, val_check As String
        Dim vGrabModbusDirect As New List(Of String)

        len_check = X0.Grab_variant(inputString, "len")
        val_check = X0.Grab_value_0(inputString, 1)

        If len_check <> "NULL" AndAlso val_check <> "NULL" Then
            vGrabModbusDirect.Add("load $" & val_check)
            vGrabModbusDirect.Add(Strings.Replace(inputString, " $" & val_check, ""))
        ElseIf len_check = "NULL" AndAlso val_check = "NULL" Then
            vGrabModbusDirect.Add(Strings.Trim(inputString & " $len=2"))
        Else
            vGrabModbusDirect.Add(inputString)
        End If

        GrabModbusDirect = vGrabModbusDirect
    End Function

    Private Function GrabDimmerDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, chString As String, LArray() As String, LArray1() As String, LArray2() As String, i As Integer
        Dim thisIOdev As String, thisChReg As Integer, ioLen As Integer

        thisIOdev = "ioexp1" : ioString = "" : chString = ""
        ioLen = 6
        LArray = inputString.Split(New String() {" $"}, StringSplitOptions.None)

        For i = LBound(LArray) To UBound(LArray)
            If Strings.InStr(LArray(i), "dev=") = 1 Then
                ioString = LArray(i).Replace("dev=", "")
            ElseIf Strings.InStr(LArray(i), "ch=") = 1 Then
                chString = LArray(i).Replace("ch=", "")
            End If
        Next i

        Dim vGrabDimmerDirect As New List(Of String)


        If Strings.Len(ioString) > 0 AndAlso Strings.Len(chString) > 0 Then

            LArray1 = ioString.Split(","c)
            LArray2 = X0.CommaExpand(chString, 1, 16).Split(","c) 'added in 5.1

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
                If Strings.Len(LArray1(i)) > 1 AndAlso IsNumeric(Strings.Right(LArray1(i), 1)) Then
                    thisIOdev = LArray1(i)
                ElseIf Strings.Len(LArray1(i)) > 1 AndAlso Not IsNumeric(Strings.Right(LArray1(i), 1)) Then
                    thisIOdev = LArray1(i) & "1"
                ElseIf Strings.Len(LArray1(i)) = 1 AndAlso IsNumeric(Strings.Right(LArray1(i), 1)) Then
                    thisIOdev = Strings.Mid(thisIOdev, 1, Strings.Len(thisIOdev) - 1) & LArray1(i)
                End If

                If IsNumeric(LArray2(i)) Then
                    If CInt(LArray2(i)) > 0 AndAlso CInt(LArray2(i)) < 9 Then
                        thisChReg = CInt(LArray2(i))
                    Else
                        thisChReg = 1
                    End If
                Else
                    thisChReg = 1
                End If

                If Strings.InStr(thisIOdev, "ioe") = 1 Then
                    thisChReg = 12 + 4 * (thisChReg - 1)
                Else
                    thisChReg = 3 + 4 * (thisChReg - 1)
                End If

                If Strings.InStr(inputString, " stop") > 1 Then
                    thisChReg += 2
                    ioLen = 2
                End If


                vGrabDimmerDirect.Add("send.modbus $dev=" & thisIOdev & " $reg=" & thisChReg & " $len=" & ioLen)
            Next i
        Else
            vGrabDimmerDirect.Add("nop")
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid dimmer instruction]")
        End If
        GrabDimmerDirect = vGrabDimmerDirect
    End Function

    Private Function GrabDimmerReadDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, chString As String, LArray() As String, LArray1() As String, LArray2() As String
        Dim thisIOdev As String, thisChReg As Integer, ioLen As Integer, i As Integer

        thisIOdev = "ioexp1" : ioString = "" : chString = ""
        ioLen = 2
        LArray = inputString.Split(New String() {" $"}, StringSplitOptions.None)

        For i = LBound(LArray) To UBound(LArray)
            If Strings.InStr(LArray(i), "dev=") = 1 Then
                ioString = Strings.Replace(LArray(i), "dev=", "")
            ElseIf Strings.InStr(LArray(i), "ch=") = 1 Then
                chString = Strings.Replace(LArray(i), "ch=", "")
            End If
        Next i

        Dim vGrabDimmerReadDirect As New List(Of String)


        If Strings.Len(ioString) > 0 AndAlso Strings.Len(chString) > 0 Then

            LArray1 = ioString.Split(","c)
            LArray2 = X0.CommaExpand(chString, 1, 16).Split(","c)

            If Strings.Len(LArray1(0)) > 1 AndAlso IsNumeric(Strings.Right(LArray1(0), 1)) Then
                thisIOdev = LArray1(0)
            ElseIf Strings.Len(LArray1(0)) > 1 AndAlso Not IsNumeric(Strings.Right(LArray1(0), 1)) Then
                thisIOdev = LArray1(0) & "1"
            ElseIf Strings.Len(LArray1(0)) = 1 AndAlso IsNumeric(Strings.Right(LArray1(0), 1)) Then
                thisIOdev = Strings.Mid(thisIOdev, 1, Strings.Len(thisIOdev) - 1) & LArray1(0)
            End If

            If IsNumeric(LArray2(0)) Then
                If CInt(LArray2(0)) > 0 AndAlso CInt(LArray2(0)) < 9 Then
                    thisChReg = CInt(LArray2(0))
                Else
                    thisChReg = 1
                End If
            Else
                thisChReg = 1
            End If

            If Strings.InStr(thisIOdev, "ioe") = 1 Then
                thisChReg = 15 + 4 * (thisChReg - 1)
            Else
                thisChReg = 6 + 4 * (thisChReg - 1)
            End If

            vGrabDimmerReadDirect.Add("read.modbus $dev=" & thisIOdev & " $reg=" & thisChReg & " $len=" & ioLen)

            If (UBound(LArray1) - LBound(LArray1) + 1) > 1 Then
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Reading from muntiple devices is not practical]")
            End If
            If (UBound(LArray2) - LBound(LArray2) + 1) > 1 Then
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Reading from muntiple channels is not practical]")
            End If

        Else
            vGrabDimmerReadDirect.Add("nop")
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid dimmer instruction]")
        End If

        GrabDimmerReadDirect = vGrabDimmerReadDirect
    End Function

    Private Function GrabDaliDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, LArray1() As String, ioName As String, i As Integer, paraString As String
        Dim DaliBus As String, sendReg As String, ioBlst As String, ioGp As String
        Dim idFactor As Integer, bgLum As String, bgID As String
        Dim dtr As Integer, isQuery As Boolean, sceneNum As Integer, readReg As String

        ioString = Strings.Replace(inputString, "#"c, "$"c)
        If InStr(ioString, "$") > 4 Then
            ioString = Mid(ioString, 1, InStr(ioString, "$"c) - 1) & " "
            paraString = X0.GetAllParameters(inputString)
            DaliBus = X0.Grab_variant(paraString, "bus")
            ioBlst = X0.Grab_variant(paraString, "ballast")
            ioGp = X0.Grab_variant(paraString, "group")
        Else
            ioString = ioString & " "
            DaliBus = "NULL"
            ioBlst = "NULL"
            ioGp = "NULL"
        End If

        idFactor = 1
        isQuery = False

        If Strings.Mid(inputString, Strings.InStr(inputString, "dali") + 4, 1) = "." Then
            ioName = "dali1"
        Else
            ioName = "dali" & Strings.Mid(inputString, Strings.InStr(inputString, "dali") + 4, 1)
        End If

        'default when bus not mentioned its bus1
        sendReg = "06"
        readReg = "04"

        If DaliBus <> "NULL" AndAlso IsNumeric(DaliBus) Then
            If CInt(DaliBus) = 2 Then
                sendReg = "05"
                readReg = "03"
            End If
        End If

        Dim vGrabDaliDirect As New List(Of String)

        If InStr(ioString, ".read ") > 0 Then
            isQuery = True

            Select Case True
                Case InStr(inputString & " ", " arc ") > 0
                    bgLum = "A0"
                Case InStr(inputString & " ", " status ") > 0
                    bgLum = "90"
                Case InStr(inputString & " ", " ispresent ") > 0
                    bgLum = "91"
                Case InStr(inputString & " ", " isfail ") > 0
                    bgLum = "92"
                Case InStr(inputString & " ", " ison ") > 0
                    bgLum = "93"
                Case InStr(inputString & " ", " dtr ") > 0, InStr(inputString & " ", " dtr0 ") > 0
                    bgLum = "98"
                Case InStr(inputString & " ", " dtr1 ") > 0
                    bgLum = "9C"
                Case InStr(inputString & " ", " dtr2 ") > 0
                    bgLum = "9D"
                Case Else
                    bgLum = "A0"
            End Select

        ElseIf InStr(ioString, ".dtr") > 0 Then
            dtr = X0.GetIndexFromBase(ioString, ".dtr")
            If dtr >= 0 AndAlso dtr <= 2 Then
                'attempt to load value to dtr0, dtr1 or dtr2
                If dtr = 1 Then
                    bgID = "C3"
                ElseIf dtr = 2 Then
                    bgID = "C5"
                Else
                    bgID = "A3"
                End If

                If InStr(inputString, "$hex=") > 0 OrElse InStr(inputString, "$dec=") > 0 OrElse InStr(inputString, "$lum=") > 0 Then
                    bgLum = X0.Grab_value_7(inputString)
                    vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & sendReg & " $hex=" & bgID & bgLum)
                    GrabDaliDirect = vGrabDaliDirect
                    Exit Function
                ElseIf dtr = 0 AndAlso InStr(inputString, " from ") > 0 Then
                    'special case to dtr0 from other things
                    If InStr(inputString & " ", " arc ") > 0 Then
                        'load dtr0 from arc
                        bgLum = "21"
                    Else
                        ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                        vGrabDaliDirect.Add("nop")
                        GrabDaliDirect = vGrabDaliDirect
                        Exit Function
                    End If
                Else
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                    vGrabDaliDirect.Add("nop")
                    GrabDaliDirect = vGrabDaliDirect
                    Exit Function
                End If

            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                vGrabDaliDirect.Add("nop")
                GrabDaliDirect = vGrabDaliDirect
                Exit Function
            End If

        ElseIf InStr(ioString, ".scene") > 0 Then
            sceneNum = X0.GetIndexFromBase(ioString, ".scene")

            If sceneNum >= 0 AndAlso sceneNum <= 15 Then
                If InStr(inputString, " from ") > 0 Then
                    'set scene x from dtr 0
                    If InStr(inputString & " ", " dtr ") > 0 OrElse InStr(inputString & " ", " dtr0 ") > 0 Then
                        bgLum = (64 + sceneNum).ToString("X2")
                    Else
                        ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                        vGrabDaliDirect.Add("nop")
                        GrabDaliDirect = vGrabDaliDirect
                        Exit Function
                    End If
                Else
                    'set specific scene to ballast, group or broadcast
                    bgLum = (16 + sceneNum).ToString("X2")
                End If
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                vGrabDaliDirect.Add("nop")
                GrabDaliDirect = vGrabDaliDirect
                Exit Function
            End If

        ElseIf InStr(inputString, " from ") > 0 Then
            'set other things from dtr0
            Select Case True
                Case InStr(ioString, ".max ") > 0
                    bgLum = "2A"
                Case InStr(ioString, ".min ") > 0
                    bgLum = "2B"
                Case InStr(ioString, ".fail ") > 0
                    bgLum = "2C"
                Case InStr(ioString, ".power ") > 0, InStr(ioString, ".poweron ") > 0
                    bgLum = "2D"
                Case InStr(ioString, ".fadetime ") > 0
                    bgLum = "2E"
                Case InStr(ioString, ".faderate ") > 0
                    bgLum = "2F"
                Case Else
                    bgLum = "2A"
            End Select
        Else
            Select Case True
                Case InStr(ioString, ".on ") > 0, InStr(ioString, ".max ") > 0
                    bgLum = "05"
                Case InStr(ioString, ".off ") > 0
                    bgLum = "00"
                Case InStr(ioString, ".ramp ") > 0
                    idFactor = 0
                    bgLum = X0.Grab_value_5(inputString)
                Case InStr(ioString, ".stop ") > 0
                    idFactor = 0
                    bgLum = "FF"
                Case InStr(ioString, ".min ") > 0
                    bgLum = "06"
                Case InStr(ioString, ".up ") > 0
                    bgLum = "01"
                Case InStr(ioString, ".down ") > 0
                    bgLum = "02"
                Case InStr(ioString, ".stepup ") > 0
                    bgLum = "03"
                Case InStr(ioString, ".stepdown ") > 0
                    bgLum = "04"
                Case InStr(ioString, ".downoff ") > 0
                    bgLum = "07"
                Case InStr(ioString, ".onup ") > 0
                    bgLum = "08"
                Case Else
                    'taken as ramp same as instr(iostring, ".ramp") > 0
                    idFactor = 0
                    bgLum = X0.Grab_value_5(inputString)
            End Select
        End If


        If ioBlst <> "NULL" OrElse ioGp <> "NULL" Then

            If ioBlst <> "NULL" Then

                ioBlst = X0.CommaExpand(ioBlst, 0, 63)
                LArray1 = ioBlst.Split(","c)

                If Not isQuery Then
                    For i = LBound(LArray1) To UBound(LArray1)
                        If IsNumeric(LArray1(i)) Then
                            If CInt(LArray1(i)) > -1 AndAlso CInt(LArray1(i)) < 64 Then

                                bgID = (CInt(LArray1(i)) * 2 + idFactor).ToString("X2")
                                vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & sendReg & " $hex=" & bgID & bgLum)
                            End If
                        End If
                    Next i

                Else
                    'is query only one ballast allowed to query from
                    If (UBound(LArray1) - LBound(LArray1) + 1) = 1 Then
                        bgID = (CInt(LArray1(0)) * 2 + idFactor).ToString("X2")
                        vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & sendReg & " $hex=" & bgID & bgLum)
                        vGrabDaliDirect.Add("wait $sec=0.005")
                        vGrabDaliDirect.Add("read.modbus $dev=" & ioName & " $reg=" & readReg)
                        GrabDaliDirect = vGrabDaliDirect
                        Exit Function
                    Else
                        ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                        vGrabDaliDirect.Add("nop")
                        GrabDaliDirect = vGrabDaliDirect
                        Exit Function
                    End If
                End If


            End If

            If ioGp <> "NULL" Then
                ioGp = X0.CommaExpand(ioGp, 0, 15)
                LArray1 = ioGp.Split(","c)

                If Not isQuery Then
                    For i = LBound(LArray1) To UBound(LArray1)
                        If IsNumeric(LArray1(i)) Then
                            If CInt(LArray1(i)) > -1 AndAlso CInt(LArray1(i)) < 16 Then
                                bgID = (CInt(LArray1(i)) * 2 + idFactor + 128).ToString("X2")
                                vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & sendReg & " $hex=" & bgID & bgLum)
                            End If
                        End If
                    Next i
                Else
                    'is query one one ballast allowed to query from
                    If (UBound(LArray1) - LBound(LArray1) + 1) = 1 Then
                        bgID = (CInt(LArray1(0)) * 2 + idFactor + 128).ToString("X2")
                        vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & sendReg & " $hex=" & bgID & bgLum)
                        vGrabDaliDirect.Add("wait $sec=0.005")
                        vGrabDaliDirect.Add("read.modbus $dev=" & ioName & " $reg=" & readReg)
                        GrabDaliDirect = vGrabDaliDirect
                        Exit Function
                    Else
                        ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                        vGrabDaliDirect.Add("nop")
                        GrabDaliDirect = vGrabDaliDirect
                        Exit Function
                    End If
                End If

            End If

        ElseIf ioBlst = "NULL" AndAlso ioGp = "NULL" Then
            'can be broadcast
            If Not isQuery Then
                vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & sendReg & " $hex=" & "FF" & bgLum)
            Else
                'query cannot be broadcast
                ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
                vGrabDaliDirect.Add("nop")
            End If
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
            vGrabDaliDirect.Add("nop")
        End If

        GrabDaliDirect = vGrabDaliDirect
    End Function

    Private Function GrabDMXDirect(inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioChString As String, ioChExp As String, LArray1() As String, i As Integer

        ioChString = X0.Grab_variant(inputString, "ch")

        Dim vGrabDMXDirect As New List(Of String)

        If ioChString <> "NULL" Then
            ioChExp = X0.CommaExpand(ioChString, 0, 63)
            LArray1 = ioChExp.Split(","c)

            For i = LBound(LArray1) To UBound(LArray1)
                If IsNumeric(LArray1(i)) Then
                    If CInt(LArray1(i)) > -1 AndAlso CInt(LArray1(i)) < 64 Then
                        vGrabDMXDirect.Add(Strings.Replace(inputString, "$ch=" & ioChString, "$ch=" & LArray1(i)))
                    End If
                End If
            Next i

        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DMX instruction]")
            vGrabDMXDirect.Add("nop")
        End If

        Return vGrabDMXDirect
    End Function

    ' grabbers core 2

    'inteligent Device Grabber
    'main function
    Private Function Grab_Dev_1(inputString As String, thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim P As Integer, i As Integer
        Dim LArray_T1() As String
        Dim LArray_GP() As String
        P = 0
        LArray_T1 = inputString.Replace(" "c, "").Split("$"c)

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(LArray_T1(i), "dev=") = 1 Then
                LArray_GP = LArray_T1(i).Split("="c)
                If IsNumeric(LArray_GP(1)) Then
                    If CInt(LArray_GP(1)) > 0 AndAlso CInt(LArray_GP(1)) < (GetDevLimit() - 2) Then ' 21 Then flexible dev
                        P = CInt(LArray_GP(1))
                    End If
                ElseIf (LArray_GP(1) = "server" OrElse LArray_GP(1) = "db" OrElse LArray_GP(1) = "ff") Then
                    P = 255
                ElseIf (LArray_GP(1) = "last" OrElse LArray_GP(1) = "previous" OrElse LArray_GP(1) = "fe") Then
                    P = 254
                ElseIf (Not IsNumeric(LArray_GP(1))) AndAlso LArray_GP(0) = "dev" Then
                    P = Grab_ID_from_Dev_Name(LArray_GP(1), thisRCU, ErrorWarnLog)
                End If
            End If
        Next i

        Return P.ToString("X2")
    End Function

    Private Function Grab_ID_from_Dev_Name(inputString As String, thisRCU As String, ByRef ErrorWarnLog As String()) As Integer
        Grab_ID_from_Dev_Name = 0
        'checking device validity
        Dim RcuDevList() As String, RcuDevNameList() As String, RCUDevType As String
        Dim DevEnumID As String, i As Integer

        RcuDevList = "01,02,03,04,04,05,05,06,06,07,07,08,0a".Split(","c) 'should be in lcase for comparison
        RcuDevNameList = "idpg,tig,tag,ioexp,ioe,iodexp,iod,gsw,gs,ioexp,ioe,bsp,dali".Split(",")
        RCUDevType = "00"
        DevEnumID = "0X" 'priority (lower 4bits) can be any value

        For i = LBound(RcuDevList) To UBound(RcuDevList)
            If Strings.InStr(inputString, RcuDevNameList(i)) Then
                RCUDevType = RcuDevList(i)
                If IsNumeric(Strings.Replace(inputString, RcuDevNameList(i), "")) Then
                    DevEnumID = CInt(Strings.Replace(inputString, RcuDevNameList(i), "") - 1).ToString("X1") & "1"
                End If
                Exit For
            End If
        Next i

        If RCUDevType <> "00" Then
            Dim ce As Excel.Range
            For Each ce In xlDeviceSheet.Range("C4:C" & GetDevLimit())
                If RCUDevType = "04" Then
                    If (Strings.LCase(ce.Text) = "04" OrElse Strings.LCase(ce.Text) = "07") AndAlso Strings.Left(Strings.Replace(ce.Offset(0, -1).Text, " "c, ""), 1) = Strings.Left(DevEnumID, 1) Then
                        Grab_ID_from_Dev_Name = ce.Row - 3
                        Exit Function
                    End If
                Else
                    If Strings.LCase(ce.Text) = RCUDevType AndAlso Strings.Left(Strings.Replace(ce.Offset(0, -1).Text, " "c, ""), 1) = Strings.Left(DevEnumID, 1) Then
                        Grab_ID_from_Dev_Name = ce.Row - 3
                        Exit Function
                    End If
                End If
            Next ce
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | Device" & thisRCU & " | " & inputString & "][Device not found, check enum]")
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | Device" & thisRCU & " | " & inputString & "][Device type not identified]")
        End If
    End Function

    Private Function GetDevLimit() As Integer
        Dim devSize As Integer
        devSize = xlDeviceSheet.Range("B4").End(Excel.XlDirection.xlDown).Row
        If devSize > 67 Then
            devSize = 67
        End If
        GetDevLimit = devSize
    End Function


    'garbage management

    Private Sub ReleaseObject(ByRef thisObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(thisObject)
            thisObject = Nothing
        Catch ex As Exception
            thisObject = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
