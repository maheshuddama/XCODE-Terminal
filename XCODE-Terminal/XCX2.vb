Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX2
    Dim xlDeviceSheet As Excel.Worksheet
    Dim xlXCODESheet As Excel.Worksheet
    Dim xlProgramSheet As Excel.Worksheet
    Dim xlSettingsSheet As Excel.Worksheet
    Dim xlCustomVarSheet As Excel.Worksheet
    Dim xlModbusVarSheet As Excel.Worksheet
    Dim xlFUNCTIONSheet As Excel.Worksheet
    Dim xlWorkFunc As Excel.WorksheetFunction

    Dim RCUname As String

    Dim X0 As XCX0

    Sub XCODE_Core1_Core2(ByVal MemmapName As String, ByRef thisWorkBook As Excel.Workbook, ByRef thisxlWorkFunc As Excel.WorksheetFunction, ByRef thisRCUList As List(Of String), ByRef ErrorWarnLog As String())

        xlWorkFunc = thisxlWorkFunc
        X0 = New XCX0(thisxlWorkFunc)

        For Each thisRCU As String In thisRCUList

            If thisRCU = "" Then
                RCUname = ""
            Else
                RCUname = "RCU" & thisRCU & " > "
            End If

            xlDeviceSheet = thisWorkBook.Sheets("Device" & thisRCU)
            xlXCODESheet = thisWorkBook.Sheets("XCODE" & thisRCU)
            xlProgramSheet = thisWorkBook.Sheets("Program" & thisRCU)
            xlSettingsSheet = thisWorkBook.Sheets("Settings" & thisRCU)

            'X0.ClearFunctionSeparaters(xlProgramSheet) 'optional
            'adding function sheet
            If Not X0.isSheetExist(thisWorkBook, "FUNCTION" & thisRCU) Then
                thisWorkBook.Sheets.Add(After:=thisWorkBook.Sheets("Settings" & thisRCU)).Name = "FUNCTION" & thisRCU
            End If
            xlFUNCTIONSheet = thisWorkBook.Sheets("FUNCTION" & thisRCU)
            'assign CustomVar and ModbusVar sheets
            If X0.isSheetExist(thisWorkBook, "CustomVar" & thisRCU) Then
                xlCustomVarSheet = thisWorkBook.Sheets("CustomVar" & thisRCU)
            Else
                xlCustomVarSheet = Nothing
            End If
            If X0.isSheetExist(thisWorkBook, "ModbusVar" & thisRCU) Then
                xlModbusVarSheet = thisWorkBook.Sheets("ModbusVar" & thisRCU)
            Else
                xlModbusVarSheet = Nothing
            End If

            'clear XCODE
            X0.ClearProgFUN(xlProgramSheet, xlFUNCTIONSheet)

            X0.print_XCODE_build(xlXCODESheet)
            X0.Check_XCODE_format(MemmapName, xlXCODESheet, thisRCU)

            Dim ce1 As Excel.Range, ce2 As Excel.Range
            Dim lineNum As Integer, Fcount As Integer, pcount As Double, pfullcount As Double
            Dim ThisTempLine As String, ThisHexCode As String, thisString As String, thisListOfString As List(Of String)
            lineNum = 0 : Fcount = 0 : pcount = 0 : pfullcount = 0

            'XCodeProgress.Show()
            'XCodeProgress.ProgressBar_val(pcount)
            pfullcount = xlXCODESheet.Range(xlXCODESheet.Range("B2"), xlXCODESheet.Range("B2").End(Excel.XlDirection.xlDown)).Count


            For Each ce1 In xlXCODESheet.Range(xlXCODESheet.Range("B2"), xlXCODESheet.Range("B2").End(Excel.XlDirection.xlDown))

                If ce1.Offset(0, 1).Text <> "" Then
                    ThisTempLine = Strings.LCase(Strings.Trim(ce1.Offset(0, 1).Text))

                    If Strings.InStr(ThisTempLine, "//") > 2 And Not Strings.Left(ThisTempLine, 2) = "0x" Then 'comment detection and filter except direct hex
                        ThisTempLine = Strings.Trim(Strings.Left(ThisTempLine, Strings.InStr(ThisTempLine, "//") - 1))
                    End If

                    If Strings.Left(ThisTempLine, 2) = "//" Then
                        'do nothing this is a comment
                    ElseIf Strings.Left(ThisTempLine, 2) = "0x" Then
                        If X0.GrabHexString(ThisTempLine) <> "FALSE" Then
                            insertXCODE(lineNum, X0.GrabComment(ThisTempLine))
                            insertHEXCODE(lineNum, X0.GrabHexString(ThisTempLine))
                            lineNum = lineNum + 1
                        End If
                    ElseIf X0.isOtherRCUcall(ThisTempLine) Then
                        insertXCODE(lineNum, "//" & ThisTempLine)
                        insertHEXCODE(lineNum, "01 01 00 00 00 00 00 00;")
                        lineNum = lineNum + 1
                    ElseIf X0.isAFunction(ThisTempLine) Then
                        If isNotDuplicateFunction(X0.GrabFunctionName(ThisTempLine)) Then
                            insertFunction(Fcount, lineNum, X0.GrabFunctionName(ThisTempLine))
                            Fcount = Fcount + 1
                        Else
                            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & ce1.Text & " | " & X0.GrabFunctionName(ThisTempLine) & "][Duplicate Function]")
                        End If
                    ElseIf X0.isAnIOEXPDirect(ThisTempLine) Then
                        thisListOfString = GrabIOEXPDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isAnIODEXPDirect(ThisTempLine) Then
                        thisListOfString = GrabIODEXPDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isARCUDirect(ThisTempLine) Then
                        thisListOfString = GrabRCUDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isAMathDirect(ThisTempLine) Then
                        thisListOfString = GrabMathDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf isOtherVarDirect(ThisTempLine) Then
                        thisListOfString = GrabOtherVarDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isASaveDirect(ThisTempLine) Then
                        thisListOfString = GrabSaveDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isAModbusDirect(ThisTempLine) Then
                        thisListOfString = GrabModbusDirect(ThisTempLine)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isADimmerDirect(ThisTempLine) Then
                        thisListOfString = GrabDimmerDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isADimmerReadDirect(ThisTempLine) Then
                        thisListOfString = GrabDimmerReadDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isADimmerLoad(ThisTempLine) Then
                        insertXCODE(lineNum, "load $hex=1")
                        lineNum = lineNum + 1
                        insertXCODE(lineNum, "//" & ThisTempLine)
                        insertHEXCODE(lineNum, X0.LoadDimmerConvert(ThisTempLine))
                        lineNum = lineNum + 1
                    ElseIf X0.isADaliDirect(ThisTempLine) Then
                        thisListOfString = GrabDaliDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isADMXDirect(ThisTempLine) Then
                        thisListOfString = GrabDMXDirect(ThisTempLine, thisRCU, ce1.Text, ErrorWarnLog)
                        For Each thisString In thisListOfString
                            insertXCODE(lineNum, thisString)
                            lineNum = lineNum + 1
                        Next
                    ElseIf X0.isADefinition(ThisTempLine) Then
                        'rcu definitions, get info later
                    Else
                        If Strings.Len(Strings.Replace(ThisTempLine, " ", "")) < 2 Then
                            'ignore multiple space empty lines and lines less then 2 charactors
                        Else
                            insertXCODE(lineNum, ThisTempLine)
                            lineNum = lineNum + 1
                        End If
                    End If
                End If

                pcount = pcount + 1
                X0.ConsoleProgress(RCUname & "XCODE > Core - 1 > Compiling", (pcount * 50 / pfullcount))
            Next ce1

            insertHEXCODE(-2, ";")
            insertHEXCODE(-1, ";code")
            insertStartupHEXCODE(0, xlProgramSheet.Range("F3").Text, thisWorkBook, thisRCU)
            lineNum = 1 : pcount = 0
            X0.ConsoleProgress(RCUname & "XCODE > Core -1 > Compiling", 50)
            pfullcount = xlProgramSheet.Range(xlProgramSheet.Range("F4"), xlProgramSheet.Range("F4").End(Excel.XlDirection.xlDown)).Count


            For Each ce2 In xlProgramSheet.Range(xlProgramSheet.Range("F4"), xlProgramSheet.Range("F4").End(Excel.XlDirection.xlDown))

                ThisTempLine = ce2.Text 'already in lcase
                ThisHexCode = "01 01 00 00 00 00 00 00;"


                If Strings.Left(ThisTempLine, 2) = "//" Then
                    ThisHexCode = ce2.Offset(0, -1).Text
                    insertXCODE(lineNum, Strings.Trim(Strings.Replace(ThisTempLine, "//", "")))
                ElseIf X0.hasAFunction(ThisTempLine) Then
                    ThisHexCode = FunctionConvert(ThisTempLine, ce2.Row, thisWorkBook, thisRCU, ErrorWarnLog) 'ce2 for nearest exit
                ElseIf X0.isARCU(ThisTempLine) Then
                    ThisHexCode = RCUConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.isADMX(ThisTempLine) Then
                    ThisHexCode = DMXConvert(ThisTempLine)
                ElseIf X0.isWaitABS(ThisTempLine) Then
                    ThisHexCode = WaitABSConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.hasTime(ThisTempLine) Then
                    ThisHexCode = X0.WaitConvert(ThisTempLine)
                ElseIf X0.isAnExit(ThisTempLine) Then
                    ThisHexCode = X0.ExitConvert
                ElseIf X0.isDebug(ThisTempLine) Then
                    ThisHexCode = X0.DebugConvert
                ElseIf X0.isVersion(ThisTempLine) Then
                    ThisHexCode = X0.VersionConvert(ThisTempLine)
                ElseIf X0.isALoad(ThisTempLine) Then
                    ThisHexCode = LoadConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.isASend(ThisTempLine) Then
                    ThisHexCode = SendConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.isARead(ThisTempLine) Then
                    ThisHexCode = ReadConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.isASave(ThisTempLine) Then
                    ThisHexCode = SaveConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.isAMath(ThisTempLine) Then
                    ThisHexCode = MathConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                ElseIf X0.isADimFlow(ThisTempLine) Then
                    ThisHexCode = DimFlowConvert(ThisTempLine, ce2.Offset(0, -2).Text, thisRCU, ErrorWarnLog)
                Else
                    'else part goes here, catching poorly written instructions
                    If Strings.InStr(ThisTempLine, ".set") > 1 Or Strings.InStr(ThisTempLine, ".unset") > 1 Or Strings.InStr(ThisTempLine, ".or") > 1 Or Strings.InStr(ThisTempLine, ".and") > 1 Or Strings.InStr(ThisTempLine, ".xor") > 1 Then  '[dropped from 5.2] Or Strings.instr(ThisTempLine, "$") > 2 Then
                        ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & ce2.Offset(0, -2).Text & " | " & ThisTempLine & "][Undefined variable | " & Strings.Left(ThisTempLine, Strings.InStr(ThisTempLine, ".") - 1) & "]")
                    ElseIf (Strings.InStr(ThisTempLine, "wait ") > 0) And (Not (Strings.InStr(ThisTempLine, "$hour") > 5 Or Strings.InStr(ThisTempLine, "$min") > 5 Or Strings.InStr(ThisTempLine, "$sec") > 5)) Then 'added from 5.2
                        ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & ce2.Offset(0, -2).Text & " | " & ThisTempLine & "][Invalid parameter | " & Strings.Replace(ThisTempLine, "wait ", "") & "]")
                    End If
                End If

                insertHEXCODE(lineNum, ThisHexCode)
                lineNum = lineNum + 1

                pcount = pcount + 1
                X0.ConsoleProgress(RCUname & "XCODE > Core - 2 > Compiling", 50 + (pcount * 50 / pfullcount))
            Next ce2

            'X0.setFunctionSeparaters(xlProgramSheet) 'optional
            FunctionCalibrate(thisWorkBook, thisRCU)

            'devlist size fix
            Dim devListSize As Integer = getDevLimit() - 3 'remove offset
            If xlSettingsSheet.Range("B29").Text <> (CStr(devListSize) & ";") Then
                xlSettingsSheet.Range("B29").FormulaR1C1 = xlWorkFunc.Dec2Hex(devListSize) & ";"
            End If


        Next


        releaseObject(xlDeviceSheet)
        releaseObject(xlXCODESheet)
        releaseObject(xlProgramSheet)
        releaseObject(xlSettingsSheet)
        releaseObject(xlCustomVarSheet)
        releaseObject(xlModbusVarSheet)
        releaseObject(xlFUNCTIONSheet)
        releaseObject(X0)
    End Sub

    'inserters xcode
    Private Sub insertXCODE(ByVal cLineNum As Integer, ByVal cXCODE As String)
        xlProgramSheet.Range("D" & (cLineNum + 3)).FormulaR1C1 = xlWorkFunc.Dec2Hex(cLineNum, 4)
        xlProgramSheet.Range("F" & (cLineNum + 3)).FormulaR1C1 = cXCODE
    End Sub

    Private Sub insertHEXCODE(ByVal cLineNum As Integer, ByVal cXCODE As String)
        xlProgramSheet.Range("E" & (cLineNum + 3)).FormulaR1C1 = cXCODE
    End Sub

    Private Sub insertFunction(ByVal cFcount As Integer, ByVal cLineNum As Integer, ByVal Fname As String)
        xlFUNCTIONSheet.Range("D" & (cFcount + 3)).FormulaR1C1 = cFcount
        xlFUNCTIONSheet.Range("E" & (cFcount + 3)).FormulaR1C1 = Fname
        xlFUNCTIONSheet.Range("F" & (cFcount + 3)).FormulaR1C1 = xlWorkFunc.Dec2Hex(cLineNum, 4)
    End Sub

    Private Sub insertStartupHEXCODE(ByVal cLineNum As Integer, ByVal cXCODE As String, ByRef thisWorkbook As Excel.Workbook, ByVal thisRCU As String)
        Dim LArray1() As String, thisFunc() As String, LineJ() As String
        Dim i As Integer, j As Integer, cFC As String

        thisFunc = Strings.Split("NULL,NULL,NULL,NULL", ",")
        LineJ = Strings.Split("0000,0000,0000,0000", ",")
        LArray1 = Strings.Split(cXCODE, " ")

        j = LBound(LineJ)
        For i = LBound(LArray1) To UBound(LArray1)
            If Strings.InStr(LArray1(i), "#") And j < 4 Then
                thisFunc(j) = Strings.Replace(LArray1(i), "#", "")
                LineJ(j) = X0.getLineFromFunctionName(thisFunc(j), 3, thisWorkbook, thisRCU)
                j = j + 1
            End If
        Next i

        cFC = Strings.Right(LineJ(0), 2) & " " & Strings.Left(LineJ(0), 2) & " " & Strings.Right(LineJ(1), 2) & " " & Strings.Left(LineJ(1), 2) & " " & Strings.Right(LineJ(2), 2) & " " & Strings.Left(LineJ(2), 2) & " " & Strings.Right(LineJ(3), 2) & " " & Strings.Left(LineJ(3), 2) & ";"
        xlProgramSheet.Range("E" & (cLineNum + 3)).FormulaR1C1 = cFC
    End Sub

    ' devices function calibration
    Private Sub FunctionCalibrate(ByRef thisWorkbook As Excel.Workbook, ByVal thisRCU As String)
        Dim ce3 As Excel.Range, pcount As Double, pfullcount As Double
        pcount = 0 : pfullcount = 0
        pfullcount = xlDeviceSheet.Range(xlDeviceSheet.Range("J4"), xlDeviceSheet.Range("J4").End(Excel.XlDirection.xlDown)).Count

        For Each ce3 In xlDeviceSheet.Range(xlDeviceSheet.Range("J4"), xlDeviceSheet.Range("J4").End(Excel.XlDirection.xlDown))
            If ce3.Offset(0, 4).Text <> "" Then
                Dim thisFunc As String, tempJumpOld As String, tempJump As String

                thisFunc = Strings.LCase(Strings.Trim(ce3.Offset(0, 4).Text))
                If Not isNotDuplicateFunction(thisFunc) Then
                    tempJumpOld = Strings.Mid(Strings.Trim(Strings.Replace(ce3.Text, ";", "")), 4, 1)
                    tempJump = tempJumpOld & Strings.Right(X0.getLineFromFunctionName(thisFunc, 3, thisWorkbook, thisRCU), 3)
                    ce3.FormulaR1C1 = Strings.Right(tempJump, 2) & " " & Strings.Left(tempJump, 2) & ";"
                    setFunctionNode(ce3.Row)
                Else
                    If ce3.Offset(0, 4).MergeCells Then
                        ce3.Offset(0, 4).MergeArea.Interior.ColorIndex = 0
                        ce3.Offset(0, 4).MergeArea.ClearContents()
                        ce3.Offset(0, 4).MergeArea.UnMerge()
                    End If

                    ce3.Offset(0, 4).Interior.ColorIndex = 0
                    ce3.Offset(0, 4).ClearContents()

                    unsetFunctionNode(ce3.Row)
                End If
            Else
                unsetFunctionNode(ce3.Row)
            End If

            pcount = pcount + 1

            X0.ConsoleProgress("Formatter > Device" & thisRCU & " inputs Calibration", (pcount * 100 / pfullcount))

        Next ce3

    End Sub

    Private Sub setFunctionNode(inputRange As String)
        xlDeviceSheet.Range("J" & inputRange).Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
        xlDeviceSheet.Range("J" & inputRange).Interior.Color = 65535
        xlDeviceSheet.Range("J" & inputRange).Interior.TintAndShade = 0
    End Sub

    Private Sub unsetFunctionNode(inputRange As String)
        xlDeviceSheet.Range("J" & inputRange).Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
        xlDeviceSheet.Range("J" & inputRange).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4
        xlDeviceSheet.Range("J" & inputRange).Interior.TintAndShade = 0.799981688894314
        xlDeviceSheet.Range("J" & inputRange).FormulaR1C1 = "00 00;"
    End Sub

    'Functions
    Private Function isNotDuplicateFunction(ByVal inputString As String) As Boolean
        If Strings.Len(xlFUNCTIONSheet.Range("E3").Text) = 0 Then
            isNotDuplicateFunction = True
        ElseIf Strings.Len(xlFUNCTIONSheet.Range("E4").Text) = 0 Then
            If xlFUNCTIONSheet.Range("E3").Text = inputString Then
                isNotDuplicateFunction = False
            Else
                isNotDuplicateFunction = True
            End If
        Else
            For Each ce As Excel.Range In xlFUNCTIONSheet.Range(xlFUNCTIONSheet.Range("E3"), xlFUNCTIONSheet.Range("E3").End(Excel.XlDirection.xlDown))
                If ce.Text = inputString Then
                    isNotDuplicateFunction = False
                    Exit Function
                End If
            Next ce
            isNotDuplicateFunction = True
        End If
    End Function

    
    'convertors

    'goto, gotowait, 03 01, 03 02, 03 03, if conditions
    Private Function FunctionConvert(ByVal inputString As String, inputLine As Integer, ByRef thisWorkbook As Excel.Workbook, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim LArray() As String, ConditionF As String, i As Integer
        Dim thisFunction As String, cFunctionConvert As String, LineJump As String, damper As String
        Dim withVar As Boolean

        cFunctionConvert = "00 00" : withVar = False : damper = "00"
        ConditionF = ConditionSet(inputString, thisRCU, ErrorWarnLog)
        withVar = X0.hasValidVar(inputString) Or hasValidC_Var(inputString)
        LArray = Strings.Split(inputString, " ")

        For i = LBound(LArray) To UBound(LArray)
            If Strings.InStr(LArray(i), "#") = 1 Then
                thisFunction = Strings.Replace(LArray(i), "#", "")
                LineJump = X0.getLineFromFunctionName(thisFunction, inputLine, thisWorkbook, thisRCU) 'input line for nearest exit
                cFunctionConvert = Strings.Right(LineJump, 2) & " " & Strings.Left(LineJump, 2)
            End If
        Next i

        If cFunctionConvert = "00 00" Then
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & inputLine & " | " & inputString & "][Invalid function call]")
            FunctionConvert = "01 01 00 00 00 00 00 00;"
        Else
            If ConditionF = "01 01" Or ConditionF = "03 01" Or ConditionF = "03 02" Then
                FunctionConvert = ConditionF & " " & cFunctionConvert & " " & X0.Grab_Time_1(inputString) & ";"
            ElseIf ConditionF = "03 03" Then
                FunctionConvert = ConditionF & " " & cFunctionConvert & " " & Strings.Left(X0.Grab_Time_1(inputString), 8) & " " & Grab_Dev_1(inputString, thisRCU, ErrorWarnLog) & ";"
            Else
                If withVar Then
                    ConditionF = Strings.Left(ConditionF, 2) & " 0A"

                    If ConditionF = "5c 0A" Then  '5c damper "02" in the end
                        damper = X0.Grab_variant(inputString, "milisec")

                        If damper <> "NULL" And IsNumeric(damper) Then
                            If CInt(damper) > 0 And CInt(damper) < 256 Then
                                damper = xlWorkFunc.Dec2Hex(damper, 2)
                            Else
                                damper = "02"
                            End If
                        Else
                            damper = "02"
                        End If
                    End If

                    If X0.hasValidVar(inputString) And (Not hasValidC_Var(inputString)) Then
                        FunctionConvert = ConditionF & " " & cFunctionConvert & " " & X0.Grab_Var_1(inputString) & " 00 00 " & damper & ";"
                    ElseIf (Not X0.hasValidVar(inputString)) And hasValidC_Var(inputString) Then
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

    Private Function ConditionSet(ByVal inputString As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String

        If Strings.InStr(inputString, "if ") = 1 Then
            If (Strings.InStr(inputString, "bit ") > 0 Or Strings.InStr(inputString, "bitwise ") > 0) And (Strings.InStr(inputString, "true ") > 0 Or Strings.InStr(inputString, "check ") > 0) Then
                ConditionSet = "09 01"
            ElseIf (Strings.InStr(inputString, "bit ") > 0 Or Strings.InStr(inputString, "bitwise ") > 0) And Strings.InStr(inputString, "run ") > 0 Then
                ConditionSet = "60 01"
            ElseIf Strings.InStr(inputString, "not equal ") > 0 Or Strings.InStr(inputString, "not equals to ") > 0 Or Strings.InStr(Strings.Replace(inputString, " ", ""), "if<>") = 1 Or Strings.InStr(Strings.Replace(inputString, " ", ""), "if!=") = 1 Then
                ConditionSet = "05 01"
            ElseIf Strings.InStr(inputString, "equal ") > 0 Or Strings.InStr(inputString, "equals to ") > 0 Or Strings.InStr(Strings.Replace(inputString, " ", ""), "if=") = 1 Then
                ConditionSet = "08 01"
            ElseIf Strings.InStr(inputString, "greater ") > 0 Or Strings.InStr(inputString, "greater than ") > 0 Or Strings.InStr(Strings.Replace(inputString, " ", ""), "if>") = 1 Then
                ConditionSet = "06 01"
            ElseIf Strings.InStr(inputString, "lesser ") > 0 Or Strings.InStr(inputString, "less than ") > 0 Or Strings.InStr(Strings.Replace(inputString, " ", ""), "if<") = 1 Then
                ConditionSet = "07 01"
            Else
                ConditionSet = "01 01"
            End If
        ElseIf Strings.InStr(inputString, "dimflow.") = 1 Then
            If Strings.InStr(inputString, ".oncheck ") > 0 Or Strings.InStr(inputString, ".checkon ") > 0 Then
                ConditionSet = "5b 0A"
            ElseIf Strings.InStr(inputString, ".offcheck ") > 0 Or Strings.InStr(inputString, ".checkoff ") > 0 Then
                ConditionSet = "5c 0A"
            Else
                ConditionSet = "01 01"
            End If
        ElseIf Strings.InStr(inputString, "run ") = 1 Or Strings.InStr(inputString, "reuse ") = 1 Or Strings.InStr(inputString, "re-run ") = 1 Or Strings.InStr(inputString, "rerun ") = 1 Then
            If Strings.InStr(inputString, "reuse ") = 1 Or Strings.InStr(inputString, "re-run ") = 1 Or Strings.InStr(inputString, "rerun ") = 1 Then
                If xlWorkFunc.Hex2Dec(Grab_Dev_1(inputString, thisRCU, ErrorWarnLog)) > 0 Then
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

    ' rcu convert
    Private Function RCUConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
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
        ElseIf Strings.InStr(inputString, "rcu.read aout2") = 1 Or Strings.InStr(inputString, "rcu.read analog2") = 1 Then
            Opt = "H3"
        ElseIf Strings.InStr(inputString, "rcu.read aout") = 1 Or Strings.InStr(inputString, "rcu.read analog") = 1 Then
            Opt = "H2"
        ElseIf Strings.InStr(inputString, "rcu.read") = 1 Then
            Opt = "03"
        ElseIf Strings.InStr(inputString, "rcu.xor ") = 1 Then
            Opt = "0B"
        ElseIf Strings.InStr(inputString, "rcu.set batt") = 1 Then
            Opt = "0D"
        ElseIf Strings.InStr(inputString, "rcu.unset batt") = 1 Then
            Opt = "0D"
        ElseIf Strings.InStr(inputString, "rcu.set ") = 1 And Strings.InStr(inputString, " all") > 0 Then
            Opt = "04"
        ElseIf (Strings.InStr(inputString, "rcu.set ") = 1 Or Strings.InStr(inputString, "rcu.unset ") = 1) And Strings.InStr(inputString, " es") > 0 Then
            Opt = "05"
        ElseIf (Strings.InStr(inputString, "rcu.set ") = 1 Or Strings.InStr(inputString, "rcu.unset ") = 1) And (Strings.InStr(inputString, " analog2") > 0 Or Strings.InStr(inputString, " aout2") > 0) Then
            Opt = "07"
        ElseIf (Strings.InStr(inputString, "rcu.set ") = 1 Or Strings.InStr(inputString, "rcu.unset ") = 1) And (Strings.InStr(inputString, " analog") > 0 Or Strings.InStr(inputString, " aout") > 0) Then
            Opt = "06"
        ElseIf Strings.InStr(inputString, "rcu.set ") = 1 Or Strings.InStr(inputString, "rcu.or ") = 1 Then
            Opt = "09"
        ElseIf Strings.InStr(inputString, "rcu.unset ") = 1 Or Strings.InStr(inputString, "rcu.and ") = 1 Then
            Opt = "0A"
        ElseIf Strings.InStr(inputString, "rcu.pwm ") = 1 Then
            Opt = "00"
        ElseIf Strings.InStr(inputString, "rcu.dim es") = 1 Then
            Opt = "G0"
        ElseIf Strings.InStr(inputString, "rcu.dim aout2") = 1 Or Strings.InStr(inputString, "rcu.dim analog2") = 1 Then
            Opt = "G2"
        ElseIf Strings.InStr(inputString, "rcu.dim aout") = 1 Or Strings.InStr(inputString, "rcu.dim analog") = 1 Then
            Opt = "G1"
        ElseIf Strings.InStr(inputString, "rcu.stop es") = 1 Then
            Opt = "I0"
        ElseIf Strings.InStr(inputString, "rcu.stop aout2") = 1 Or Strings.InStr(inputString, "rcu.stop analog2") = 1 Then
            Opt = "I2"
        ElseIf Strings.InStr(inputString, "rcu.stop aout") = 1 Or Strings.InStr(inputString, "rcu.stop analog") = 1 Then
            Opt = "I1"
        End If

        If Opt = "03" Or Opt = "08" Or Opt = "0C" Or Opt = "0E" Then
            RCUConvert = "02 " & Opt & " 00 00 00 00 00 00;"
        ElseIf Opt = "H1" Or Opt = "H2" Or Opt = "H3" Then
            RCUConvert = "53 0" & Strings.Right(Opt, 1) & " 00 00 00 00 00 00;"
        ElseIf X0.hasValidVar(inputString) Then
            If Strings.Left(Opt, 1) = "G" Or Strings.Left(Opt, 1) = "I" Then
                RCUConvert = "5" & Strings.Right(Opt, 1) & " 0A 00 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            Else
                RCUConvert = "02 " & Opt & " 0A 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            End If
        ElseIf hasValidC_Var(inputString) Then
            If Strings.Left(Opt, 1) = "G" Or Strings.Left(Opt, 1) = "I" Then
                RCUConvert = "5" & Strings.Right(Opt, 1) & " 0A 00 00 " & Grab_Var_2(inputString) & " 00 00 00;"
            Else
                RCUConvert = "02 " & Opt & " 0A 00 " & Grab_Var_2(inputString) & " 00 00 00;"
            End If
        Else
            If Smode = "00" Then
                If Strings.Left(Opt, 1) = "G" Or Strings.Left(Opt, 1) = "I" Then
                    RCUConvert = "5" & Strings.Right(Opt, 1) & " 00 00 00 00 00 00 00;"
                Else
                    RCUConvert = "02 " & Opt & " 00 00 00 00 00 00;"
                End If
            Else
                If Strings.InStr(inputString, "$var=") > 4 Then
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                    RCUConvert = "01 01 00 00 00 00 00 00;"
                Else
                    If Opt = "05" Or Opt = "06" Or Opt = "07" Then
                        If Strings.InStr(inputString, "rcu.unset ") = 1 Then
                            RCUConvert = "02 " & Opt & " 01 00 00 00 00 00;"
                        Else
                            RCUConvert = "02 " & Opt & " 01 00 " & X0.Grab_value_4(inputString) & " 00 00 00;"
                        End If
                    ElseIf Opt = "0D" Then
                        If Strings.InStr(inputString, "rcu.unset") = 1 Then
                            RCUConvert = "02 " & Opt & " 01 00 00 00 00 00;"
                        ElseIf Strings.InStr(inputString, "rcu.set") = 1 And Strings.InStr(inputString, " $") < 1 Then
                            RCUConvert = "02 " & Opt & " 01 00 01 00 00 00;"
                        Else
                            RCUConvert = "02 " & Opt & " 01 00 " & X0.Grab_value_4(inputString) & " 00 00 00;"
                        End If
                    ElseIf Strings.Left(Opt, 1) = "G" Or Strings.Left(Opt, 1) = "I" Then
                        RCUConvert = "5" & Strings.Right(Opt, 1) & " 01 00 00 " & X0.Grab_Time_3(inputString) & " " & X0.Grab_value_4(inputString) & ";"
                    Else
                        RCUConvert = "02 " & Opt & " 01 00 " & X0.Grab_value_1(inputString, "OR") & ";"
                    End If
                End If
            End If
        End If
    End Function

    ' dmx convert
    Private Function DMXConvert(ByVal inputString As String) As String
        Dim Opt As String, Smode As String, ioCh As String
        Opt = "01" : Smode = "01"

        If Strings.InStr(inputString, " acc") > 1 Then
            Smode = "00"
        End If

        ioCh = X0.Grab_variant(inputString, "ch")

        If Strings.InStr(inputString, "dmx.send ") = 1 Or Strings.InStr(inputString, "dmx.ramp ") = 1 Or Strings.InStr(inputString, "dmx.set ") = 1 Then
            Opt = "01"
        ElseIf Strings.InStr(inputString, "dmx.stop ") = 1 Or Strings.InStr(inputString, "dmx.hold ") = 1 Then
            Opt = "02"
        ElseIf Strings.InStr(inputString, "dmx.read ") = 1 Then
            Opt = "03"
        End If

        If ioCh <> "NULL" Then
            ioCh = xlWorkFunc.Dec2Hex(CInt(ioCh), 2)

            If Opt = "03" Then
                DMXConvert = "61 02 00 " & ioCh & " 00 00 00 00;"
            ElseIf Opt = "02" Then
                DMXConvert = "61 01 01 " & ioCh & " 00 00 00 00;"
            Else
                If Smode = "00" Then
                    DMXConvert = "61 01 00 " & ioCh & " 00 00 00 00;"
                ElseIf X0.hasValidVar(inputString) Then
                    DMXConvert = "61 01 0A " & ioCh & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
                ElseIf hasValidC_Var(inputString) Then
                    DMXConvert = "61 01 0A " & ioCh & " " & Grab_Var_2(inputString) & " 00 00 00;"
                Else
                    DMXConvert = "61 01 01 " & ioCh & " " & X0.Grab_Time_3(inputString) & " " & X0.Grab_value_6(inputString) & ";"
                End If
            End If
        Else
            DMXConvert = "01 01 00 00 00 00 00 00;"
        End If
    End Function

    'abs convert
    Private Function WaitABSConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        If X0.hasValidVar(inputString) Then
            WaitABSConvert = "0A 0A 00 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf hasValidC_Var(inputString) Then
            WaitABSConvert = "0A 0A 00 00 " & Grab_Var_2(inputString) & " 00 00 00;"
        ElseIf X0.hasTime(inputString) Then
            WaitABSConvert = "0A 01 00 00 " & X0.Grab_Time_1(inputString) & ";"
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid Wait ABS instruction]")
            WaitABSConvert = "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' load convert
    Private Function LoadConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
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
                    thisLen = Strings.Right("00" & xlWorkFunc.Hex2Dec(thisLen), 2)
                Else
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Incompatible length | " & X0.Grab_variant(inputString, "len") & "]")
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
        ElseIf X0.hasValidVar(inputString) Then
            LoadConvert = "10 80 0A " & Smode & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf hasValidC_Var(inputString) Then
            LoadConvert = "10 80 0A " & Smode & " " & Grab_Var_2(inputString) & " 00 00 00;"
        ElseIf hasValidM_Var(inputString) And Strings.InStr(inputString, " mod") > 1 And Strings.InStr(inputString, " flo") > 1 Then 'added in ver 5.3
            LoadConvert = "10 80 15 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf hasValidM_Var(inputString) And Strings.InStr(inputString, " mod") > 1 Then
            LoadConvert = "10 80 12 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf Strings.InStr(inputString, " cbs") > 1 Then
            LoadConvert = "10 85 01 " & Smode & " " & X0.Grab_element(inputString) & ";"
        ElseIf Strings.InStr(inputString, " enum") > 1 Then
            LoadConvert = "10 80 10 " & Smode & " 00 00 00 00;"
        ElseIf Strings.InStr(inputString, " time") > 1 Then
            If Strings.InStr(inputString, " unix") > 1 Or Strings.InStr(inputString, " utc") > 1 Then
                LoadConvert = "10 84 00 " & Smode & " 00 00 00 00;"
            ElseIf Strings.InStr(inputString, " now") > 1 Then
                LoadConvert = "10 80 01 " & Smode & " " & X0.Give_Time_1() & ";"
            Else
                LoadConvert = "10 82 00 " & Smode & " 00 00 00 00;"
            End If
        Else
            If Strings.InStr(inputString, "$var=") > 4 Then
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                LoadConvert = "01 01 00 00 00 00 00 00;"
            Else
                LoadConvert = "10 80 01 " & Smode & " " & X0.Grab_value_1(inputString, "OR") & ";"
            End If
        End If
    End Function

    ' send convert
    Private Function SendConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        If (Strings.InStr(inputString, "send.mod ") = 1 Or Strings.InStr(inputString, "send.modbus ") = 1) And (Strings.InStr(inputString, " $dev=") > 5 Or Strings.InStr(inputString, " $type=") > 5) And Strings.InStr(inputString, " $reg=") > 5 Then
            Dim thisDev As String = Grab_Dev_1(inputString, thisRCU, ErrorWarnLog)
            If xlWorkFunc.Hex2Dec(thisDev) > 0 Then
                '20 & 22
                If Strings.InStr(inputString, "$len=") Then
                    SendConvert = "22 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_1(inputString, "OR") & ";"
                Else
                    SendConvert = "20 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_3(inputString, "OR") & " 00 00;"
                End If
            ElseIf xlWorkFunc.Hex2Dec(X0.Grab_Dev_Type_1(inputString)) > 0 Then
                '21 & 23
                If Strings.InStr(inputString, "$len=") Then
                    SendConvert = "23 " & X0.Grab_Dev_Type_1(inputString) & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_1(inputString, "OR") & ";"
                Else
                    SendConvert = "21 " & X0.Grab_Dev_Type_1(inputString) & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_3(inputString, "OR") & " 00 00;"
                End If
            Else
                SendConvert = "01 01 00 00 00 00 00 00;"
            End If
        Else

            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid modbus instruction]")
            SendConvert = "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' read convert
    Private Function ReadConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        If Strings.InStr(inputString, "read.mod ") = 1 Or Strings.InStr(inputString, "read.modbus ") = 1 Then
            Dim thisDev As String = Grab_Dev_1(inputString, thisRCU, ErrorWarnLog)
            If xlWorkFunc.Hex2Dec(thisDev) > 0 Then
                '20 & 22
                If Strings.InStr(inputString, "$len=") Then
                    ReadConvert = "26 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " " & X0.Grab_value_1(inputString, "OR") & ";"
                Else
                    ReadConvert = "26 " & thisDev & " " & X0.Grab_Reg_1(inputString) & " 02 00 00 00;"
                End If
            Else
                If xlWorkFunc.Hex2Dec(thisDev) = 0 Then
                    ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid device]")
                End If
                ReadConvert = "01 01 00 00 00 00 00 00;"
            End If
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid modbus instruction]")
            ReadConvert = "01 01 00 00 00 00 00 00;"
        End If
    End Function

    ' save convert
    Private Function SaveConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
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
                    thisLen = Strings.Right("00" & xlWorkFunc.Hex2Dec(thisLen), 2)
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
        ElseIf X0.hasValidVar(inputString) Then
            SaveConvert = "10 81 0A " & Smode & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
        ElseIf hasValidC_Var(inputString) Then
            SaveConvert = "10 81 0A " & Smode & " " & Grab_Var_2(inputString) & " 00 00 00;"
        ElseIf hasValidM_Var(inputString) And Strings.InStr(inputString, " mod") > 1 And Strings.InStr(inputString, " flo") > 1 Then 'added in ver 5.3
            SaveConvert = "10 81 15 " & Smode & " " & Grab_Var_3(inputString) & " 00 00 00;"
        ElseIf hasValidM_Var(inputString) And Strings.InStr(inputString, " mod") > 1 Then
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
    Private Function MathConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim Opt As String, Smode As String
        Opt = "80" : Smode = "00"

        If Strings.InStr(inputString, "shift ") > 1 Then
            Smode = "01"
        End If

        If Strings.InStr(inputString, "math.add ") = 1 Or Strings.InStr(inputString, "math.+ ") = 1 Then
            Opt = "01"
        ElseIf Strings.InStr(inputString, "math.sub ") = 1 Or Strings.InStr(inputString, "math.- ") = 1 Then
            Opt = "02"
        ElseIf Strings.InStr(inputString, "math.mul ") = 1 Or Strings.InStr(inputString, "math.* ") = 1 Or Strings.InStr(inputString, "math.x ") = 1 Then
            Opt = "03"
        ElseIf Strings.InStr(inputString, "math.or ") = 1 Or Strings.InStr(inputString, "math.|| ") = 1 Then
            Opt = "04"
        ElseIf Strings.InStr(inputString, "math.xor ") = 1 Or Strings.InStr(inputString, "math.x| ") = 1 Then
            Opt = "05"
        ElseIf Strings.InStr(inputString, "math.and ") = 1 Or Strings.InStr(inputString, "math.&& ") = 1 Then
            Opt = "06"
        ElseIf Strings.InStr(inputString, "math.shiftl ") = 1 Or Strings.InStr(inputString, "math.<< ") = 1 Then
            Opt = "07"
        ElseIf Strings.InStr(inputString, "math.shiftr ") = 1 Or Strings.InStr(inputString, "math.>> ") = 1 Then
            Opt = "08"
        ElseIf Strings.InStr(inputString, "math.div ") = 1 Or Strings.InStr(inputString, "math./ ") = 1 Then
            Opt = "09"
        ElseIf Strings.InStr(inputString, "math.mod ") = 1 Or Strings.InStr(inputString, "math.% ") = 1 Then
            Opt = "0A"
        End If

        If Opt = "80" Then
            'not math function
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid math instruction]")
            MathConvert = "01 01 00 00 00 00 00 00;"
        Else
            If X0.hasValidVar(inputString) Then
                MathConvert = "10 " & Opt & " 0A " & Smode & " " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            ElseIf hasValidC_Var(inputString) Then
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
    Private Function DimFlowConvert(ByVal inputString As String, ByVal thisLine As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim Dmode As String
        Dmode = "01 01"

        If Strings.InStr(inputString, ".start ") > 0 Or Strings.InStr(inputString, ".onstart ") > 0 Or Strings.InStr(inputString, ".starton ") > 0 Then
            Dmode = "5a 0A"
        ElseIf Strings.InStr(inputString, ".getdir") > 0 Or Strings.InStr(inputString, ".dirget ") > 0 Then
            Dmode = "5e 0A"
        ElseIf Strings.InStr(inputString, ".setdir") > 0 Or Strings.InStr(inputString, ".dirset ") > 0 Then
            Dmode = "5d 0A"
        End If

        If Dmode <> "01 01" Then
            If X0.hasValidVar(inputString) Then
                DimFlowConvert = Dmode & " 00 00 " & X0.Grab_Var_1(inputString) & " 00 00 00;"
            ElseIf hasValidC_Var(inputString) Then
                DimFlowConvert = Dmode & " 00 00 " & Grab_Var_2(inputString) & " 00 00 00;"
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Undefined variable | " & X0.Grab_variant(inputString, "var") & "]")
                DimFlowConvert = "01 01 00 00 00 00 00 00;"
            End If
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Program" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid instruction]")
            DimFlowConvert = "01 01 00 00 00 00 00 00;"
        End If
    End Function

    '------------------------------------------------------





    ' is other var direct
    Private Function isOtherVarDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, ".") > 0 Then
            Dim VarArray() As String, tempS As String
            VarArray = Strings.Split(inputString, ".")
            tempS = "load $var=" & VarArray(0) 'create a simple load instruction with var
            If X0.hasValidVar(tempS) Or hasValidC_Var(tempS) Or hasValidM_Var(tempS) Then
                isOtherVarDirect = True
            Else
                isOtherVarDirect = False
            End If
        Else
            isOtherVarDirect = False
        End If
    End Function

    'custom var
    Private Function hasValidC_Var(ByVal inputString As String) As Boolean 'Custom var, moified in ver 5.0
        hasValidC_Var = False
        If C_VarListNotEmpty() Then
            Dim LArray_T2() As String, V As String, i As Integer
            V = ""
            LArray_T2 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If Strings.InStr(Strings.Replace(LArray_T2(i), " ", ""), "var=") Then
                    V = Strings.Replace(LArray_T2(i), "var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
                Dim rngVar As Excel.Range
                For Each rngVar In xlCustomVarSheet.Range(xlCustomVarSheet.Range("B3").Offset(1, 0), xlCustomVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
                    If V = Strings.Trim(Strings.LCase(rngVar.Text)) Or V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
                        hasValidC_Var = True
                        Exit Function
                    End If
                Next rngVar
            End If
        End If
    End Function

    Private Function C_VarListNotEmpty() As Boolean
        C_VarListNotEmpty = False
        If Not IsNothing(xlCustomVarSheet) Then
            If Strings.Len(xlCustomVarSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
                C_VarListNotEmpty = True
            End If
        End If
    End Function

    Private Function Grab_Var_2(ByVal inputString As String) As String 'Custom var
        Dim V As String, LArray_T2() As String, P As String, i As Integer
        Dim rngVar As Excel.Range
        V = "" : P = ""
        LArray_T2 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_T2) To UBound(LArray_T2)
            If Strings.InStr(Strings.Replace(LArray_T2(i), " ", ""), "var=") = 1 Then
                V = Strings.Replace(LArray_T2(i), "var=", "")
                Exit For
            End If
        Next i
        For Each rngVar In xlCustomVarSheet.Range(xlCustomVarSheet.Range("B3").Offset(1, 0), xlCustomVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
            If V = Strings.Trim(Strings.LCase(rngVar.Text)) Or V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
                P = rngVar.Text
                Exit For
            End If
        Next rngVar
        Grab_Var_2 = Strings.Right("00" & P, 2)
    End Function

    ' modbus var
    Private Function hasValidM_Var(ByVal inputString As String) As Boolean 'Modbus var, moified in ver 5.0
        hasValidM_Var = False
        If M_VarListNotEmpty() Then
            Dim LArray_T2() As String, V As String, i As Integer
            V = ""
            LArray_T2 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If Strings.InStr(Strings.Replace(LArray_T2(i), " ", ""), "var=") Then
                    V = Strings.Replace(LArray_T2(i), "var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
                Dim rngVar As Excel.Range
                For Each rngVar In xlModbusVarSheet.Range(xlModbusVarSheet.Range("B3").Offset(1, 0), xlModbusVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
                    If V = Strings.Trim(Strings.LCase(rngVar.Text)) Or V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
                        hasValidM_Var = True
                        Exit Function
                    End If
                Next rngVar
            End If
        End If
    End Function

    Private Function M_VarListNotEmpty() As Boolean
        M_VarListNotEmpty = False
        If Not IsNothing(xlModbusVarSheet) Then
            If Strings.Len(xlModbusVarSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
                M_VarListNotEmpty = True
            End If
        End If
    End Function

    Private Function Grab_Var_3(inputString As String) As String 'Modbus var
        Dim V As String, LArray_T2() As String, P As String
        Dim rngVar As Excel.Range

        V = "" : P = ""
        LArray_T2 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T2) To UBound(LArray_T2)
            If Strings.InStr(Strings.Replace(LArray_T2(i), " ", ""), "var=") Then
                V = Strings.Replace(LArray_T2(i), "var=", "")
                Exit For
            End If
        Next i

        For Each rngVar In xlModbusVarSheet.Range(xlModbusVarSheet.Range("B3").Offset(1, 0), xlModbusVarSheet.Range("B3").End(Excel.XlDirection.xlDown))
            If V = Strings.Trim(Strings.LCase(rngVar.Text)) Or V = Strings.Trim(Strings.LCase(rngVar.Offset(0, 1).Text)) Then
                P = rngVar.Text
                Exit For
            End If
        Next rngVar

        Grab_Var_3 = Strings.Right("00" & P, 2)
    End Function


    ' grabbers core 1
    Private Function GrabIOEXPDirect(ByVal inputString As String, ByVal thisRCU As String, ByVal thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, LArray() As String, ioName As String, L As Integer
        Dim vGrabIOEXPDirect As New List(Of String)

        'detect ioexp
        If Strings.Mid(inputString, Strings.InStr(inputString, "ioexp") + 5, 1) = "." Then
            ioName = "ioexp1"
        Else
            ioName = "ioexp" & Strings.Mid(inputString, Strings.InStr(inputString, "ioexp") + 5, 1)
        End If

        If Strings.InStr(inputString, "read") = (Strings.InStr(inputString, ".") + 1) Then
            If Strings.InStr(inputString, " relay") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=4 $len=2")
            ElseIf Strings.InStr(inputString, " aout8") > 9 Or Strings.InStr(inputString, " analog8") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=43 $len=2")
            ElseIf Strings.InStr(inputString, " aout7") > 9 Or Strings.InStr(inputString, " analog7") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=39 $len=2")
            ElseIf Strings.InStr(inputString, " aout6") > 9 Or Strings.InStr(inputString, " analog6") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=35 $len=2")
            ElseIf Strings.InStr(inputString, " aout5") > 9 Or Strings.InStr(inputString, " analog5") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=31 $len=2")
            ElseIf Strings.InStr(inputString, " aout4") > 9 Or Strings.InStr(inputString, " analog4") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=27 $len=2")
            ElseIf Strings.InStr(inputString, " aout3") > 9 Or Strings.InStr(inputString, " analog3") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=23 $len=2")
            ElseIf Strings.InStr(inputString, " aout2") > 9 Or Strings.InStr(inputString, " analog2") > 9 Then
                vGrabIOEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=19 $len=2")
            ElseIf Strings.InStr(inputString, " aout") > 9 Or Strings.InStr(inputString, " analog") > 9 Then
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

            ioString = Strings.Replace(Strings.Replace(inputString, " ", ""), "#", "$")
            ioString = Strings.Mid(ioString, 1, Strings.InStr(ioString, "$") - 1)
            LArray = Strings.Split(ioString, ".")

            If UBound(LArray) > 0 Then

                vGrabIOEXPDirect.Add("load $var=" & ioName)

                For L = LBound(LArray) + 1 To UBound(LArray)
                    If (LArray(L) = "set" Or LArray(L) = "or") And X0.Grab_value_0(inputString, L) <> "NULL" Then
                        vGrabIOEXPDirect.Add("math.or $" & X0.Grab_value_0(inputString, L))
                    ElseIf (LArray(L) = "unset" Or LArray(L) = "and") And X0.Grab_value_0(inputString, L) <> "NULL" Then
                        vGrabIOEXPDirect.Add("math.and $" & X0.Grab_value_0(inputString, L))
                    ElseIf (LArray(L) = "xor") And X0.Grab_value_0(inputString, L) <> "NULL" Then
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

        GrabIOEXPDirect = vGrabIOEXPDirect
    End Function

    Private Function GrabIODEXPDirect(ByVal inputString As String, ByVal thisRCU As String, ByVal thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioName As String
        Dim vGrabIODEXPDirect As New List(Of String)

        'detect ioexp
        If Strings.Mid(inputString, Strings.InStr(inputString, "iodexp") + 6, 1) = "." Then
            ioName = "iodexp1"
        Else
            ioName = "iodexp" & Strings.Mid(inputString, Strings.InStr(inputString, "iodexp") + 6, 1)
        End If

        If Strings.InStr(inputString, "read") = (Strings.InStr(inputString, ".") + 1) Then
            If Strings.InStr(inputString, " aout8") > 9 Or Strings.InStr(inputString, " analog8") > 9 Or Strings.InStr(inputString, " ch8") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=34 $len=2")
            ElseIf Strings.InStr(inputString, " aout7") > 9 Or Strings.InStr(inputString, " analog7") > 9 Or Strings.InStr(inputString, " ch7") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=30 $len=2")
            ElseIf Strings.InStr(inputString, " aout6") > 9 Or Strings.InStr(inputString, " analog6") > 9 Or Strings.InStr(inputString, " ch6") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=26 $len=2")
            ElseIf Strings.InStr(inputString, " aout5") > 9 Or Strings.InStr(inputString, " analog5") > 9 Or Strings.InStr(inputString, " ch5") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=22 $len=2")
            ElseIf Strings.InStr(inputString, " aout4") > 9 Or Strings.InStr(inputString, " analog4") > 9 Or Strings.InStr(inputString, " ch4") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=18 $len=2")
            ElseIf Strings.InStr(inputString, " aout3") > 9 Or Strings.InStr(inputString, " analog3") > 9 Or Strings.InStr(inputString, " ch3") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=14 $len=2")
            ElseIf Strings.InStr(inputString, " aout2") > 9 Or Strings.InStr(inputString, " analog2") > 9 Or Strings.InStr(inputString, " ch2") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=10 $len=2")
            ElseIf Strings.InStr(inputString, " aout") > 9 Or Strings.InStr(inputString, " analog") > 9 Or Strings.InStr(inputString, " ch") > 9 Then
                vGrabIODEXPDirect.Add("read.modbus $dev=" & ioName & " $reg=6 $len=2")
            Else
                vGrabIODEXPDirect.Add("nop")
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IODEXP instruction]")
            End If
        Else
            vGrabIODEXPDirect(1) = "nop"
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid IODEXP instruction]")
        End If

        GrabIODEXPDirect = vGrabIODEXPDirect
    End Function

    Private Function GrabRCUDirect(ByVal inputString As String) As List(Of String)
        Dim LArray() As String, LArray1() As String
        Dim L As Integer
        LArray = Strings.Split(Strings.Replace(inputString, "#", ""), " $")
        LArray1 = Strings.Split(LArray(0), ".")

        Dim vGrabRCUDirect As New List(Of String)

        For L = 0 To UBound(LArray1) - 1
            vGrabRCUDirect.Add("rcu." & LArray1(L + 1) & " $" & X0.Grab_value_0(inputString, L + 1))
        Next L

        GrabRCUDirect = vGrabRCUDirect
    End Function

    Private Function GrabMathDirect(ByVal inputString As String, ByVal thisRCU As String, ByVal thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)

        Dim LArray() As String, LArray1() As String
        Dim L As Integer
        LArray = Strings.Split(Strings.Replace(inputString, "#", ""), " $")
        LArray1 = Strings.Split(LArray(0), ".")

        Dim vGrabMathDirect As New List(Of String)

        For L = 0 To UBound(LArray1) - 1
            If Not X0.Grab_value_0(inputString, L + 1) = "NULL" Then
                vGrabMathDirect.Add("math." & LArray1(L + 1) & " $" & X0.Grab_value_0(inputString, L + 1))
            Else
                ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & " at """ & LArray1(L + 1) & """][Invalid math data type]")
                vGrabMathDirect.Add("math." & LArray1(L + 1))
            End If
        Next L

        GrabMathDirect = vGrabMathDirect
    End Function

    Private Function GrabSaveDirect(ByVal inputString As String) As List(Of String)
        Dim ioString As String, LArray() As String, LArray1() As String, ioLen As String
        Dim vGrabSaveDirect As New List(Of String)

        ioString = "" : ioLen = "2"

        If Strings.InStr(inputString, " $var=") > 3 And Strings.InStr(inputString, " $len=") > 3 Then
            'added in v5.3 specially for save settings with length
            'can inprove this entire function later
            LArray = Strings.Split(inputString, " $") 'filter with space

            For i = LBound(LArray) To UBound(LArray)
                If Strings.InStr(LArray(i), "var=") = 1 Then
                    ioString = Strings.Replace(LArray(i), "var=", "")
                End If
                If Strings.InStr(LArray(i), "len=") = 1 Then
                    ioLen = Strings.Replace(LArray(i), "len=", "")
                End If
            Next i
            LArray1 = Strings.Split(ioString, ",")

            For i = LBound(LArray1) To UBound(LArray1)
                vGrabSaveDirect.Add(LArray(0) & " $var=" & LArray1(i) & " $len=" & ioLen)
            Next i
        ElseIf Strings.InStr(inputString, " $var=") > 3 Then
            LArray = Strings.Split(inputString, "$")

            For i = LBound(LArray) To UBound(LArray)
                If Strings.InStr(LArray(i), "var=") = 1 Then
                    ioString = Strings.Replace(LArray(i), "var=", "")
                End If
            Next i
            LArray1 = Strings.Split(ioString, ",")

            For i = LBound(LArray1) To UBound(LArray1)
                vGrabSaveDirect.Add(LArray(0) & "$var=" & LArray1(i))
            Next i
        Else
            vGrabSaveDirect.Add(inputString)
        End If

        GrabSaveDirect = vGrabSaveDirect
    End Function

    Private Function GrabOtherVarDirect(ByVal inputString As String) As List(Of String)
        Dim LArray() As String, LArray1() As String
        Dim OtherVar As String, L As Integer, Smode As String
        Dim vGrabOtherVarDirect As New List(Of String)

        LArray = Strings.Split(Strings.Replace(inputString, "#", ""), " $")
        LArray1 = Strings.Split(LArray(0), ".")
        OtherVar = LArray1(0)
        Smode = ""

        If hasValidM_Var("load $var=" & OtherVar) Then
            Smode = " modbus"
        End If

        If LArray1(1) = "set" Then
            If hasValidM_Var(inputString) Then
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

        GrabOtherVarDirect = vGrabOtherVarDirect
    End Function

    Private Function GrabModbusDirect(ByVal inputString As String) As List(Of String)
        Dim len_check As String, val_check As String
        Dim vGrabModbusDirect As New List(Of String)

        len_check = X0.Grab_variant(inputString, "len")
        val_check = X0.Grab_value_0(inputString, 1)

        If len_check <> "NULL" And val_check <> "NULL" Then
            vGrabModbusDirect.Add("load $" & val_check)
            vGrabModbusDirect.Add(Strings.Replace(inputString, " $" & val_check, ""))
        ElseIf len_check = "NULL" And val_check = "NULL" Then
            vGrabModbusDirect.Add(Strings.Trim(inputString & " $len=2"))
        Else
            vGrabModbusDirect.Add(inputString)
        End If

        GrabModbusDirect = vGrabModbusDirect
    End Function

    Private Function GrabDimmerDirect(ByVal inputString As String, thisRCU As String, thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, chString As String, LArray() As String, LArray1() As String, LArray2() As String, i As Integer
        Dim thisIOdev As String, thisChReg As Integer, ioLen As Integer

        thisIOdev = "ioexp1" : ioString = "" : chString = ""
        ioLen = 6
        LArray = Strings.Split(inputString, " $")

        For i = LBound(LArray) To UBound(LArray)
            If Strings.InStr(LArray(i), "dev=") = 1 Then
                ioString = Strings.Replace(LArray(i), "dev=", "")
            ElseIf Strings.InStr(LArray(i), "ch=") = 1 Then
                chString = Strings.Replace(LArray(i), "ch=", "")
            End If
        Next i

        Dim vGrabDimmerDirect As New List(Of String)


        If Strings.Len(ioString) > 0 And Strings.Len(chString) > 0 Then

            LArray1 = Strings.Split(ioString, ",")
            LArray2 = Strings.Split(X0.CommaExpand(chString, 1, 16), ",") 'added in 5.1

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
                If Strings.Len(LArray1(i)) > 1 And IsNumeric(Strings.Right(LArray1(i), 1)) Then
                    thisIOdev = LArray1(i)
                ElseIf Strings.Len(LArray1(i)) > 1 And Not IsNumeric(Strings.Right(LArray1(i), 1)) Then
                    thisIOdev = LArray1(i) & "1"
                ElseIf Strings.Len(LArray1(i)) = 1 And IsNumeric(Strings.Right(LArray1(i), 1)) Then
                    thisIOdev = Strings.Mid(thisIOdev, 1, Strings.Len(thisIOdev) - 1) & LArray1(i)
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

                If Strings.InStr(thisIOdev, "ioe") = 1 Then
                    thisChReg = 12 + 4 * (thisChReg - 1)
                Else
                    thisChReg = 3 + 4 * (thisChReg - 1)
                End If

                If Strings.InStr(inputString, " stop") > 1 Then
                    thisChReg = thisChReg + 2
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

    Private Function GrabDimmerReadDirect(ByVal inputString As String, ByVal thisRCU As String, ByVal thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, chString As String, LArray() As String, LArray1() As String, LArray2() As String
        Dim thisIOdev As String, thisChReg As Integer, ioLen As Integer, i As Integer

        thisIOdev = "ioexp1" : ioString = "" : chString = ""
        ioLen = 2
        LArray = Strings.Split(inputString, " $")

        For i = LBound(LArray) To UBound(LArray)
            If Strings.InStr(LArray(i), "dev=") = 1 Then
                ioString = Strings.Replace(LArray(i), "dev=", "")
            ElseIf Strings.InStr(LArray(i), "ch=") = 1 Then
                chString = Strings.Replace(LArray(i), "ch=", "")
            End If
        Next i

        Dim vGrabDimmerReadDirect As New List(Of String)


        If Strings.Len(ioString) > 0 And Strings.Len(chString) > 0 Then

            LArray1 = Strings.Split(ioString, ",")
            LArray2 = Strings.Split(X0.CommaExpand(chString, 1, 16), ",")

            If Strings.Len(LArray1(0)) > 1 And IsNumeric(Strings.Right(LArray1(0), 1)) Then
                thisIOdev = LArray1(0)
            ElseIf Strings.Len(LArray1(0)) > 1 And Not IsNumeric(Strings.Right(LArray1(0), 1)) Then
                thisIOdev = LArray1(0) & "1"
            ElseIf Strings.Len(LArray1(0)) = 1 And IsNumeric(Strings.Right(LArray1(0), 1)) Then
                thisIOdev = Strings.Mid(thisIOdev, 1, Strings.Len(thisIOdev) - 1) & LArray1(0)
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

    Private Function GrabDaliDirect(ByVal inputString As String, ByVal thisRCU As String, ByVal thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioString As String, LArray1() As String, ioName As String, i As Integer
        Dim Dalibus As String, ioBlst As String, ioGp As String
        Dim idFactor As Integer, bgLum As String, bgID As String

        ioString = Strings.Replace(Strings.Replace(inputString, " ", ""), "#", "$")
        ioString = Strings.Mid(ioString, 1, Strings.InStr(ioString, "$") - 1)
        Dalibus = X0.Grab_variant(inputString, "bus")
        ioBlst = X0.Grab_variant(inputString, "ballast")
        ioGp = X0.Grab_variant(inputString, "group")
        idFactor = 0
        bgLum = X0.Grab_value_5(inputString)

        If Strings.Mid(inputString, Strings.InStr(inputString, "dali") + 4, 1) = "." Then
            ioName = "dali1"
        Else
            ioName = "dali" & Strings.Mid(inputString, Strings.InStr(inputString, "dali") + 4, 1)
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

        If Strings.InStr(ioString, ".up") > 0 Or Strings.InStr(ioString, ".on") > 0 Then
            idFactor = 1
            bgLum = "05"
        ElseIf Strings.InStr(ioString, ".down") > 0 Or Strings.InStr(ioString, ".off") > 0 Then
            idFactor = 1
            bgLum = "00"
        ElseIf Strings.InStr(ioString, ".stop") > 0 Then
            idFactor = 0
            bgLum = "FF"
        Else
            idFactor = 0
            bgLum = X0.Grab_value_5(inputString)
        End If

        Dim vGrabDaliDirect As New List(Of String)


        If ioBlst <> "NULL" Or ioGp <> "NULL" Then
            If ioBlst <> "NULL" Then
                ioBlst = X0.CommaExpand(ioBlst, 0, 63)
                LArray1 = Strings.Split(ioBlst, ",")
                For i = LBound(LArray1) To UBound(LArray1)
                    If IsNumeric(LArray1(i)) Then
                        If CInt(LArray1(i)) > -1 And CInt(LArray1(i)) < 64 Then

                            bgID = xlWorkFunc.Dec2Hex(CInt(LArray1(i)) * 2 + idFactor, 2)
                            vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & Dalibus & " $hex=" & bgID & bgLum)
                        End If
                    End If
                Next i
            End If

            If ioGp <> "NULL" Then
                ioGp = X0.CommaExpand(ioGp, 0, 15)
                LArray1 = Strings.Split(ioGp, ",")
                For i = LBound(LArray1) To UBound(LArray1)
                    If IsNumeric(LArray1(i)) Then
                        If CInt(LArray1(i)) > -1 And CInt(LArray1(i)) < 16 Then
                            bgID = xlWorkFunc.Dec2Hex(CInt(LArray1(i)) * 2 + idFactor + 128, 2)
                            vGrabDaliDirect.Add("send.modbus $dev=" & ioName & " $reg=" & Dalibus & " $hex=" & bgID & bgLum)
                        End If
                    End If
                Next i
            End If
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DALI instruction]")
            vGrabDaliDirect.Add("nop")
        End If

        GrabDaliDirect = vGrabDaliDirect
    End Function

    Private Function GrabDMXDirect(ByVal inputString As String, ByVal thisRCU As String, ByVal thisLine As String, ByRef ErrorWarnLog As String()) As List(Of String)
        Dim ioChString As String, ioChExp As String, LArray1() As String, i As Integer

        ioChString = X0.Grab_variant(inputString, "ch")

        Dim vGrabDMXDirect As New List(Of String)

        If ioChString <> "NULL" Then
            ioChExp = X0.CommaExpand(ioChString, 0, 63)
            LArray1 = Strings.Split(ioChExp, ",")

            For i = LBound(LArray1) To UBound(LArray1)
                If IsNumeric(LArray1(i)) Then
                    If CInt(LArray1(i)) > -1 And CInt(LArray1(i)) < 64 Then
                        vGrabDMXDirect.Add(Strings.Replace(inputString, "$ch=" & ioChString, "$ch=" & LArray1(i)))
                    End If
                End If
            Next i

        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | line " & thisLine & " | " & inputString & "][Invalid DMX instruction]")
            vGrabDMXDirect.Add("nop")
        End If

        GrabDMXDirect = vGrabDMXDirect
    End Function

    ' grabbers core 2

    'inteligent Device Grabber
    'main function
    Private Function Grab_Dev_1(ByVal inputString As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As String
        Dim P As Integer, i As Integer
        Dim LArray_T1() As String
        Dim LArray_GP() As String
        P = 0
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_T1) To UBound(LArray_T1)

            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "dev=") Then
                LArray_GP = Strings.Split(Strings.Replace(LArray_T1(i), " ", ""), "=")
                If IsNumeric(LArray_GP(1)) And LArray_GP(0) = "dev" Then
                    If CInt(LArray_GP(1)) > 0 And CInt(LArray_GP(1)) < (getDevLimit() - 2) Then ' 21 Then flexible dev
                        P = CInt(LArray_GP(1))
                    End If
                ElseIf (LArray_GP(1) = "server" Or LArray_GP(1) = "db" Or LArray_GP(1) = "ff") And LArray_GP(0) = "dev" Then
                    P = 255
                ElseIf (LArray_GP(1) = "last" Or LArray_GP(1) = "previous" Or LArray_GP(1) = "fe") And LArray_GP(0) = "dev" Then
                    P = 254
                ElseIf (Not IsNumeric(LArray_GP(1))) And LArray_GP(0) = "dev" Then
                    P = Grab_ID_from_Dev_Name(LArray_GP(1), thisRCU, ErrorWarnLog)
                End If
            End If
        Next i

        Grab_Dev_1 = xlWorkFunc.Dec2Hex(P, 2)
    End Function

    Private Function Grab_ID_from_Dev_Name(ByVal inputString As String, ByVal thisRCU As String, ByRef ErrorWarnLog As String()) As Integer
        Grab_ID_from_Dev_Name = 0
        'checking device validity
        Dim RcuDevList() As String, RcuDevNameList() As String, RCUDevType As String
        Dim DevEnumID As String, i As Integer

        RcuDevList = Strings.Split("01,02,03,04,04,05,05,06,06,07,07,08,0a", ",") 'should be in lcase for comparison
        RcuDevNameList = Strings.Split("idpg,tig,tag,ioexp,ioe,iodexp,iod,gsw,gs,ioexp,ioe,bsp,dali", ",")
        RCUDevType = "00"
        DevEnumID = "01"

        For i = LBound(RcuDevList) To UBound(RcuDevList)
            If Strings.InStr(inputString, RcuDevNameList(i)) Then
                RCUDevType = RcuDevList(i)
                If IsNumeric(Strings.Replace(inputString, RcuDevNameList(i), "")) Then
                    DevEnumID = xlWorkFunc.Dec2Hex(CInt(Strings.Replace(inputString, RcuDevNameList(i), "") - 1), 1) & "1"
                End If
                Exit For
            End If
        Next i

        If RCUDevType <> "00" Then
            Dim ce As Excel.Range
            For Each ce In xlDeviceSheet.Range("C4:C" & getDevLimit())
                If RCUDevType = "04" Then
                    If (Strings.LCase(ce.Text) = "04" Or Strings.LCase(ce.Text) = "07") And Strings.Replace(ce.Offset(0, -1).Text, " ", "") = DevEnumID Then
                        Grab_ID_from_Dev_Name = ce.Row - 3
                        Exit Function
                    End If
                Else
                    If Strings.LCase(ce.Text) = RCUDevType And Strings.Replace(ce.Offset(0, -1).Text, " ", "") = DevEnumID Then
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

    Private Function getDevLimit() As Integer
        Dim devSize As Integer
        devSize = xlDeviceSheet.Range("B4").End(Excel.XlDirection.xlDown).Row
        If devSize > 67 Then
            devSize = 67
        End If
        getDevLimit = devSize
    End Function


    'garbage management

    Private Sub releaseObject(ByRef thisObject As Object)
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
