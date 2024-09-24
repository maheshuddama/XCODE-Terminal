Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX4
    Dim xlXLogSheet As Excel.Worksheet
    Dim xlXCODESheet As Excel.Worksheet
    Dim xlProgramSheet As Excel.Worksheet
    Dim xlSettingsSheet As Excel.Worksheet
    Dim xlWorkFunc As Excel.WorksheetFunction

    Dim RCUname As String

    Dim X0 As XCX0
    Dim X5 As New XCX5

    Sub XCODE_Finalize(ByVal thisUser As String, ByVal MemmapName As String, ByRef thisWorkBook As Excel.Workbook, ByRef thisxlWorkFunc As Excel.WorksheetFunction,
                       ByRef xlSheetList As List(Of String), ByRef thisRCUList As List(Of String), ByRef ErrorWarnLog As String(), ByVal unAttended As Boolean)
        xlWorkFunc = thisxlWorkFunc
        X0 = New XCX0(thisxlWorkFunc, unAttended)

        Check_ProgBase(thisWorkBook, thisRCUList)
        Set_Version(thisUser, MemmapName, thisWorkBook, xlSheetList, thisRCUList)

        releaseObject(xlXLogSheet)
        releaseObject(xlXCODESheet)
        releaseObject(xlProgramSheet)
        releaseObject(xlSettingsSheet)
        releaseObject(X0)
        releaseObject(X5)
    End Sub

    Private Sub Check_ProgBase(ByRef thisWorkbook As Excel.Workbook, ByRef thisRCUList As List(Of String))
        Dim thisRCU As String, pfullcount As Integer
        Dim i As Integer, ProgBaseVal As String

        i = 0 : ProgBaseVal = "128"

        X0.ConsoleMsg("XC Progress:> Checking Setting.......")

        For Each thisRCU In thisRCUList

            If thisRCU = "" Then
                RCUname = ""
            Else
                RCUname = "RCU" & thisRCU & " > "
            End If

            xlProgramSheet = thisWorkbook.Sheets("Program" & thisRCU)
            xlSettingsSheet = thisWorkbook.Sheets("Settings" & thisRCU)

            pfullcount = xlProgramSheet.Range("E4").End(Excel.XlDirection.xlDown).Row

            If pfullcount > 2002 Then
                ProgBaseVal = "0"
            End If

            i = i + 1
            X0.ConsoleProgress(RCUname & "Checking Setting", (i * 100 / thisRCUList.Count))
        Next

        'need to work on 2 loops to determine if at least one rcu script is above the standard block size
        i = 0
        For Each thisRCU In thisRCUList

            If thisRCU = "" Then
                RCUname = ""
            Else
                RCUname = "RCU" & thisRCU & " > "
            End If

            xlSettingsSheet = thisWorkbook.Sheets("Settings" & thisRCU)

            If xlSettingsSheet.Range("C26").Text <> ProgBaseVal Then 'optimization write only if its different.
                'xlSettingsSheet.Range("B26").FormulaR1C1 = "=DEC2HEX(RC[1],2)&"";"""
                'xlSettingsSheet.Range("C26").FormulaR1C1 = ProgBaseVal

                xlSettingsSheet.Range("B26").Resize(1, 2).Value = New Object(,) {{"=DEC2HEX(RC[1],2)&"";""", ProgBaseVal}}
            End If

            i = i + 1
            X0.ConsoleProgress(RCUname & "Formatter > Updating ProgBase Setting", (i * 100 / thisRCUList.Count))
        Next

    End Sub

    'advancing the version for XCODE terminal this is unattended version update no user inputs
    Private Sub Set_Version(ByVal thisUser As String, ByVal MemmapName As String, ByRef thisWorkbook As Excel.Workbook, ByRef xlSheetList As List(Of String), ByRef thisRCUList As List(Of String))
        '00 00 00 00 hex(XCODE version) 2Byte[bin7(year) bin4(month) bin5(day)] hex(build version)
        Dim buildVersion As Integer, tempVLine As String, vData() As String
        Dim firstRCU As String, XLogComment As String

        firstRCU = thisRCUList(0)

        xlProgramSheet = thisWorkbook.Sheets("Program" & firstRCU)

        tempVLine = xlProgramSheet.Range("F4").Text
        vData = Strings.Split(Strings.Mid(tempVLine, Strings.InStr(tempVLine, "[") + 1, Strings.InStr(tempVLine, "]") - Strings.InStr(tempVLine, "[") - 1), ",")
        buildVersion = CInt(vData(1))

        open_XLog(thisWorkbook, xlSheetList, firstRCU)

        XLogComment = "- Compiled using XC Terminal Ver-" & X0.getXCVer & "." & X0.getXCbuildVer 'fixed unattended comment

        If (XLogComment = "test" Or XLogComment = "beta") AndAlso buildVersion > -1 Then 'can't skip first built
            X0.ConsoleMsg("XC Progress:> Versioning > Version " & buildVersion)
            increment_version(MemmapName, (buildVersion - 1), thisWorkbook, thisRCUList)
        Else
            If Len(XLogComment) < 4 AndAlso buildVersion < 0 Then
                'first built without comment
                XLogComment = "- Initial built"
            ElseIf Len(XLogComment) < 4 AndAlso buildVersion > -1 Then
                XLogComment = "- Compiled without logging changes"
            End If

            X0.ConsoleMsg("XC Progress:> Versioning > Version " & (buildVersion + 1))
            write_XLog(thisUser, MemmapName, thisWorkbook, buildVersion, XLogComment)
            increment_version(MemmapName, buildVersion, thisWorkbook, thisRCUList)
        End If

        close_XLog(thisWorkbook)

        'xlXCODESheet = thisWorkbook.Sheets("XCODE" & firstRCU)
        'xlXCODESheet.Select()

    End Sub

    Private Sub open_XLog(ByRef thisWorkbook As Excel.Workbook, ByRef xlSheetList As List(Of String), ByVal firstRCU As String)
        If Not X0.isSheetExist_From_List(xlSheetList, "XLog") Then
            thisWorkbook.Sheets.Add(Before:=thisWorkbook.Sheets("Device" & firstRCU)).Name = "XLog"
            xlSheetList.Add("XLog")
            populate_XLog(thisWorkbook)
        Else
            thisWorkbook.Sheets("XLog").Unprotect("1234XCODE5")
        End If
    End Sub

    Private Sub populate_XLog(ByRef thisWorkbook As Excel.Workbook)
        xlXLogSheet = thisWorkbook.Sheets("XLog")

        'xlXLogSheet.Range("B2").FormulaR1C1 = "File"
        'xlXLogSheet.Range("B2").Offset(0, 1).FormulaR1C1 = "Version"
        'xlXLogSheet.Range("B2").Offset(0, 2).FormulaR1C1 = "User"
        'xlXLogSheet.Range("B2").Offset(0, 3).FormulaR1C1 = "Date"
        'xlXLogSheet.Range("B2").Offset(0, 4).FormulaR1C1 = "Change Log"

        xlXLogSheet.Range("B2").Resize(1, 5).Value = New Object(,) {{"File", "Version", "User", "Date", "Change Log"}}
    End Sub

    Private Sub write_XLog(ByVal thisUser As String, ByVal MemmapName As String, ByRef thisWorkbook As Excel.Workbook, ByVal PreVersion As Integer, ByVal thisComment As String)
        xlXLogSheet = thisWorkbook.Sheets("XLog")
        'B3 is the fixed origin row = 3 

        If Strings.LCase(xlXLogSheet.Range("B3").Text) = Strings.LCase(MemmapName) Then
            'wrting next log
            xlXLogSheet.Range("B3").ClearContents()
            Push_XLog_Down(thisWorkbook, 3)
        Else
            'can be new file or adding Xlog for the first time
            Push_XLog_Down(thisWorkbook, 3)
            Push_XLog_Down(thisWorkbook, 3)
            Push_XLog_Down(thisWorkbook, 3)
            Push_XLog_Down(thisWorkbook, 3)
            Push_XLog_Down(thisWorkbook, 3)

            xlXLogSheet.Range("B7").FormulaR1C1 = "Historical Data"
            With xlXLogSheet.Range("B7:F1000").Font
                .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
                .TintAndShade = -0.349986266670736
            End With
        End If

        'xlXLogSheet.Range("B3").FormulaR1C1 = MemmapName
        'xlXLogSheet.Range("B3").Offset(0, 1).FormulaR1C1 = "'" & xlWorkFunc.Dec2Hex(PreVersion + 1, 2)
        'xlXLogSheet.Range("B3").Offset(0, 2).FormulaR1C1 = "'" & thisUser
        'xlXLogSheet.Range("B3").Offset(0, 3).FormulaR1C1 = "'" & Day(DateTime.Now) & "/" & Month(DateTime.Now) & "/" & Year(DateTime.Now)
        'xlXLogSheet.Range("B3").Offset(0, 4).FormulaR1C1 = thisComment

        xlXLogSheet.Range("B3").Resize(1, 5).Value = New Object(,) {{MemmapName,
                                                                     "'" & xlWorkFunc.Dec2Hex(PreVersion + 1, 2),
                                                                     "'" & thisUser,
                                                                     "'" & Day(DateTime.Now) & "/" & Month(DateTime.Now) & "/" & Year(DateTime.Now),
                                                                     thisComment}}

        xlXLogSheet.Range("B3:F3").Font.Bold = False
        xlXLogSheet.Range("B3:F3").VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    End Sub

    Private Sub close_XLog(ByRef thisWorkbook As Excel.Workbook)
        xlXLogSheet = thisWorkbook.Sheets("XLog")

        xlXLogSheet.Range("B2:F2").Font.Bold = True
        xlXLogSheet.Columns("C:C").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        xlXLogSheet.Range("B2:F2").EntireColumn.AutoFit()
        xlXLogSheet.Range("F:F").NumberFormat = "@"
        xlXLogSheet.Range("F:F").ColumnWidth = 64
        xlXLogSheet.Range("3:3").Rows.AutoFit()
        'protect and close
        xlXLogSheet.Protect(Password:="1234XCODE5")
    End Sub

    Private Sub Push_XLog_Down(ByRef thisWorkbook As Excel.Workbook, ByVal newLine As Integer)
        xlXLogSheet = thisWorkbook.Sheets("XLog")
        Dim newRow As Excel.Range = xlXLogSheet.Rows(newLine)
        newRow.Insert()
    End Sub

    Private Sub push_code_down(ByRef thisXCODESheet As Excel.Worksheet, ByVal inputText As String)
        If Strings.Len(thisXCODESheet.ActiveCell.Offset(1, 0).Text) > 0 Then
            thisXCODESheet.CutCopyMode = False
            thisXCODESheet.ActiveCell.Offset(1, 1).Range("A1").Select()
            thisXCODESheet.ActiveCell.Rows("1:1").EntireRow.Select()
            thisXCODESheet.Selection.Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
            thisXCODESheet.ActiveCell.Offset(-1, 1).Range("A1").Select()
            thisXCODESheet.Selection.AutoFill(Destination:=thisXCODESheet.ActiveCell.Range("A1:A3"), Type:=Excel.XlAutoFillType.xlFillDefault)
            thisXCODESheet.ActiveCell.Offset(1, 1).Range("A1").Select()
            thisXCODESheet.ActiveCell.FormulaR1C1 = inputText
        Else
            thisXCODESheet.ActiveCell.Offset(1, 0).Select()
            thisXCODESheet.ActiveCell.FormulaR1C1 = inputText
        End If
    End Sub

    Private Sub increment_version(ByVal MemmapName As String, ByVal buildVersion As Integer, ByRef thisWorkbook As Excel.Workbook, ByRef thisRCUList As List(Of String))
        Dim tempVLine As String
        Dim vLine As Integer, m As Integer
        Dim this_Color As Integer

        tempVLine = "00 00 00 00 " & xlWorkFunc.Dec2Hex(CDbl(X0.getXCVer) * 10, 2) & " " & X0.buildDate_hex & " " & xlWorkFunc.Dec2Hex(buildVersion + 1, 2) & ";"

        For Each thisRCU As String In thisRCUList

            xlXCODESheet = thisWorkbook.Sheets("XCODE" & thisRCU)
            xlProgramSheet = thisWorkbook.Sheets("Program" & thisRCU)

            vLine = 1 : m = 3
            While (vLine = 1 AndAlso m < 100)
                If Strings.InStr(Strings.LCase(xlXCODESheet.Range("C" & m).Text), "version ") = 1 Then
                    vLine = m
                End If
                m = m + 1
            End While

            If vLine > 1 Then
                xlXCODESheet.Range("C" & vLine).FormulaR1C1 = "version [" & X0.encript(X0.getsufixname(MemmapName)) & "," & (buildVersion + 1) & "]"
                'xlProgramSheet.Range("F4").FormulaR1C1 = xlXCODESheet.Range("C" & vLine).Text
                'xlProgramSheet.Range("E4").FormulaR1C1 = tempVLine
                xlProgramSheet.Range("E4").Resize(1, 2).Value = New Object(,) {{tempVLine, xlXCODESheet.Range("C" & vLine).Text}}
                this_Color = X5.set_C_Scheme_2(xlXCODESheet.Range("C1").Text)
                X5.keyword_version(this_Color, xlXCODESheet.Range("C" & vLine))
            End If
        Next
    End Sub


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
