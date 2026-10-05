Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX4
    Dim xlXLogSheet As Excel.Worksheet
    Dim xlXCODESheet As Excel.Worksheet
    Dim xlProgramSheet As Excel.Worksheet
    Dim xlSettingsSheet As Excel.Worksheet

    Dim RCUname As String

    Dim X0 As XCX0
    Dim X5 As New XCX5

    Sub XCODE_Finalize(thisUser As String, MemmapName As String, ByRef thisWorkBook As Excel.Workbook, ByRef xlSheetList As List(Of String), ByRef thisRCUList As List(Of String), unAttended As Boolean)

        X0 = New XCX0(unAttended)

        Check_ProgBase(thisWorkBook, thisRCUList)
        Set_Version(thisUser, MemmapName, thisWorkBook, xlSheetList, thisRCUList)

        ReleaseObject(xlXLogSheet)
        ReleaseObject(xlXCODESheet)
        ReleaseObject(xlProgramSheet)
        ReleaseObject(xlSettingsSheet)
        ReleaseObject(X0)
        ReleaseObject(X5)
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
    Private Sub Set_Version(thisUser As String, MemmapName As String, ByRef thisWorkbook As Excel.Workbook, ByRef xlSheetList As List(Of String), ByRef thisRCUList As List(Of String))
        '00 00 00 00 hex(XCODE version) 2Byte[bin7(year) bin4(month) bin5(day)] hex(build version)
        Dim buildVersion As Integer, tempVLine As String, vData() As String
        Dim firstRCU As String, XLogComment As String

        firstRCU = thisRCUList(0)

        xlProgramSheet = thisWorkbook.Sheets("Program" & firstRCU)

        tempVLine = xlProgramSheet.Range("F4").Text
        vData = Strings.Split(Strings.Mid(tempVLine, Strings.InStr(tempVLine, "[") + 1, Strings.InStr(tempVLine, "]") - Strings.InStr(tempVLine, "[") - 1), ","c)
        buildVersion = CInt(vData(1))

        Open_XLog(thisWorkbook, xlSheetList, firstRCU)

        XLogComment = "- Compiled using XC Terminal Ver-" & X0.GetXCVer & "." & X0.GetXCbuildVer 'fixed unattended comment

        If (XLogComment = "test" OrElse XLogComment = "beta") AndAlso buildVersion > -1 Then 'can't skip first built
            X0.ConsoleMsg("XC Progress:> Versioning > Version " & buildVersion)
            Increment_version(MemmapName, (buildVersion - 1), thisWorkbook, thisRCUList)
        Else
            If XLogComment.Length < 4 AndAlso buildVersion < 0 Then
                'first built without comment
                XLogComment = "- Initial built"
            ElseIf Len(XLogComment) < 4 AndAlso buildVersion > -1 Then
                XLogComment = "- Compiled without logging changes"
            End If

            X0.ConsoleMsg("XC Progress:> Versioning > Version " & (buildVersion + 1))
            Write_XLog(thisUser, MemmapName, thisWorkbook, buildVersion, XLogComment)
            Increment_version(MemmapName, buildVersion, thisWorkbook, thisRCUList)
        End If

        Close_XLog(thisWorkbook)

        'xlXCODESheet = thisWorkbook.Sheets("XCODE" & firstRCU)
        'xlXCODESheet.Select()

    End Sub

    Private Sub Open_XLog(ByRef thisWorkbook As Excel.Workbook, ByRef xlSheetList As List(Of String), firstRCU As String)
        If Not X0.isSheetExist_From_List(xlSheetList, "XLog") Then
            thisWorkbook.Sheets.Add(Before:=thisWorkbook.Sheets("Device" & firstRCU)).Name = "XLog"
            xlSheetList.Add("XLog")
            Populate_XLog(thisWorkbook)
        Else
            thisWorkbook.Sheets("XLog").Unprotect("1234XCODE5")
        End If
    End Sub

    Private Sub Populate_XLog(ByRef thisWorkbook As Excel.Workbook)
        xlXLogSheet = thisWorkbook.Sheets("XLog")

        'xlXLogSheet.Range("B2").FormulaR1C1 = "File"
        'xlXLogSheet.Range("B2").Offset(0, 1).FormulaR1C1 = "Version"
        'xlXLogSheet.Range("B2").Offset(0, 2).FormulaR1C1 = "User"
        'xlXLogSheet.Range("B2").Offset(0, 3).FormulaR1C1 = "Date"
        'xlXLogSheet.Range("B2").Offset(0, 4).FormulaR1C1 = "Change Log"

        xlXLogSheet.Range("B2").Resize(1, 5).Value = New Object(,) {{"File", "Version", "User", "Date", "Change Log"}}
    End Sub

    Private Sub Write_XLog(thisUser As String, MemmapName As String, ByRef thisWorkbook As Excel.Workbook, PreVersion As Integer, thisComment As String)
        xlXLogSheet = thisWorkbook.Sheets("XLog")
        'B3 is the fixed origin row = 3 

        If Strings.LCase(xlXLogSheet.Range("B3").Text) = MemmapName.ToLower() Then
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
                                                                     "'" & (PreVersion + 1).ToString("X2"),
                                                                     "'" & thisUser,
                                                                     "'" & Day(DateTime.Now) & "/" & Month(DateTime.Now) & "/" & Year(DateTime.Now),
                                                                     thisComment}}

        xlXLogSheet.Range("B3:F3").Font.Bold = False
        xlXLogSheet.Range("B3:F3").VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    End Sub

    Private Sub Close_XLog(ByRef thisWorkbook As Excel.Workbook)
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

    Private Sub Push_XLog_Down(ByRef thisWorkbook As Excel.Workbook, newLine As Integer)
        xlXLogSheet = thisWorkbook.Sheets("XLog")
        Dim newRow As Excel.Range = xlXLogSheet.Rows(newLine)
        newRow.Insert()
    End Sub

    Private Sub Push_code_down(ByRef thisXCODESheet As Excel.Worksheet, inputText As String)
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

    Private Sub Increment_version(MemmapName As String, buildVersion As Integer, ByRef thisWorkbook As Excel.Workbook, ByRef thisRCUList As List(Of String))
        Dim tempVLine As String
        Dim vLine As Integer, m As Integer
        Dim this_Color As Integer

        tempVLine = "00 00 00 00 " & CInt(CDbl(X0.GetXCVer) * 10).ToString("X2") & " " & X0.BuildDate_hex & " " & (buildVersion + 1).ToString("X2") & ";"

        For Each thisRCU As String In thisRCUList

            xlXCODESheet = thisWorkbook.Sheets("XCODE" & thisRCU)
            xlProgramSheet = thisWorkbook.Sheets("Program" & thisRCU)

            vLine = 1 : m = 3
            While vLine = 1 AndAlso m < 100
                If Strings.InStr(Strings.LCase(xlXCODESheet.Range("C" & m).Text), "version ") = 1 Then
                    vLine = m
                End If
                m = m + 1
            End While

            If vLine > 1 Then
                xlXCODESheet.Range("C" & vLine).FormulaR1C1 = "version [" & X0.Encript(X0.Getsufixname(MemmapName)) & "," & (buildVersion + 1) & "]"
                'xlProgramSheet.Range("F4").FormulaR1C1 = xlXCODESheet.Range("C" & vLine).Text
                'xlProgramSheet.Range("E4").FormulaR1C1 = tempVLine
                xlProgramSheet.Range("E4").Resize(1, 2).Value = New Object(,) {{tempVLine, xlXCODESheet.Range("C" & vLine).Text}}
                this_Color = X5.Set_C_Scheme_2(xlXCODESheet.Range("C1").Text)
                X5.Keyword_version(this_Color, xlXCODESheet.Range("C" & vLine))
            End If
        Next
    End Sub


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
