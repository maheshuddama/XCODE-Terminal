Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX1
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkFunc As Excel.WorksheetFunction
    Dim xlPath As String
    Dim startTime As String, thisUserName As String, eMsg As String
    Dim xlSheetList As New List(Of String)
    Dim thisRCUList As List(Of String)
    Dim RCUNameDictionary As New Dictionary(Of String, String)
    Dim IPDictionary As New Dictionary(Of String, String)

    Dim X0 As XCX0
    Dim X2 As New XCX2
    Dim X3 As New XCX3
    Dim X4 As New XCX4

    Sub XCM(MemmapName As String, ByRef ErrorWarnLog As String(), Optional unAttended As Boolean = False)
        X0 = New XCX0(unAttended)
        startTime = X0.getTimeStamp
        eMsg = ""

        If CreateWorkbook(MemmapName, ErrorWarnLog) Then
            xlApp.ScreenUpdating = False
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual

            X2.XCODE_Core1_Core2(MemmapName, xlWorkBook, xlWorkFunc, xlSheetList, thisRCUList, ErrorWarnLog, unAttended, RCUNameDictionary, IPDictionary)
            xlApp.Calculate()

            X3.XCODE_Core3(xlWorkBook, xlWorkFunc, thisRCUList, ErrorWarnLog, unAttended, IPDictionary)
            xlApp.Calculate()

            X4.XCODE_Finalize(thisUserName, MemmapName, xlWorkBook, xlWorkFunc, xlSheetList, thisRCUList, ErrorWarnLog, unAttended)
            xlApp.ScreenUpdating = True
            xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            UpdateReport(thisUserName, X0.getXCVer & "." & X0.getXCbuildVer, MemmapName)
            CreateSolutions(MemmapName, ErrorWarnLog)
            removeBackups()
            CompleteWorkbook() 'save excel
            CreateErrorWarnReport("XC compiled with error(s)", MemmapName, thisUserName, startTime, ErrorWarnLog, False)
            DumpResources(unAttended)
        Else
            CreateErrorWarnReport("XC did not compile due to critical error(s)", MemmapName, xlApp.UserName, startTime, ErrorWarnLog, True)
            DumpResources(unAttended)
        End If
        Console.WriteLine() 'finally move to a new line
    End Sub

    Private Function CreateWorkbook(MemmapName As String, ByRef ErrorWarnLog As String()) As Boolean
        CreateWorkbook = False

        xlPath = My.Computer.FileSystem.CurrentDirectory

        Console.Write("XC Progress:> Engaging Memmap file '" & MemmapName & "'...")

        xlApp = New Excel.Application()

        'try to open memmap
        Try
            xlWorkBook = xlApp.Workbooks.Open(xlPath & "\" & MemmapName)
        Catch ex As Exception
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), ex.Message)
            Exit Function
        End Try

        xlWorkFunc = xlApp.WorksheetFunction
        thisUserName = xlApp.UserName

        Populate_Sheet_List()
        thisRCUList = X0.getRCUList(xlSheetList)

        If thisRCUList(0) <> "NULL" Then

            If X0.NoVariableErrors(xlWorkBook, xlSheetList, thisRCUList, ErrorWarnLog) Then
                CreateWorkbook = CheckCompatibility(xlWorkBook, ErrorWarnLog)
            Else
                'there was some errors, its included in the error report
            End If
        Else
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "One or more sheets does not exist in the Mem map. A valid Mem map should be having at least one set of sheets of 'Device','XCODE','Program','Settings','CustomVar' or their integer elevated name such as 'Device1' etc. Carefully check in the Mem map if these sheet names having an extra space charactor before or end to the name.")
        End If
    End Function

    Private Sub Populate_Sheet_List()
        For Each xlSheet As Excel.Worksheet In xlWorkBook.Sheets
            xlSheetList.Add(xlSheet.Name)
        Next
    End Sub

    Private Sub CreateErrorWarnReport(header As String, MemmapName As String, thisUser As String, thisStartTime As String, ByRef ErrorWarnLog As String(), Optional Critical As Boolean = False)

        If Strings.Len(ErrorWarnLog(0)) > 5 OrElse Strings.Len(ErrorWarnLog(1)) > 5 Then
            'there's some error or warning to report.
            Dim fileh As System.IO.StreamWriter
            Dim errFile As String, endTime As String
            endTime = X0.getTimeStamp
            errFile = "Error Report - XC - " & MemmapName & " - [" & endTime & "].txt"

            fileh = My.Computer.FileSystem.OpenTextFileWriter(xlPath & "\" & errFile, False)

            With fileh
                .WriteLine(header & vbNewLine)
                .WriteLine("Project" & vbTab & vbTab & ": " & MemmapName)
                .WriteLine("Start time" & vbTab & ": " & X0.expandTimeStamp(startTime))
                .WriteLine("End time" & vbTab & ": " & X0.expandTimeStamp(endTime))
                .WriteLine("Compiler" & vbTab & ": " & "XC-Terminal version " & X0.getXCVer & "." & X0.getXCbuildVer)
                .WriteLine("User" & vbTab & vbTab & ": " & thisUser & vbNewLine & vbNewLine)
                If Strings.Len(ErrorWarnLog(0)) > 5 Then
                    .WriteLine("====================[Error Log]====================")
                    .WriteLine(ErrorWarnLog(0) & vbNewLine)
                End If
                If Strings.Len(ErrorWarnLog(1)) > 5 Then
                    .WriteLine("====================[Warning Log]====================")
                    .WriteLine(ErrorWarnLog(1))
                End If
                .Close()
            End With


            If Critical Then
                X0.ConsoleMsg("XC:> Error File '" & errFile & "'") 'overwrite
            Else
                X0.ConsoleMsg("XC:> Error File '" & errFile & "'", False) 'newline
            End If

            If eMsg = "[Result]" Then
                eMsg = "[Error,Result]"
            Else
                eMsg = "[Error]"
            End If

            releaseObject(fileh)
        End If

    End Sub

    Private Sub UpdateReport(thisUser As String, thisVer As String, MemmapName As String)
        Dim xlWorkSheet As Excel.Worksheet, XLogQuery As String
        Dim thisPath As String

        X0.ConsoleMsg("XC Progress:> Finalizing....")

        thisUser = X0.URLEncode(thisUser)
        thisVer = X0.URLEncode(thisVer)
        MemmapName = X0.URLEncode(MemmapName)
        thisPath = X0.URLEncode(xlPath)

        If X0.isSheetExist_From_List(xlSheetList, "ReportXlog") Then
            xlWorkSheet = xlWorkBook.Sheets("ReportXlog")
            xlWorkSheet.Range("A:Z").ClearContents()
        Else
            xlWorkBook.Sheets.Add(Before:=xlWorkBook.Sheets("XLog")).name = "ReportXlog"
            xlWorkSheet = xlWorkBook.Sheets("ReportXlog")
            xlSheetList.Add("ReportXlog")
        End If

        XLogQuery = X0.getLogAuth
        XLogQuery = XLogQuery & "?entry.84145384=" & thisUser & "&entry.1362894441=" & thisVer & "&entry.1003842767=" & MemmapName & "&entry.696172307=" & thisPath & "&fvv=1&partialResponse=%5Bnull%2Cnull%2C%228790538719799163109%22%5D&pageHistory=0&fbzx=8790538719799163109"

        Try
            With xlWorkSheet.QueryTables.Add(Connection:="URL;" & XLogQuery, Destination:=xlWorkSheet.Range("$A$1"))
                .Name = "XLogXC"
                .RefreshStyle = Excel.XlCellInsertionMode.xlOverwriteCells
                .WebSelectionType = Excel.XlWebSelectionType.xlSpecifiedTables
                .PreserveFormatting = False
                .BackgroundQuery = True
                .Refresh(BackgroundQuery:=False)
                .Delete()
            End With
        Catch ex As Exception

        Finally
            xlWorkSheet.Range("A:Z").ClearContents()
            xlApp.DisplayAlerts = False
            xlWorkBook.Sheets("ReportXlog").delete()
            xlSheetList.Remove("ReportXlog")
            releaseObject(xlWorkSheet) 'locally released 
            xlApp.DisplayAlerts = True
        End Try

    End Sub

    Private Function CheckCompatibility(ByRef thisWorkbook As Excel.Workbook, ByRef ErrorWarnLog As String()) As Boolean
        Dim xlWorkSheet As Excel.Worksheet, ce As Excel.Range, thisCon As String
        CheckCompatibility = False

        X0.ConsoleMsg("XC Progress:> Initiate Compiling...")

        If X0.isSheetExist_From_List(xlSheetList, "Compatible") Then
            xlWorkSheet = xlWorkBook.Sheets("Compatible")
            xlWorkSheet.Range("A:Z").ClearContents()
        Else
            xlWorkBook.Sheets.Add(Before:=xlWorkBook.Sheets("Device" & thisRCUList(0))).name = "Compatible" 'add compatibility sheet at the begining
            xlSheetList.Add("Compatible")
            xlWorkSheet = xlWorkBook.Sheets("Compatible")
        End If



        thisCon = X0.getOpenAuth

        Try
            With xlWorkSheet.QueryTables.Add(Connection:=thisCon, Destination:=xlWorkSheet.Range("$A$1"))
                .Name = "XC"
                .FieldNames = False
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = False
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = False
                .RefreshPeriod = 0
                .WebSelectionType = Excel.XlWebSelectionType.xlSpecifiedTables
                .WebFormatting = Excel.XlWebFormatting.xlWebFormattingNone
                .WebTables = "1"
                .WebPreFormattedTextToColumns = True
                .WebConsecutiveDelimitersAsOne = True
                .WebSingleBlockTextImport = False
                .WebDisableDateRecognition = False
                .WebDisableRedirections = False
                .Refresh(BackgroundQuery:=False)
                .Delete()
            End With

            For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("B1"), xlWorkSheet.Range("B1").End(Excel.XlDirection.xlDown))
                If ce.Text = "ver" & X0.getXCVer Then
                    If ce.Offset(0, 1).Text = "TRUE" Then
                        CheckCompatibility = True
                    Else
                        ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "Incompatible version. " & ce.Offset(0, 2).Text)
                    End If
                End If
            Next ce

        Catch ex As Exception
            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "Failed to initiate Memmap.")
        Finally
            xlWorkSheet.Range("A:Z").ClearContents()
            xlApp.DisplayAlerts = False
            xlWorkBook.Sheets("Compatible").delete()
            xlSheetList.Remove("Compatible")
            releaseObject(xlWorkSheet)
            xlApp.DisplayAlerts = True
        End Try

    End Function


    Private Sub removeBackups()
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        'remove backup
        xlApp.DisplayAlerts = False
        For Each thisRCU As String In thisRCUList
            xlWorkSheet = xlWorkBook.Sheets("FUNCTION" & thisRCU)
            xlWorkSheet.Range("A:Z").ClearContents()
            xlWorkBook.Sheets("FUNCTION" & thisRCU).delete()
        Next
        'making the first xcode sheet active
        xlWorkSheet = xlWorkBook.Sheets("XCODE" & thisRCUList(0))
        xlWorkSheet.Select()

        releaseObject(xlWorkSheet)
        xlApp.DisplayAlerts = True
    End Sub

    Private Sub DumpResources(Optional unAttended As Boolean = False)
        'release all objects in an error.
        Try
            If Not IsNothing(xlWorkBook) Then
                xlApp.DisplayAlerts = False
                xlWorkBook.Close()
                xlApp.DisplayAlerts = True
            End If
        Catch ex As Exception
            ' no exceptions
        Finally
            releaseObject(xlWorkBook)
            releaseObject(xlWorkFunc)

            releaseObject(X0)
            releaseObject(X2)
            releaseObject(X3)
            releaseObject(X4)

            If Not IsNothing(xlApp) Then
                xlApp.Quit()
            End If

            releaseObject(xlApp)

        End Try

        If unAttended AndAlso Strings.Len(eMsg) > 0 Then
            Console.WriteLine()
            Console.Write("XC:>" & eMsg)
            'MsgBox(eMsg)
        End If
    End Sub

    Private Sub CreateSolutions(MemmapName As String, ByRef ErrorWarnLog As String())
        Dim Notification As String, evalPath As Integer
        Dim Defined_Name As String, rcount As Integer, rfullcount As Integer

        Notification = "" : Defined_Name = "NULL"
        evalPath = X0.getFilePathFormat(xlPath)

        If evalPath > 0 Then

            'If evalPath = 2 Then
            '    Notification = "OneDrive synchronise pending" & vbNewLine & vbNewLine
            'End If

            rcount = 0
            rfullcount = thisRCUList.Count

            For Each thisRCU As String In thisRCUList
                Defined_Name = getDefinedName(thisRCU)

                Notification = Notification & " '" & DeviceXF(MemmapName, thisRCU, Defined_Name) & "'"
                Notification = Notification & " '" & ProgramXF(MemmapName, thisRCU, Defined_Name) & "'"
                Notification = Notification & " '" & ParamsXF(MemmapName, thisRCU, Defined_Name) & "'"

                rcount = rcount + 1
                X0.ConsoleProgress("Printer > Collecting pages - " & MemmapName, (rcount * 100 / rfullcount))
            Next
            X0.ConsoleMsg("XC:> Solution Files " & Notification)
            eMsg = "[Result]"
        Else
            ErrorWarnLog(1) = X0.addErrorOrWarn(ErrorWarnLog(1), "[XCODE | Printer | " & xlPath & "][Destination path could not access]")
            eMsg = "[Error]"
        End If
    End Sub

    Private Sub CompleteWorkbook()
        Try
            xlWorkBook.Save()

            'xlApp.DisplayAlerts = False
            'xlWorkBook.Close()
            'xlApp.DisplayAlerts = True
        Catch ex As Exception

        Finally
            'discard objects

            'releaseObject(xlWorkBook)
            'releaseObject(xlWorkFunc)

            'releaseObject(X0)
            'releaseObject(X2)
            'releaseObject(X3)
            'releaseObject(X4)

            'If Not IsNothing(xlApp) Then
            '    xlApp.Quit()
            'End If

            'releaseObject(xlApp)
        End Try
    End Sub

    'solution write

    Private Function getDefinedName(thisRCU As String) As String
        'Dim xlWorkSheet As Excel.Worksheet
        'Dim ce As Excel.Range, ThisTempLine As String
        'getDefinedName = "NULL"

        'xlWorkSheet = xlWorkBook.Sheets("XCODE" & thisRCU)

        'For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("B2"), xlWorkSheet.Range("B2").End(Excel.XlDirection.xlDown))

        '    ThisTempLine = Strings.LCase(Strings.Trim(ce.Offset(0, 1).Text))

        '    If Strings.InStr(ThisTempLine, "//") > 2 AndAlso Not Strings.Left(ThisTempLine, 2) = "0x" Then 'comment detection and filter except direct hex
        '        ThisTempLine = Strings.Trim(Strings.Left(ThisTempLine, Strings.InStr(ThisTempLine, "//") - 1))
        '    End If

        '    If Strings.InStr(ThisTempLine, "define rcu ") = 1 Then
        '        getDefinedName = X0.Grab_name(ThisTempLine)
        '        releaseObject(xlWorkSheet)
        '        Exit Function
        '    End If

        'Next ce

        'releaseObject(xlWorkSheet)
        If RCUNameDictionary.ContainsKey("XCODE" & thisRCU) Then
            getDefinedName = RCUNameDictionary("XCODE" & thisRCU)
        Else
            getDefinedName = "NULL"
        End If
    End Function

    Private Function DeviceXF(MemmapName As String, thisRCU As String, Def_Name As String) As String
        Dim xlWorkSheet As Excel.Worksheet
        Dim ce As Excel.Range, RCUName As String, thisName As String
        Dim withCBS As Boolean, CBSline As String, CBSint As Integer
        Dim fileh As System.IO.StreamWriter

        withCBS = False

        If thisRCU = "" Then
            RCUName = ""
        Else
            RCUName = "-RCU" & thisRCU
        End If

        If Def_Name = "NULL" Then
            thisName = X0.getsufixname(MemmapName) & RCUName
        Else
            thisName = Def_Name
        End If

        xlWorkSheet = xlWorkBook.Sheets("Device" & thisRCU)
        fileh = My.Computer.FileSystem.OpenTextFileWriter(X0.getfilepath(xlPath) & "\Devices." & thisName & ".hex", False)

        With fileh

            For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("B4"), xlWorkSheet.Range("B4").End(Excel.XlDirection.xlDown))
                If ce.Text <> "" Then
                    .WriteLine(ce.Text & vbTab & ce.Offset(0, 1).Text & vbTab & ce.Offset(0, 2).Text & vbTab & ce.Offset(0, 3).Text & vbTab & ce.Offset(0, 4).Text & vbTab & ce.Offset(0, 5).Text)

                    If Strings.LCase(Strings.Trim(ce.Offset(0, 1).Text)) = "0e" AndAlso (Not withCBS) Then
                        withCBS = True
                        CBSline = Strings.Replace(Strings.Trim(ce.Offset(0, 4).Text), " ", "")
                        CBSline = Strings.Right(CBSline, 2) & Strings.Left(CBSline, 2)
                        CBSint = xlWorkFunc.Hex2Dec(CBSline) + 4
                    End If
                End If
            Next ce

            If withCBS Then
                For Each ce In xlWorkSheet.Range("J4:J" & (CBSint - 1))
                    If ce.Text <> "" Then
                        .WriteLine(ce.Text)
                    End If
                Next ce
                For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("G" & CBSint), xlWorkSheet.Range("G" & CBSint).End(Excel.XlDirection.xlDown))
                    If ce.Text <> "" Then
                        .WriteLine(ce.Text & vbTab & ce.Offset(0, 1).Text & vbTab & ce.Offset(0, 2).Text & vbTab & ce.Offset(0, 3).Text)
                    End If
                Next ce
            Else
                For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("J4"), xlWorkSheet.Range("J4").End(Excel.XlDirection.xlDown))
                    If ce.Text <> "" Then
                        .WriteLine(ce.Text)
                    End If
                Next ce
            End If

            .Close()
        End With

        releaseObject(fileh)
        releaseObject(xlWorkSheet)

        DeviceXF = "Devices." & thisName & ".hex"
    End Function

    Private Function ProgramXF(MemmapName As String, thisRCU As String, Def_Name As String) As String
        Dim xlWorkSheet As Excel.Worksheet
        Dim ce As Excel.Range, RCUName As String, thisName As String
        Dim fileh As System.IO.StreamWriter

        If thisRCU = "" Then
            RCUName = ""
        Else
            RCUName = "-RCU" & thisRCU
        End If

        If Def_Name = "NULL" Then
            thisName = X0.getsufixname(MemmapName) & RCUName
        Else
            thisName = Def_Name
        End If

        xlWorkSheet = xlWorkBook.Sheets("Program" & thisRCU)
        fileh = My.Computer.FileSystem.OpenTextFileWriter(X0.getfilepath(xlPath) & "\Program." & thisName & ".hex", False)

        With fileh
            For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("E1"), xlWorkSheet.Range("E1").End(Excel.XlDirection.xlDown))
                If ce.Text <> "" Then
                    .WriteLine(ce.Text)
                End If
            Next ce
            .Close()
        End With


        releaseObject(fileh)
        releaseObject(xlWorkSheet)

        ProgramXF = "Program." & thisName & ".hex"
    End Function

    Private Function ParamsXF(MemmapName As String, thisRCU As String, Def_Name As String) As String
        Dim xlWorkSheet As Excel.Worksheet
        Dim ce As Excel.Range, RCUName As String, thisName As String
        Dim fileh As System.IO.StreamWriter

        If thisRCU = "" Then
            RCUName = ""
        Else
            RCUName = "-RCU" & thisRCU
        End If

        If Def_Name = "NULL" Then
            thisName = X0.getsufixname(MemmapName) & RCUName
        Else
            thisName = Def_Name
        End If

        xlWorkSheet = xlWorkBook.Sheets("Settings" & thisRCU)
        fileh = My.Computer.FileSystem.OpenTextFileWriter(X0.getfilepath(xlPath) & "\Params." & thisName & ".hex", False)

        With fileh
            For Each ce In xlWorkSheet.Range(xlWorkSheet.Range("B1"), xlWorkSheet.Range("B1").End(Excel.XlDirection.xlDown))
                If ce.Text <> "" Then
                    .WriteLine(ce.Text)
                End If
            Next ce
            .Close()
        End With

        releaseObject(fileh)
        releaseObject(xlWorkSheet)

        ParamsXF = "Params." & thisName & ".hex"
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
