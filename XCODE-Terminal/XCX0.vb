Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions

Public Class XCX0
    Dim xlWorkFunc As Excel.WorksheetFunction

    'constructors

    Public Sub New()

    End Sub

    Public Sub New(ByRef ThisWorkFunc As Excel.WorksheetFunction)
        xlWorkFunc = ThisWorkFunc
    End Sub

    ' XC terminal elements

    Sub XC_Terminal_Header()
        Dim logo As String = "" & vbNewLine & _
         "    ██╗  ██╗ ██████╗" & vbNewLine & _
         "    ╚██╗██╔╝██╔════╝" & vbNewLine & _
         "     ╚███╔╝ ██║     " & vbNewLine & _
         "     ██╔██╗ ██║     " & vbNewLine & _
         "    ██╔╝ ██╗╚██████╗" & vbNewLine & _
         "    ╚═╝  ╚═╝ ╚═════╝"
        Console.WriteLine(logo & vbNewLine & "     XC-Terminal Ver(" & getXCVer() & "." & getXCbuildVer() & ")" & vbNewLine)
    End Sub

    Sub ConsoleProgress(ByVal StageDiscription As String, ByVal StageProgress As Double)
        Dim progressBarWidth As Integer = 20
        ConsoleEraseThisLine()
        Console.Write("XC Progress:> " & StageDiscription & "> [{0}{1}]{2}%", New String("█"c, StageProgress * progressBarWidth / 100), New String(" "c, progressBarWidth - (StageProgress * progressBarWidth / 100)), CInt(StageProgress))
    End Sub

    Sub ConsoleMsg(ByVal thisMessage As String, Optional ByVal Overwrite As Boolean = True)
        If Overwrite Then
            ConsoleEraseThisLine()
            Console.Write(thisMessage)
        Else
            Console.WriteLine()
            Console.Write(thisMessage)
        End If
    End Sub

    Sub ConsoleEraseThisLine()
        Dim thisLenth As Integer = Console.CursorLeft
        Console.SetCursorPosition(0, Console.CursorTop)
        Console.Write(New String(" "c, thisLenth))
        Console.SetCursorPosition(0, Console.CursorTop)
    End Sub

    Function get_memmap_name_from_words(ByVal inputWords As String(), Optional ByVal index As Integer = 1) As String
        Dim w As Integer
        get_memmap_name_from_words = inputWords(index)

        If UBound(inputWords) > index Then
            For w = (index + 1) To UBound(inputWords)
                get_memmap_name_from_words = get_memmap_name_from_words & " " & inputWords(w)
                If Strings.Right(Strings.LCase(inputWords(w)), 5) = ".xlsx" Then
                    Exit For
                End If
            Next w
        End If
    End Function

    ' Mem map evaluation
    Function getRCUList(ByRef thisWorkbook As Excel.Workbook) As List(Of String)
        Dim L As Integer
        Dim thisRCUList As New List(Of String)

        For L = 1 To 64
            If isSheetExist(thisWorkbook, "Device" & L) AndAlso
                isSheetExist(thisWorkbook, "XCODE" & L) AndAlso
                isSheetExist(thisWorkbook, "Program" & L) AndAlso
                isSheetExist(thisWorkbook, "Settings" & L) AndAlso
                isSheetExist(thisWorkbook, "CustomVar" & L) Then
                thisRCUList.Add(L)
            End If
            ConsoleProgress("Listing RCUs", CInt((L / 64) * 100))
        Next L

        If thisRCUList.Count = 0 Then
            If isSheetExist(thisWorkbook, "Device") AndAlso
                isSheetExist(thisWorkbook, "XCODE") AndAlso
                isSheetExist(thisWorkbook, "Program") AndAlso
                isSheetExist(thisWorkbook, "Settings") AndAlso
                isSheetExist(thisWorkbook, "CustomVar") Then
                thisRCUList.Add("")
            Else
                thisRCUList.Add("NULL")
            End If
        End If

        getRCUList = thisRCUList

    End Function

    Function NoVariableErrors(ByRef thisWorkbook As Excel.Workbook, ByRef thisRCUList As List(Of String), ByRef ErrorWarnLog As String()) As Boolean
        NoVariableErrors = True
        Dim xlVarSheet As Excel.Worksheet
        Dim ce As Excel.Range
        Dim C_VarShort As String, C_VarLong As String
        Dim M_VarShort As String, M_VarLong As String

        For Each thisRCU As String In thisRCUList


            xlVarSheet = thisWorkbook.Sheets("CustomVar" & thisRCU) 'this sheet should exist because its being checked before
            C_VarShort = "" : C_VarLong = ","

            If isVarSheetNotEmpty(xlVarSheet) Then
                For Each ce In xlVarSheet.Range(xlVarSheet.Range("B3").Offset(1, 0), xlVarSheet.Range("B3").End(Excel.XlDirection.xlDown))

                    If Not isValidHexString(ce.Text) Then
                        ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Non-Hexadecimal value as variable short name]")
                        NoVariableErrors = True
                    End If

                    If Strings.InStr(Strings.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                        ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Custom Variable definition with spaces]")
                        NoVariableErrors = True
                    End If

                    If Strings.InStr(C_VarShort, Strings.Right("00" & Strings.LCase(ce.Text), 2) & ",") > 0 Then
                        ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Duplicate variable]")
                        NoVariableErrors = True
                    Else
                        C_VarShort = C_VarShort & Strings.Right("00" & Strings.LCase(ce.Text), 2) & ","
                    End If

                    If Strings.Len(ce.Offset(0, 1).Text) > 0 Then
                        If Strings.InStr(C_VarLong, "," & Strings.LCase(ce.Offset(0, 1).Text) & ",") > 0 Then
                            ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Duplicate variable definition]")
                            NoVariableErrors = True
                        Else
                            C_VarLong = C_VarLong & Strings.LCase(ce.Offset(0, 1).Text) & ","
                        End If
                    Else
                        ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Empty variable definition]")
                        NoVariableErrors = True
                    End If
                Next ce
            End If


            'optional check for Modbus var

            If isSheetExist(thisWorkbook, "ModbusVar" & thisRCU) Then

                xlVarSheet = thisWorkbook.Sheets("ModbusVar" & thisRCU)
                M_VarShort = "" : M_VarLong = ","

                If isVarSheetNotEmpty(xlVarSheet) Then
                    For Each ce In xlVarSheet.Range(xlVarSheet.Range("B3").Offset(1, 0), xlVarSheet.Range("B3").End(Excel.XlDirection.xlDown))

                        If Not isValidHexString(ce.Text) Then
                            ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Non-Hexadecimal value as variable short name]")
                            NoVariableErrors = False
                        End If

                        If Strings.InStr(Strings.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                            ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Modbus Variable definition with spaces]")
                            NoVariableErrors = False
                        End If

                        If Strings.InStr(M_VarShort, Strings.Right("00" & Strings.LCase(ce.Text), 2) & ",") > 0 Then
                            ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Duplicate variable]")
                            NoVariableErrors = False
                        Else
                            M_VarShort = M_VarShort & Strings.Right("00" & Strings.LCase(ce.Text), 2) & ","
                        End If

                        If Strings.Len(ce.Offset(0, 1).Text) > 0 Then
                            If Strings.InStr(M_VarLong, "," & Strings.LCase(ce.Offset(0, 1).Text) & ",") > 0 Then
                                ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Duplicate variable definition]")
                                NoVariableErrors = False
                            Else
                                M_VarLong = M_VarLong & Strings.LCase(ce.Offset(0, 1).Text) & ","
                            End If
                        Else
                            ErrorWarnLog(0) = addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Empty variable definition]")
                            NoVariableErrors = False
                        End If
                    Next ce
                End If


            End If

        Next

    End Function

    Function isVarSheetNotEmpty(ByRef thisWorkSheet As Excel.Worksheet) As Boolean
        isVarSheetNotEmpty = False
        If Strings.Len(thisWorkSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
            isVarSheetNotEmpty = True
        End If
    End Function

    Function isSheetExist(ByRef thisWorkBook As Excel.Workbook, ByVal thisSheet As String) As Boolean
        isSheetExist = False

        Dim xlSheet As Excel.Worksheet

        For Each xlSheet In thisWorkBook.Sheets
            If xlSheet.Name = thisSheet Then
                isSheetExist = True
                Exit For
            End If
        Next xlSheet

    End Function

    Function isValidHexString(ByVal inputString As String) As Boolean
        isValidHexString = Regex.IsMatch(inputString, "^[0-9A-Fa-f]+$")
    End Function

    Function isValidBinString(ByVal inputString As String) As Boolean
        isValidBinString = Regex.IsMatch(inputString, "^[01]+$")
    End Function

    Function Evaluate(ByVal thisExpression As String) As Double
        Dim table As New DataTable()
        Dim result As Object = table.Compute(thisExpression, Nothing)
        Evaluate = Convert.ToDouble(result)
    End Function

    Function getFilePathFormat(ByVal filePath As String) As Integer 'added in XCODE 4.9.2, detect if local path accessible
        '0 other not accessible, 1 as local path, 2 as onedrive path,
        'removed letter c checking from path. in XCODE ver 5.2
        filePath = Strings.LCase(filePath)
        If InStr(filePath, ":\") = 2 Then
            getFilePathFormat = 1
        ElseIf InStr(filePath, "http") = 1 Then
            Dim onedrivePath As String
            onedrivePath = LCase(Environ("onedrive"))
            If InStr(onedrivePath, ":\") = 2 Then
                getFilePathFormat = 2
            Else
                getFilePathFormat = 0
            End If
        Else
            getFilePathFormat = 0
        End If
    End Function

    Function getfilepath(ByVal filePath As String, Optional ByVal toLower As Boolean = True) As String 'added in XCODE 4.9.2, redirect to onedrive local path
        If toLower Then
            filePath = Strings.LCase(filePath)
        End If

        If InStr(filePath, "http") = 1 Then
            Dim onedrivePath As String
            onedrivePath = Environ("onedrive")
            getfilepath = Strings.Replace(onedrivePath & Mid(filePath, InStr(filePath, "Documents") + 9), "/", "\")
        Else
            getfilepath = filePath
        End If
    End Function

    Private Function isValidDate(ByVal inputString As String) As Boolean
        Dim DArr() As String, Y1 As Integer, M1 As Integer, D1 As Integer
        isValidDate = False
        If Strings.InStr(inputString, "/") > 2 Then
            DArr = Strings.Split(inputString, "/")
            If UBound(DArr) = 2 Then
                If IsNumeric(DArr(0)) And IsNumeric(DArr(1)) And IsNumeric(DArr(2)) Then
                    Y1 = CInt(Strings.Right("00" & DArr(0), 2))
                    M1 = CInt(DArr(1))
                    D1 = CInt(DArr(2))

                    If Y1 > 10 And Y1 < 100 And M1 > 0 And M1 < 13 And D1 > 0 And D1 < 32 Then
                        isValidDate = True
                    End If
                End If
            End If
        End If
    End Function

    Private Function isValidTime(ByVal inputString As String) As Boolean
        Dim TArr() As String, H1 As Integer, M1 As Integer, S1 As Integer
        isValidTime = False
        If Strings.InStr(inputString, ":") > 2 Then
            TArr = Strings.Split(inputString, ":")
            If UBound(TArr) = 2 Then
                If IsNumeric(TArr(0)) And IsNumeric(TArr(1)) And IsNumeric(TArr(2)) Then
                    H1 = CInt(TArr(0))
                    M1 = CInt(TArr(1))
                    S1 = CInt(TArr(2))

                    If H1 < 24 And M1 < 60 And S1 < 60 Then
                        isValidTime = True
                    End If
                End If
            End If
        End If
    End Function

    Function CommaExpand(ByVal inputString As String, ByVal lowerLimit As Integer, ByVal upperLimit As Integer) As String
        Dim relayvalues() As String, rangeSplit() As String
        Dim Lrelay As Integer, Urelay As Integer, i As Integer, k As Integer
        Dim ArrExp(upperLimit - lowerLimit) As Integer

        If Strings.InStr(inputString, ".") > 0 Then
            'addWarn("[" & inputString & "][A dot detected instead of a comma]")
            inputString = Strings.Replace(inputString, ".", ",")
        End If

        If Strings.InStr(inputString, "to") > 0 Then
            inputString = Strings.Replace(inputString, "to", "-")
        End If

        relayvalues = Strings.Split(inputString, ",")
        CommaExpand = ""

        For i = LBound(relayvalues) To UBound(relayvalues)

            If Strings.InStr(relayvalues(i), "-") Then

                rangeSplit = Strings.Split(relayvalues(i), "-")

                If IsNumeric(rangeSplit(0)) And (Not IsNumeric(rangeSplit(1))) Then

                    If CInt(rangeSplit(0)) >= lowerLimit And CInt(rangeSplit(0)) <= upperLimit Then
                        ArrExp(CInt(rangeSplit(0)) - lowerLimit) = 1
                    End If

                ElseIf IsNumeric(rangeSplit(1)) And (Not IsNumeric(rangeSplit(0))) Then

                    If CInt(rangeSplit(1)) >= lowerLimit And CInt(rangeSplit(1)) <= upperLimit Then
                        ArrExp(CInt(rangeSplit(1)) - lowerLimit) = 1
                    End If

                ElseIf (Not IsNumeric(rangeSplit(1))) And (Not IsNumeric(rangeSplit(0))) Then

                ElseIf IsNumeric(rangeSplit(0)) And IsNumeric(rangeSplit(1)) Then
                    Lrelay = CInt(rangeSplit(0))
                    Urelay = CInt(rangeSplit(1))
                    If Lrelay > Urelay Then
                        Lrelay = Lrelay + Urelay
                        Urelay = Lrelay - Urelay
                        Lrelay = Lrelay - Urelay
                    End If

                    If Urelay > upperLimit Then
                        Urelay = upperLimit
                    End If
                    If Lrelay < lowerLimit Then
                        Lrelay = lowerLimit
                    End If

                    For k = Lrelay To Urelay
                        ArrExp(k - lowerLimit) = 1
                    Next k
                End If
            ElseIf IsNumeric(relayvalues(i)) Then
                If CInt(relayvalues(i)) >= lowerLimit And CInt(relayvalues(i)) <= upperLimit Then
                    ArrExp(CInt(relayvalues(i)) - lowerLimit) = 1
                End If
            End If

        Next i

        For i = LBound(ArrExp) To UBound(ArrExp)

            If ArrExp(i) = 1 Then
                CommaExpand = CommaExpand & "," & CStr(i + lowerLimit)
            End If

        Next i

        If CommaExpand = "" Then
            CommaExpand = 0
        Else
            CommaExpand = CommaExpand.Substring(1)
        End If

    End Function

    Private Function hexadder(ByVal hex1 As String, ByVal hex2 As String) As String
        hexadder = Strings.Right("00" & Hex(CLng(xlWorkFunc.Hex2Dec(hex1)) + CLng(xlWorkFunc.Hex2Dec(hex2))), 2)
    End Function

    'encription/decription
    '53 70 84 85
    '84 101 115 116
    'T e s t

    Function encript(ByVal inputstring As String) As String
        Dim i As Integer
        encript = ""
        For i = 1 To Strings.Len(inputstring)
            encript = encript & Strings.Right("00" & (Strings.Asc(Strings.Mid(inputstring, i, 1)) - 31), 2)
        Next i
    End Function

    Function decript(ByVal encripted As String) As String
        Dim x As Integer
        decript = ""
        For x = 1 To Strings.Len(encripted) / 2
            decript = decript & Chr(Strings.Mid(encripted, 2 * x - 1, 2) + 31)
        Next x
    End Function

    'connection
    Function getOpenAuth() As String
        getOpenAuth = decript("5451452873858581842716166980688415728080727770156880781684818370666984737070858416691618145070488739197279688559537667241418642181194948498977775685225117896456908558257081681670697485328684813084736683747972")
    End Function

    Function getLogAuth() As String
        getLogAuth = decript("73858581842716166980688415728080727770156880781671808378841669167016183934428150455270505620815026668339583652646784672179855519377857495358395651568856408138785034724650429123883416718083785170848180798470")
    End Function

    'URL encode
    Function URLEncode(ByVal StringToEncode As String, Optional ByVal PlusSpace As Boolean = True) As String
        Dim TempAns As String = "", CurChr As Integer = 1

        Do Until CurChr - 1 = Len(StringToEncode)
            Select Case Asc(Strings.Mid(StringToEncode, CurChr, 1))
                Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122
                    TempAns = TempAns & Strings.Mid(StringToEncode, CurChr, 1)
                Case 32
                    If PlusSpace = True Then
                        TempAns = TempAns & "+"
                    Else
                        TempAns = TempAns & "%" & Hex(32)
                    End If
                Case Else
                    TempAns = TempAns & "%" & Strings.Right("0" & Hex(Asc(Strings.Mid(StringToEncode, CurChr, 1))), 2)
            End Select

            CurChr = CurChr + 1
        Loop

        URLEncode = TempAns
    End Function

    'my version 
    Function getXCVer() As String
        getXCVer = "0.9" 'major minor versions
    End Function

    Function getXCbuildVer() As String
        getXCbuildVer = "0" ' build version 
    End Function

    'utc functions
    Function getTimeStamp() As String
        getTimeStamp = Format(Now, "yyMMdd-HHmmss")
    End Function

    Function expandTimeStamp(ByVal thisTimeStamp As String) As String ' ver 5.0
        Dim Months() As String, thisH As Integer, AMPM As String
        Months = Split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ",")

        thisH = CInt(Strings.Mid(thisTimeStamp, 8, 2))
        If thisH > 12 Then
            AMPM = " PM"
            thisH = thisH - 12
        Else
            AMPM = " AM"
        End If

        expandTimeStamp = thisH & ":" & Strings.Mid(thisTimeStamp, 10, 2) & ":" & Strings.Mid(thisTimeStamp, 12, 2) & AMPM & " - " & Months(CInt(Strings.Mid(thisTimeStamp, 3, 2)) - 1) & " " & Strings.Mid(thisTimeStamp, 5, 2) & ", 20" & Strings.Left(thisTimeStamp, 2)
    End Function

    Function buildDate_hex() As String
        buildDate_hex = DateTime.Now.Year Mod 1000
        buildDate_hex = xlWorkFunc.Dec2Bin(buildDate_hex, 7)
        buildDate_hex = buildDate_hex & xlWorkFunc.Dec2Bin(DateTime.Now.Month, 4) & xlWorkFunc.Dec2Bin(DateTime.Now.Day, 5)
        buildDate_hex = xlWorkFunc.Bin2Hex(Strings.Left(buildDate_hex, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Right(buildDate_hex, 8), 2)
    End Function

    Function Give_Time_1() As String
        'Dim YY As Integer, MM As Integer, DD As Integer, h As Integer, m As Integer, s As Integer

        Dim YY As String = DateTime.Now.ToString("yy")
        Dim MM As String = DateTime.Now.ToString("MM")
        Dim DD As String = DateTime.Now.ToString("dd")
        Dim h As String = DateTime.Now.ToString("HH")
        Dim m As String = DateTime.Now.ToString("mm")
        Dim s As String = DateTime.Now.ToString("ss")

        'YY = Year(Now)
        'MM = Month(Now)
        'DD = Day(Now)
        'h = Hour(Format(Now, "MM/dd/yyyy HH:mm:ss"))
        'm = Minute(Now)
        's = Second(Now)
        ' sec 60(64)6 : min 60(64)6 : hr 24(32)5 : day 31(32)5 : month 12(16)4 : year 64(64)6
        Give_Time_1 = xlWorkFunc.Dec2Bin(Strings.Right(YY, 2), 6) & xlWorkFunc.Dec2Bin(MM, 4) & xlWorkFunc.Dec2Bin(DD, 5) & xlWorkFunc.Dec2Bin(h, 5) & xlWorkFunc.Dec2Bin(m, 6) & xlWorkFunc.Dec2Bin(s, 6)
        Give_Time_1 = xlWorkFunc.Bin2Hex(Strings.Right(Give_Time_1, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Mid(Give_Time_1, 17, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Mid(Give_Time_1, 9, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Left(Give_Time_1, 8), 2)
    End Function

   

    'error and warn apend
    Function addErrorOrWarn(ByVal OriginalString As String, ByVal thisErrorOrWarn As String) As String
        If Strings.InStr(OriginalString & " ", thisErrorOrWarn & " ") = 0 Then
            addErrorOrWarn = OriginalString & vbNewLine & thisErrorOrWarn
        Else
            addErrorOrWarn = OriginalString
        End If
    End Function

    'file functions
    Function isFileExist(ByVal Fpath As String) As Boolean
        If System.IO.File.Exists(Fpath) Then
            isFileExist = True
        Else
            isFileExist = False
        End If
    End Function

    Function IsFileInUse(ByVal FileName As String) As Boolean
        Try
            Using f As New IO.FileStream(FileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            End Using
        Catch Ex As Exception
            Return True
        End Try
        Return False
    End Function

    Function isAnExcelFile(ByVal FileName As String) As Boolean
        If Strings.Right(Strings.LCase(FileName), 5) = ".xlsx" Then
            isAnExcelFile = True
        Else
            isAnExcelFile = False
        End If
    End Function

    'is IP
    'ip checker
    Function isAValidIP(ByVal inputip As String) As Boolean
        isAValidIP = False
        If Strings.InStr(inputip, ".") > 0 Then

            Dim LArray_ip() As String, L As Integer
            LArray_ip = Strings.Split(inputip, ".")

            If UBound(LArray_ip) = 3 Then

                For L = LBound(LArray_ip) To UBound(LArray_ip)
                    If IsNumeric(LArray_ip(L)) Then
                        If CInt(LArray_ip(L)) < 0 Or CInt(LArray_ip(L)) > 255 Then
                            Exit Function
                        End If
                    Else
                        Exit Function 'if atleast one non numeric found exit function with false
                    End If
                Next L

                isAValidIP = True
            End If
        End If
    End Function

    'name version
    Function getsufixname(ByVal MemmapName As String) As String
        getsufixname = Strings.Replace(Strings.Replace(Strings.Replace(MemmapName, "RCU mem map.", ""), ".xlsx", ""), ".XLSX", "")
    End Function

    'XCODE formatting functions
    Sub ClearProgFUN(ByRef thisProgramSheet As Excel.Worksheet, ByRef thisFunctionSheet As Excel.Worksheet)
        thisProgramSheet.Range("A:J").ClearContents()
        thisProgramSheet.Range("D:D").NumberFormat = "@"

        thisFunctionSheet.Range("A:J").ClearContents()
        thisFunctionSheet.Range("F:F").NumberFormat = "@"
    End Sub

    Sub ClearFunctionSeparaters(ByRef thisWorksheet As Excel.Worksheet)
        Dim i As Integer
        If Strings.Len(thisWorksheet.Range("E4").Text) = 0 Then
            For i = 0 To 5
                thisWorksheet.Range("E4").Offset(0, i).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            Next i
        ElseIf Strings.Len(thisWorksheet.Range("E4").Offset(1, 0).Text) = 0 Then
            For i = 0 To 5
                thisWorksheet.Range("E4").Offset(0, i).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                thisWorksheet.Range("E5").Offset(0, i).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            Next i
        Else
            Dim ce As Excel.Range
            For Each ce In thisWorksheet.Range(thisWorksheet.Range("E4"), thisWorksheet.Range("E4").End(Excel.XlDirection.xlDown))
                For i = 0 To 5
                    ce.Offset(0, i).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                Next i
            Next ce
        End If
    End Sub

    Sub setFunctionSeparaters(ByRef thisProgramSheet As Excel.Worksheet)
        Dim ce As Excel.Range, i As Integer
        For Each ce In thisProgramSheet.Range(thisProgramSheet.Range("E5"), thisProgramSheet.Range("E5").End(Excel.XlDirection.xlDown))
            If ce.Text = "04 01 00 00 00 00 00 00;" Or (Strings.Left(ce.Text, 5) = "01 01" And Strings.Mid(ce.Text, 7, 5) <> "00 00") Then
                For i = 0 To 5
                    With ce.Offset(0, i).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                        .LineStyle = Excel.XlLineStyle.xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = Excel.XlBorderWeight.xlMedium
                    End With
                Next i
            End If
        Next ce
    End Sub

    'XCODE modifiers
    Sub print_XCODE_build(ByRef thisXCODESheet As Excel.Worksheet)
        Dim XCODEConfig As String, XConfigArr() As String
        XCODEConfig = Strings.LCase(Strings.Trim(thisXCODESheet.Range("C1").Text))
        If Strings.InStr(XCODEConfig, "]") > 0 And Strings.InStr(XCODEConfig, "[") > 0 Then
            XCODEConfig = Strings.Mid(XCODEConfig, Strings.InStr(XCODEConfig, "[") + 1, Strings.InStr(XCODEConfig, "]") - Strings.InStr(XCODEConfig, "[") - 1)
            XConfigArr = Strings.Split(XCODEConfig, ",")
            If UBound(XConfigArr) > 0 Then
                thisXCODESheet.Range("C1").FormulaR1C1 = "XCODE [" & getXCVer() & "," & XConfigArr(1) & "]"
            End If
        End If
    End Sub

    Sub Check_XCODE_format(ByVal MemmapName As String, ByRef thisXCODESheet As Excel.Worksheet, ByVal thisRCU As String)
        Dim L As Integer, m As Integer, L1 As Integer, L2 As Integer, L3 As Integer, L4 As Integer, thisLine As String, V As Integer
        L = 0 : L1 = 1 : L2 = 1 : L3 = 1 : L4 = 1 : m = 1 : V = 0 'get first 2 lines

        While (L < 2 And m < 100)
            If Strings.Len(thisXCODESheet.Range("C" & m + 2).Text) > 2 Then
                thisLine = Strings.Trim(Strings.LCase(thisXCODESheet.Range("C" & m + 2).Text))
                If Strings.InStr(thisLine, "//") <> 1 Then
                    If Strings.InStr(thisLine, "//") > 4 Then
                        thisLine = Strings.Trim(Strings.Left(thisLine, Strings.InStr(thisLine, "//") - 1))
                    End If

                    If thisLine = "debug" Then
                        L = L + 1
                        L1 = m + 2
                    ElseIf thisLine = "exit" And L1 > 1 Then
                        L = L + 1
                        L2 = m + 2
                    ElseIf Strings.InStr(thisLine, "version") = 1 Then
                        L = L + 1
                        L3 = m + 2
                    ElseIf thisLine = "goto #debugging" And L3 > 1 Then
                        L = L + 1
                        L4 = m + 2
                    End If
                    V = V + 1
                End If
            End If
            m = m + 1
        End While


        If V = 2 And L1 > 1 And L2 > L1 Then
            thisXCODESheet.Range("C" & L1).FormulaR1C1 = "version [" & encript(getsufixname(MemmapName)) & ",-1]"
            thisXCODESheet.Range("C" & L2).ClearContents()
            thisXCODESheet.Range("C" & L1 + 1).FormulaR1C1 = "goto #debugging"
            ConsoleMsg("XC Progress:> Formatting XCODE" & thisRCU)
            move_add_debugging(thisXCODESheet)
        ElseIf V = 2 And L3 > 1 And L4 > L3 Then
            move_add_debugging(thisXCODESheet)
            Check_Version_Format(MemmapName, thisXCODESheet, L3)
        End If
    End Sub

    Private Sub move_add_debugging(ByRef thisXCODESheet As Excel.Worksheet)
        Dim ce As Excel.Range, This_Line As String, debugging As Boolean
        debugging = True
        For Each ce In thisXCODESheet.Range(thisXCODESheet.Range("B2"), thisXCODESheet.Range("B2").End(Excel.XlDirection.xlDown))
            If ce.Offset(0, 1).Text <> "" Then
                This_Line = Strings.LCase(Strings.Trim(ce.Offset(0, 1).Text))
                If Strings.InStr(This_Line, "function debugging") = 1 Then
                    debugging = False
                    Exit For
                End If
            End If
        Next ce

        If debugging Then
            Dim k As Integer, L1 As Integer, L2 As Integer
            k = 27 : L1 = 1 : L2 = 1

            While k < 200 And L2 = 1
                This_Line = Strings.LCase(Strings.Trim(thisXCODESheet.Range("C" & k).Text)) & " "
                If Strings.InStr(This_Line, "exit ") = 1 Or Strings.InStr(This_Line, "goto ") = 1 Then
                    L1 = k
                ElseIf Strings.InStr(This_Line, "function ") = 1 And L1 > 1 Then
                    L2 = k
                End If
                k = k + 1
            End While

            thisXCODESheet.Select()
            thisXCODESheet.Range("C" & L1).Select()

            push_code_down(thisXCODESheet, "")
            push_code_down(thisXCODESheet, "")
            push_code_down(thisXCODESheet, "function debugging")
            push_code_down(thisXCODESheet, "debug")
            push_code_down(thisXCODESheet, "exit")
            push_code_down(thisXCODESheet, "")
            push_code_down(thisXCODESheet, "")
        End If
    End Sub

    Private Sub Check_Version_Format(ByVal MemmapName As String, ByRef thisXCODESheet As Excel.Worksheet, vLine As Integer)
        Dim thisLine As String, vData() As String
        thisLine = Strings.Trim(Strings.LCase(thisXCODESheet.Range("C" & vLine).Text))

        If Strings.InStr(thisLine, "[") < Strings.InStr(thisLine, "]") Then
            vData = Strings.Split(Strings.Mid(thisLine, Strings.InStr(thisLine, "[") + 1, Strings.InStr(thisLine, "]") - Strings.InStr(thisLine, "[") - 1), ",")
            If vData(0) <> encript(getsufixname(MemmapName)) Then
                thisXCODESheet.Range("C" & vLine).FormulaR1C1 = "version [" & encript(getsufixname(MemmapName)) & ",-1]"
            End If
        Else
            thisXCODESheet.Range("C" & vLine).FormulaR1C1 = "version [" & encript(getsufixname(MemmapName)) & ",-1]"
        End If

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

    'xcode core 1
    ' rcu var
    Function hasValidVar(ByVal inputString As String) As Boolean 'RCU var
        hasValidVar = False
        If Strings.InStr(Strings.LCase(inputString), " $var") > 1 Then
            Dim RcuVariableList() As String, RcuNameList() As String
            RcuVariableList = Strings.Split("04,05,06,07,08,09,0A,0B,0C,0D,0E,0F,10,11,12,13,14,14,33,15,16,17,18,19,1A,1B,1C,1D,1E,1F,20,21,22,23,24,25,26,27", ",")
            RcuNameList = Strings.Split("anti_ice,checkin,thermo1,thermo2,thermo3,thermo4,thermo5,thermo6,temp1,temp2,temp3,temp4,temp5,temp6,season,intervention,room_empty,room_empty_temp,room_empty_thermo,user,workflow,gs1,gs2,gs3,gs4,gs5,gs6,gs7,gs8,gs9,gs10,ioexp1,ioexp2,ioexp3,ioexp4,ioexp5,ioexp6,tag_minmax", ",")
            Dim LArray_T2() As String, V As String
            V = ""
            LArray_T2 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If Strings.InStr(Strings.Replace(LArray_T2(i), " ", ""), "var=") Then
                    V = Strings.Replace(LArray_T2(i), "var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
                For i = LBound(RcuVariableList) To UBound(RcuVariableList)
                    If V = RcuVariableList(i) Or V = RcuNameList(i) Then
                        hasValidVar = True
                        Exit For
                    End If
                Next i
            End If
        End If
    End Function

    Function Grab_Var_1(ByVal inputString As String) As String 'RCU var
        Dim V As String, LArray_T1() As String, P As String, i As Integer
        Dim RcuVariableList() As String, RcuNameList() As String
        RcuVariableList = Strings.Split("04,05,06,07,08,09,0A,0B,0C,0D,0E,0F,10,11,12,13,14,14,33,15,16,17,18,19,1A,1B,1C,1D,1E,1F,20,21,22,23,24,25,26,27", ",")
        RcuNameList = Strings.Split("anti_ice,checkin,thermo1,thermo2,thermo3,thermo4,thermo5,thermo6,temp1,temp2,temp3,temp4,temp5,temp6,season,intervention,room_empty,room_empty_temp,room_empty_thermo,user,workflow,gs1,gs2,gs3,gs4,gs5,gs6,gs7,gs8,gs9,gs10,ioexp1,ioexp2,ioexp3,ioexp4,ioexp5,ioexp6,tag_minmax", ",")
        V = "" : P = ""
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "var=") = 1 Then
                V = Strings.Replace(LArray_T1(i), "var=", "")
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

    ' is function 
    Function isAFunction(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "function ") = 1 Then
            isAFunction = True
        Else
            isAFunction = False
        End If
    End Function

    Function getLineFromFunctionName(ByVal inputString As String, ByVal inputLine As Integer, ByRef thisWorkbook As Excel.Workbook, ByVal thisRCU As String) As String
        Dim thisXlProgramSheet As Excel.Worksheet
        Dim thisXlFUNCTIONSheet As Excel.Worksheet
        Dim ce As Excel.Range
        getLineFromFunctionName = "0000"

        If inputString = "exit" Or inputString = "quit" Or inputString = "end" Then

            thisXlProgramSheet = thisWorkbook.Sheets("Program" & thisRCU)

            If Strings.Len(thisXlProgramSheet.Range("F" & (inputLine + 1)).Text) = 0 Then
                'to remove this condition from else
            ElseIf Strings.Len(thisXlProgramSheet.Range("F" & (inputLine + 2)).Text) = 0 Then
                If thisXlProgramSheet.Range("F" & (inputLine + 1)).Text = inputString Then
                    getLineFromFunctionName = xlWorkFunc.Dec2Hex(inputLine - 3, 4)
                End If
            Else
                For Each ce In thisXlProgramSheet.Range(thisXlProgramSheet.Range("F" & (inputLine + 1)), thisXlProgramSheet.Range("F" & (inputLine + 1)).End(Excel.XlDirection.xlDown))
                    If ce.Text = inputString Then
                        getLineFromFunctionName = xlWorkFunc.Dec2Hex(ce.Row - 3, 4)
                        Exit Function
                    End If
                Next ce
            End If
        ElseIf IsNumeric(inputString) Then
            If (inputLine + CInt(inputString) - 3) > 0 Then
                getLineFromFunctionName = xlWorkFunc.Dec2Hex(inputLine + CInt(inputString) - 3, 4)
            End If
        Else

            thisXlFUNCTIONSheet = thisWorkbook.Sheets("FUNCTION" & thisRCU)

            If Strings.Len(thisXlFUNCTIONSheet.Range("E3").Text) = 0 Then
                'to remove this condition from else
            ElseIf Strings.Len(thisXlFUNCTIONSheet.Range("E4").Text) = 0 Then
                If thisXlFUNCTIONSheet.Range("E3").Text = inputString Then
                    getLineFromFunctionName = thisXlFUNCTIONSheet.Range("F3").Text
                End If
            Else
                For Each ce In thisXlFUNCTIONSheet.Range(thisXlFUNCTIONSheet.Range("E3"), thisXlFUNCTIONSheet.Range("E3").End(Excel.XlDirection.xlDown))
                    If ce.Text = inputString Then
                        getLineFromFunctionName = ce.Offset(0, 1).Text
                        Exit Function
                    End If
                Next ce
            End If
        End If
    End Function


    'is other rcu header functions
    Function isOtherRCUcall(ByVal inputString As String) As Boolean
        If (Strings.InStr(inputString, "run ") = 1 Or Strings.InStr(inputString, "reuse ") = 1) And Strings.InStr(inputString, " @rcu") > 1 Then
            isOtherRCUcall = True
        Else
            isOtherRCUcall = False
        End If
    End Function

    ' is ioexp direct
    Function isAnIOEXPDirect(ByVal inputString As String) As Boolean
        inputString = Strings.Replace(inputString, " ", "")
        If Strings.InStr(inputString, "ioexp") = 1 And (Strings.InStr(inputString, ".") = 6 Or Strings.InStr(inputString, ".") = 7) Then
            isAnIOEXPDirect = True
        Else
            isAnIOEXPDirect = False
        End If
    End Function

    ' is iodexp direct
    Function isAnIODEXPDirect(ByVal inputString As String) As Boolean
        inputString = Strings.Replace(inputString, " ", "")
        If Strings.InStr(inputString, "iodexp") = 1 And (Strings.InStr(inputString, ".") = 7 Or Strings.InStr(inputString, ".") = 8) Then
            isAnIODEXPDirect = True
        Else
            isAnIODEXPDirect = False
        End If
    End Function

    ' is rcu direct
    Function isARCUDirect(ByVal inputString As String) As Boolean
        inputString = Strings.Replace(inputString, " ", "")
        If Strings.InStr(inputString, "rcu.") = 1 And _
        (Strings.InStr(inputString, ".set$relay=") > 0 Or Strings.InStr(inputString, ".or$relay=") > 0 Or _
        Strings.InStr(inputString, ".unset$relay=") > 0 Or Strings.InStr(inputString, ".and$relay=") > 0 Or _
        Strings.InStr(inputString, ".xor$relay=") > 0) Then
            isARCUDirect = True
        Else
            isARCUDirect = False
        End If
    End Function

    ' is math direct
    Function isAMathDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "math.") = 1 Then
            isAMathDirect = True
        Else
            isAMathDirect = False
        End If
    End Function

    ' is save direct
    Function isASaveDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "save") = 1 Then
            isASaveDirect = True
        Else
            isASaveDirect = False
        End If
    End Function

    'extend modbus direct
    Function isAModbusDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "send.mod") = 1 Then
            isAModbusDirect = True
        Else
            isAModbusDirect = False
        End If
    End Function

    'expansion dimmer direct
    Function isADimmerDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "send.dim") = 1 Then
            isADimmerDirect = True
        Else
            isADimmerDirect = False
        End If
    End Function

    ' dimmer read direct
    Function isADimmerReadDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "read.dim") = 1 Then
            isADimmerReadDirect = True
        Else
            isADimmerReadDirect = False
        End If
    End Function

    ' dimmer load direct 
    Function isADimmerLoad(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "load dimmer ") = 1 Then
            isADimmerLoad = True
        Else
            isADimmerLoad = False
        End If
    End Function

    Function LoadDimmerConvert(ByVal inputString As String) As String
        LoadDimmerConvert = "10 80 01 01 " & Grab_value_2(inputString) & " " & Grab_Time_2(inputString) & ";"
    End Function

    'dali direct
    Function isADaliDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "dali") = 1 Then
            isADaliDirect = True
        Else
            isADaliDirect = False
        End If
    End Function

    'dmx
    Function isADMXDirect(ByVal inputString As String) As Boolean
        If Strings.InStr(Strings.Replace(inputString, " ", ""), "dmx.") = 1 Then
            isADMXDirect = True
        Else
            isADMXDirect = False
        End If
    End Function

    'definition header
    Function isADefinition(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "define ") = 1 Then
            isADefinition = True
        Else
            isADefinition = False
        End If
    End Function

    



    'xcode core - 2

    ' has function call
    Function hasAFunction(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, " #") > 0 Then
            hasAFunction = True
        Else
            hasAFunction = False
        End If
    End Function

    Function isARCU(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "rcu.") = 1 Then
            isARCU = True
        Else
            isARCU = False
        End If
    End Function

    'dmx
    Function isADMX(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "dmx.") = 1 Then
            isADMX = True
        Else
            isADMX = False
        End If
    End Function

    ' abs wait
    Function isWaitABS(ByVal inputString As String) As Boolean   'XCODE 5.1 wait abs instruction check
        If Strings.InStr(inputString, "wait abs ") = 1 Then
            isWaitABS = True
        Else
            isWaitABS = False
        End If
    End Function

    Function WaitConvert(ByVal inputString As String) As String
        WaitConvert = "01 01 00 00 " & Grab_Time_1(inputString) & ";"
    End Function

    ' is an exit
    'exit handling
    Function isAnExit(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "exit") = 1 Or Strings.InStr(inputString, "quit") = 1 Or Strings.InStr(inputString, "end") = 1 Then
            isAnExit = True
        Else
            isAnExit = False
        End If
    End Function

    Function ExitConvert() As String
        ExitConvert = "04 01 00 00 00 00 00 00;"
    End Function

    'debug handling
    Function isDebug(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "debug") = 1 Then
            isDebug = True
        Else
            isDebug = False
        End If
    End Function

    Function DebugConvert() As String
        DebugConvert = "40 01 01 00 00 00 00 00;"
    End Function

    'version handling
    Function isVersion(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "version") = 1 Then
            isVersion = True
        Else
            isVersion = False
        End If
    End Function

    Function VersionConvert(ByVal inputString As String) As String
        Dim buildVersion As Integer, vData() As String
        vData = Strings.Split(Strings.Mid(inputString, Strings.InStr(inputString, "[") + 1, Strings.InStr(inputString, "]") - Strings.InStr(inputString, "[") - 1), ",")
        buildVersion = CInt(vData(1))
        VersionConvert = "00 00 00 00 " & xlWorkFunc.Dec2Hex(CDbl(getXCVer()) * 10, 2) & " " & buildDate_hex() & " " & xlWorkFunc.Dec2Hex(buildVersion, 2) & ";"
    End Function

    ' load handling
    Function isALoad(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "load ") = 1 Then
            isALoad = True
        Else
            isALoad = False
        End If
    End Function

    ' is send
    Function isASend(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "send.") = 1 Then
            isASend = True
        Else
            isASend = False
        End If
    End Function

    'read handling
    Function isARead(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "read.") = 1 Then
            isARead = True
        Else
            isARead = False
        End If
    End Function

    'save handling
    Function isASave(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "save ") = 1 Then
            isASave = True
        Else
            isASave = False
        End If
    End Function

    ' math handling
    Function isAMath(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "math.") = 1 Then
            isAMath = True
        Else
            isAMath = False
        End If
    End Function

    'dimflow handling
    Function isADimFlow(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, "dimflow.") = 1 Then
            isADimFlow = True
        Else
            isADimFlow = False
        End If
    End Function

    ' grabbers
    ' end line worker functions
    Function GrabHexString(ByVal inputString As String) As String
        Dim HTL1 As String
        HTL1 = Strings.Left(Strings.Replace(Strings.Trim(Strings.Replace(inputString, "0x", "")), " ", ""), 16)
        If isValidHexString(HTL1) Then
            GrabHexString = Strings.Left(HTL1, 2) & " " & Strings.Mid(HTL1, 3, 2) & " " & Strings.Mid(HTL1, 5, 2) & " " & Strings.Mid(HTL1, 7, 2) & " " & Strings.Mid(HTL1, 9, 2) & " " & Strings.Mid(HTL1, 11, 2) & " " & Strings.Mid(HTL1, 13, 2) & " " & Strings.Right(HTL1, 2) & ";"
        Else
            GrabHexString = "FALSE"
        End If
    End Function

    Function GrabComment(ByVal inputString As String) As String
        If Strings.InStr(inputString, "//") Then
            Dim GC1 As String() = Strings.Split(inputString, "//")
            GrabComment = "//" & GC1(1)
        Else
            GrabComment = "//Direct hex"
        End If
    End Function

    Function GrabFunctionName(ByVal inputString As String) As String
        Dim LArray() As String
        LArray = Strings.Split(inputString, " ")
        GrabFunctionName = LArray(1)
    End Function

    'value grabbers
    'value for OR, relay wise

    Function Grab_value_0(ByVal inputString As String, index As Integer) As String 'return indexed value otherwise null
        Dim LArray_v1() As String, tempS As String, indexcount As Integer, i As Integer
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        indexcount = 0
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            tempS = Strings.Replace(LArray_v1(i), " ", "")
            If Strings.InStr(tempS, "bin=") = 1 Or Strings.InStr(tempS, "dec=") = 1 Or Strings.InStr(tempS, "hex=") = 1 _
            Or Strings.InStr(tempS, "relay=") = 1 Or Strings.InStr(tempS, "bit=") = 1 Or Strings.InStr(tempS, "bin=") = 1 _
            Or Strings.InStr(tempS, "fan=") = 1 Or Strings.InStr(tempS, "display=") = 1 Or Strings.InStr(tempS, "var=") Then
                Grab_value_0 = tempS
                indexcount = indexcount + 1
                If indexcount = index Then
                    Exit Function
                End If
            End If
        Next i
        Grab_value_0 = "NULL"
    End Function

    Function Grab_value_1(ByVal inputString As String, ByVal Vmode As String) As String
        Dim P As String, Q As Long, R As String, t As Integer, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = "" : t = 0
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bin=") = 1 Then
                P = GrabPara_BIN(LArray_v1(i))
                R = xlWorkFunc.Bin2Hex(Strings.Left(P, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 And (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "relay")
                R = xlWorkFunc.Bin2Hex(Strings.Left(P, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 And (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "relay")
                R = xlWorkFunc.Bin2Hex(Strings.Left(P, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 And (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "bit")
                R = xlWorkFunc.Bin2Hex(Strings.Left(P, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 And (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "bit")
                R = xlWorkFunc.Bin2Hex(Strings.Left(P, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Lng(LArray_v1(i), "dec")
                R = xlWorkFunc.Dec2Hex(Q, 8)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "hex=") = 1 Then
                R = GrabPara_HEX(LArray_v1(i))
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "fan=") = 1 Then
                R = GrabPara_FAN(LArray_v1(i))
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "display=") = 1 Then
                Grab_value_1 = GrabPara_Display(LArray_v1(i))
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "len=") = 1 Then
                t = GrabPara_Int(LArray_v1(i), "len")
                R = xlWorkFunc.Dec2Hex(t, 8)
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "clock=") = 1 Then
                R = GrabPara_Clock(LArray_v1(i))
                Grab_value_1 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
                Exit Function
            End If
        Next i
        Grab_value_1 = "00 00 00 00"
    End Function

    Private Function Grab_value_2(inputString As String) As String
        Dim P As String, Q As Double, R As String, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = ""
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = xlWorkFunc.Dec2Hex(P * 100, 8)
                Grab_value_2 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = xlWorkFunc.Dec2Hex(Q * 100, 8)
                Grab_value_2 = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2)
                Exit Function
            End If
        Next i
        Grab_value_2 = "00 00"
    End Function

    Function Grab_value_3(inputString As String, Vmode As String) As String
        Dim P As String, Q As Long, R As String, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = ""
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bin=") = 1 Then
                P = GrabPara_BIN(LArray_v1(i))
                R = xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_3 = Strings.Left(R, 2) & " " & Strings.Mid(R, 3, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 And (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "relay")
                R = xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_3 = Strings.Left(R, 2) & " " & Strings.Mid(R, 3, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 And (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "relay")
                R = xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_3 = Strings.Left(R, 2) & " " & Strings.Mid(R, 3, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 And (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "bit")
                R = xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_3 = Strings.Left(R, 2) & " " & Strings.Mid(R, 3, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 And (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "bit")
                R = xlWorkFunc.Bin2Hex(Strings.Mid(P, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(P, 8), 2)
                Grab_value_3 = Strings.Left(R, 2) & " " & Strings.Mid(R, 3, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Lng(LArray_v1(i), "dec")
                R = xlWorkFunc.Dec2Hex(Q, 4)
                Grab_value_3 = Strings.Left(R, 2) & " " & Strings.Mid(R, 3, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "hex=") = 1 Then
                R = GrabPara_HEX(LArray_v1(i))
                Grab_value_3 = Strings.Mid(R, 5, 2) & " " & Strings.Right(R, 2)
                Exit Function
            End If
        Next i
        Grab_value_3 = "00 00"
    End Function

    Function Grab_value_4(inputString As String) As String
        Dim P As String, Q As Double, R As String, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = ""
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = xlWorkFunc.Dec2Hex(P, 8)
                Grab_value_4 = Strings.Right(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = xlWorkFunc.Dec2Hex(Q, 8)
                Grab_value_4 = Strings.Right(R, 2)
                Exit Function
            End If
        Next i
        Grab_value_4 = "00"
    End Function

    Function Grab_value_5(inputString As String) As String 'for dali
        Dim P As Double, R As Double, i As Integer
        Dim LArray_v1() As String
        P = 0
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = P * 254 / 100
                Grab_value_5 = xlWorkFunc.Dec2Hex(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = P * 254 / 100
                Grab_value_5 = xlWorkFunc.Dec2Hex(R, 2)
                Exit Function
            End If
        Next i
        Grab_value_5 = "00"
    End Function

    Function Grab_value_6(inputString As String) As String 'for dmx
        Dim P As Double, R As Double, i As Integer
        Dim LArray_v1() As String
        P = 0
        LArray_v1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = P * 255 / 100
                Grab_value_6 = xlWorkFunc.Dec2Hex(R, 2)
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = P * 255 / 100
                Grab_value_6 = xlWorkFunc.Dec2Hex(R, 2)
                Exit Function
            End If
        Next i
        Grab_value_6 = "00"
    End Function

    'variant
    Function Grab_variant(ByVal inputString As String, ByVal dataType As String) As String
        Dim LArray_v1() As String, i As Integer
        Grab_variant = "NULL"
        LArray_v1 = Strings.Split(Strings.Replace(Strings.Replace(Strings.LCase(inputString), "#", "$"), " ", ""), "$")

        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(LArray_v1(i), dataType & "=") = 1 Then
                Grab_variant = Strings.Replace(LArray_v1(i), dataType & "=", "")
                If Strings.Len(Grab_variant) < 1 Then
                    Grab_variant = "NULL"
                End If
                Exit Function
            End If
        Next
    End Function

    'grab ip
    Function Grab_IP_addr(ByVal inputString As String) As String
        Dim LArray_ip() As String, R As String, k As Integer

        LArray_ip = Strings.Split(Strings.Replace(inputString, " ", ""), ".")
        R = ""

        For k = LBound(LArray_ip) To UBound(LArray_ip)
            R = R & xlWorkFunc.Dec2Hex(LArray_ip(k), 2)
        Next k

        Grab_IP_addr = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
    End Function

    ' grab var
    ' settings
    'added in v5.3 read write settings
    Function Grab_Var_4(ByVal inputString As String) As String 'settings
        Dim V As String, LArray_T1() As String, i As Integer

        V = "00000000"
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "var=") Then
                V = Strings.Replace(LArray_T1(i), "var=", "")
                Exit For
            End If
        Next i

        If isValidHexString(V) Then
            If xlWorkFunc.Hex2Dec(V) < 65536 Then
                V = Strings.Right("00000000" & V, 8)
            End If
        End If

        Grab_Var_4 = Strings.Right(V, 2) & " " & Strings.Mid(V, 5, 2) & " " & Strings.Mid(V, 3, 2) & " " & Strings.Left(V, 2)
    End Function

    ' grab elements
    Function Grab_element(ByVal inputString As String) As String
        Dim element As String
        element = Grab_variant(inputString, "element")
        If element = "NULL" Then
            Grab_element = "00 00 00 00"
        Else
            element = Strings.Right("00000000" & element, 8)
            Grab_element = Strings.Right(element, 2) & " " & Strings.Mid(element, 5, 2) & " " & Strings.Mid(element, 3, 2) & " " & Strings.Left(element, 2)
        End If
    End Function

    'grab register
    Function Grab_Reg_1(ByVal inputString As String) As String
        Dim P As Integer, i As Integer
        Dim LArray_T1() As String
        Dim LArray_GP() As String

        P = 0
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "reg=") Then
                LArray_GP = Strings.Split(Strings.Replace(LArray_T1(i), " ", ""), "=")
                If IsNumeric(LArray_GP(1)) And LArray_GP(0) = "reg" Then
                    If CInt(LArray_GP(1)) > -1 And CInt(LArray_GP(1)) < 65536 Then
                        P = CInt(LArray_GP(1))
                        Exit For
                    End If
                End If
            End If
        Next i

        Grab_Reg_1 = Strings.Left(xlWorkFunc.Dec2Hex(P, 4), 2) & " " & Strings.Right(xlWorkFunc.Dec2Hex(P, 4), 2)
    End Function

    ' grab dev type
    Function Grab_Dev_Type_1(ByVal inputString As String) As String
        Dim P As Integer, i As Integer
        Dim LArray_T1() As String
        Dim LArray_GP() As String

        P = 0
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T1) To UBound(LArray_T1)

            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "devtype=") Or Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "type=") Then
                LArray_GP = Strings.Split(Strings.Replace(LArray_T1(i), " ", ""), "=")
                If IsNumeric(LArray_GP(1)) And (LArray_GP(0) = "devtype" Or LArray_GP(0) = "type") Then
                    If CInt(LArray_GP(1)) > 0 And CInt(LArray_GP(1)) < 11 Then
                        P = CInt(LArray_GP(1))
                        Exit For
                    End If
                ElseIf (LArray_GP(1) = "server" Or LArray_GP(1) = "db" Or LArray_GP(1) = "ff") And (LArray_GP(0) = "devtype" Or LArray_GP(0) = "type") Then
                    P = 255
                    Exit For
                ElseIf (Not IsNumeric(LArray_GP(1))) And (LArray_GP(0) = "devtype" Or LArray_GP(0) = "type") Then
                    P = Grab_Type_from_Name(LArray_GP(1))
                    Exit For
                End If
            End If
        Next i

        Grab_Dev_Type_1 = xlWorkFunc.Dec2Hex(P, 2)
    End Function

    Private Function Grab_Type_from_Name(ByVal inputString As String) As Integer
        Dim RcuDevList() As String, RcuDevNameList() As String

        RcuDevList = Strings.Split("01,02,03,04,04,05,05,06,06,07,07,08,10", ",")
        RcuDevNameList = Strings.Split("idpg,tig,tag,ioexp,ioe,iodexp,iod,gs,gsw,ioexp,ioe,bsp,dali", ",")
        Grab_Type_from_Name = 0

        For i = LBound(RcuDevList) To UBound(RcuDevList)
            If Strings.InStr(inputString, RcuDevNameList(i)) Then
                Grab_Type_from_Name = CInt(RcuDevList(i))
                Exit Function
            End If
        Next i

    End Function


    'Grab_Time_1 > return time Ticks
    'Grab_Value_1 > retrun value input to functions
    'Grab_Dev_1 > return device for 03 03

    Function hasTime(ByVal inputString As String) As Boolean
        If Strings.InStr(inputString, " $hour") > 0 Or Strings.InStr(inputString, " $min") > 0 Or Strings.InStr(inputString, " $sec") > 0 Then
            hasTime = True
        Else
            hasTime = False
        End If
    End Function

    'time for run wait
    Function Grab_Time_1(ByVal inputString As String) As String
        Dim P As Double, Q As Double, R As Double, t As Double, s As String
        Dim LArray_T1() As String, i As Integer

        P = 0 : Q = 0 : R = 0
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "hour=") Then
                P = GrabPara_Dbl(LArray_T1(i), "hour")
            ElseIf Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "min=") Then
                Q = GrabPara_Dbl(LArray_T1(i), "min")
            ElseIf Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "sec=") Then
                R = GrabPara_Dbl(LArray_T1(i), "sec")
            End If
        Next i

        t = ((P * 3600) + (Q * 60) + R) * 10000
        s = xlWorkFunc.Dec2Hex(t, 8)

        Grab_Time_1 = Strings.Right(s, 2) & " " & Strings.Mid(s, 5, 2) & " " & Strings.Mid(s, 3, 2) & " " & Strings.Left(s, 2)
    End Function

    'time for dimmer load
    Private Function Grab_Time_2(ByVal inputString As String) As String
        Dim P As Double, Q As Double, R As Double, t As Double, s As String
        Dim LArray_T1() As String, i As Integer

        P = 0 : Q = 0 : R = 0
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "hour=") Then
                P = GrabPara_Dbl(LArray_T1(i), "hour")
            ElseIf Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "min=") Then
                Q = GrabPara_Dbl(LArray_T1(i), "min")
            ElseIf Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "sec=") Then
                R = GrabPara_Dbl(LArray_T1(i), "sec")
            End If
        Next i

        t = ((P * 3600) + (Q * 60) + R) * 100
        s = xlWorkFunc.Dec2Hex(t, 8)

        Grab_Time_2 = Strings.Right(s, 2) & " " & Strings.Mid(s, 5, 2)
    End Function

    'time for RCU dimmer
    Function Grab_Time_3(ByVal inputString As String) As String
        Dim P As Double, Q As Double, R As Double, t As Double, s As String
        Dim LArray_T1() As String, i As Integer

        P = 0 : Q = 0 : R = 0
        LArray_T1 = Strings.Split(Strings.Replace(inputString, "#", " $"), " $")

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "hour=") Then
                P = GrabPara_Dbl(LArray_T1(i), "hour")
            ElseIf Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "min=") Then
                Q = GrabPara_Dbl(LArray_T1(i), "min")
            ElseIf Strings.InStr(Strings.Replace(LArray_T1(i), " ", ""), "sec=") Then
                R = GrabPara_Dbl(LArray_T1(i), "sec")
            End If
        Next i

        t = ((P * 3600) + (Q * 60) + R) * 10000
        s = xlWorkFunc.Dec2Hex(t, 8)

        Grab_Time_3 = Strings.Right(s, 2) & " " & Strings.Mid(s, 5, 2) & " " & Strings.Mid(s, 3, 2)
    End Function

    ' level 0 grabbers
    Private Function GrabPara_Dbl(inputString As String, checkPara As String) As Double
        Dim LArray_GP() As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If IsNumeric(Evaluate(LArray_GP(1))) And LArray_GP(0) = checkPara Then
            If CDbl(Evaluate(LArray_GP(1))) > 0 And CDbl(Evaluate(LArray_GP(1))) < 4294967296.0# Then
                GrabPara_Dbl = CDbl(Evaluate(LArray_GP(1)))
            Else
                GrabPara_Dbl = 0
            End If
        Else
            GrabPara_Dbl = 0
        End If
    End Function

    Private Function GrabPara_BIN(inputString As String) As String
        Dim LArray_GP() As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If isValidBinString(LArray_GP(1)) And LArray_GP(0) = "bin" Then
            GrabPara_BIN = Strings.Right("00000000000000000000000000000000" & LArray_GP(1), 32)
        Else
            GrabPara_BIN = "00000000000000000000000000000000"
        End If
    End Function

    Private Function GrabPara_HEX(inputString As String) As String
        Dim LArray_GP() As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If isValidHexString(LArray_GP(1)) And LArray_GP(0) = "hex" Then
            GrabPara_HEX = Strings.Right("00000000" & UCase(LArray_GP(1)), 8)
        Else
            GrabPara_HEX = "00000000"
        End If
    End Function

    Private Function GrabPara_Lng(ByVal inputString As String, ByVal checkPara As String) As Long
        Dim LArray_GP() As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If IsNumeric(Evaluate(LArray_GP(1))) And LArray_GP(0) = checkPara Then
            If CLng(Evaluate(LArray_GP(1))) > 0 And CLng(Evaluate(LArray_GP(1))) < 4294967296.0# Then
                GrabPara_Lng = CLng(Evaluate(LArray_GP(1)))
            Else
                GrabPara_Lng = 0
            End If
        Else
            GrabPara_Lng = 0
        End If
    End Function

    Private Function GrabPara_Int(ByVal inputString As String, ByVal checkPara As String) As Integer
        Dim LArray_GP() As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If IsNumeric(Evaluate(LArray_GP(1))) And LArray_GP(0) = checkPara Then
            If CInt(Evaluate(LArray_GP(1))) > 0 And CInt(Evaluate(LArray_GP(1))) < 4294967296.0# Then
                GrabPara_Int = CInt(Evaluate(LArray_GP(1)))
            Else
                GrabPara_Int = 0
            End If
        Else
            GrabPara_Int = 0
        End If
    End Function

    Private Function GrabPara_BitPOS_OR(ByVal inputString As String, ByVal checkPara As String) As String
        Dim LArray_GP() As String, RelayBits() As String, i As Integer
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        RelayBits = Strings.Split("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
        GrabPara_BitPOS_OR = ""
        If LArray_GP(0) = checkPara Then
            Dim Sbits() As String
            LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
            Sbits = Strings.Split(LArray_GP(1), ",")

            For i = LBound(Sbits) To UBound(Sbits)
                If IsNumeric(Sbits(i)) Then
                    If CInt(Sbits(i)) > 0 And CInt(Sbits(i)) < 33 Then
                        RelayBits(CInt(Sbits(i)) - 1) = "1"
                    End If
                End If
            Next i
        End If
        For i = LBound(RelayBits) To UBound(RelayBits)
            GrabPara_BitPOS_OR = RelayBits(i) & GrabPara_BitPOS_OR
        Next i
    End Function

    Private Function GrabPara_BitPOS_AND(ByVal inputString As String, ByVal checkPara As String) As String
        Dim LArray_GP() As String, RelayBits() As String, i As Integer
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        RelayBits = Strings.Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1", ",")
        GrabPara_BitPOS_AND = ""
        If LArray_GP(0) = checkPara Then
            Dim Sbits() As String
            LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
            Sbits = Strings.Split(LArray_GP(1), ",")

            For i = LBound(Sbits) To UBound(Sbits)
                If IsNumeric(Sbits(i)) Then
                    If CInt(Sbits(i)) > 0 And CInt(Sbits(i)) < 33 Then
                        RelayBits(CInt(Sbits(i)) - 1) = "0"
                    End If
                End If
            Next i
        End If
        For i = LBound(RelayBits) To UBound(RelayBits)
            GrabPara_BitPOS_AND = RelayBits(i) & GrabPara_BitPOS_AND
        Next i
    End Function

    Private Function GrabPara_FAN(ByVal inputString As String) As String
        Dim LArray_GP() As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If LArray_GP(1) = "fan1" And LArray_GP(0) = "fan" Then
            GrabPara_FAN = "003F0655"
        ElseIf LArray_GP(1) = "fan2" And LArray_GP(0) = "fan" Then
            GrabPara_FAN = "003F0565"
        ElseIf LArray_GP(1) = "fan3" And LArray_GP(0) = "fan" Then
            GrabPara_FAN = "003F0556"
        ElseIf (LArray_GP(1) = "auto" Or LArray_GP(1) = "fan4") And LArray_GP(0) = "fan" Then
            GrabPara_FAN = "003F0955"
        ElseIf (LArray_GP(1) = "eco" Or LArray_GP(1) = "fan5") And LArray_GP(0) = "fan" Then
            GrabPara_FAN = "003F0595"
        ElseIf (LArray_GP(1) = "off" Or LArray_GP(1) = "fan6") And LArray_GP(0) = "fan" Then
            GrabPara_FAN = "003F0559"
        Else
            GrabPara_FAN = "00000000"
        End If
    End Function

    Private Function GrabPara_Display(ByVal inputString As String, Optional lengthByte As Integer = 4) As String
        Dim LArray_GP() As String, thischar As String, nextchar As String
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")

        If LArray_GP(0) = "display" Then

            Dim DisplayChar() As String, DisplayHexChar() As String, CharSpace() As String, HexCharSpace() As String, rPoint As Integer
            DisplayChar = Strings.Split("a,b,/c,\c,c,d,/e,e,f,g,/h,h,i,j,l,/n,n,/o,\o,o,p,q,r,s,t,/u,\u,u,y,0,1,2,3,4,5,6,7,8,9,.,-,_, ", ",")
            DisplayHexChar = Strings.Split("77,1F,0E,43,4D,3E,6F,4F,47,5D,17,37,10,38,0D,16,75,1E,63,7D,67,73,06,5B,0F,1C,23,3D,3B,7D,30,6E,7A,33,5B,5F,70,7F,7B,80,02,08,00", ",")
            CharSpace = Strings.Split(",,,", ",")
            HexCharSpace = Strings.Split("00,00,00,00", ",")

            If Strings.Left(LArray_GP(1), 1) = "[" Then
                LArray_GP(1) = Strings.Mid(LArray_GP(1), 2, Strings.Len(LArray_GP(1)))
            End If
            If Strings.Right(LArray_GP(1), 1) = "]" Then
                LArray_GP(1) = Strings.Mid(LArray_GP(1), 1, Strings.Len(LArray_GP(1)) - 1)
            End If
            LArray_GP(1) = "        " & LArray_GP(1)
            rPoint = -1

            For i = 0 To Strings.Len(LArray_GP(1)) - 4
                thischar = Strings.Mid(LArray_GP(1), Strings.Len(LArray_GP(1)) - i, 1)
                nextchar = Strings.Mid(LArray_GP(1), Strings.Len(LArray_GP(1)) - 1 - i, 1)

                If thischar <> "/" And thischar <> "\" Then
                    If thischar <> "." Then
                        rPoint = rPoint + 1
                        If rPoint < 4 Then
                            CharSpace(rPoint) = thischar & CharSpace(rPoint)
                        End If
                    Else
                        If thischar = "." And nextchar <> "." Then
                            If rPoint < 3 Then
                                CharSpace(rPoint + 1) = thischar & CharSpace(rPoint + 1)
                            End If
                        ElseIf thischar = "." And nextchar = "." Then
                            rPoint = rPoint + 1
                            If rPoint < 4 Then
                                CharSpace(rPoint) = thischar & CharSpace(rPoint)
                            End If
                        End If
                    End If
                Else
                    If rPoint > -1 And rPoint < 4 Then
                        CharSpace(rPoint) = thischar & CharSpace(rPoint)
                    End If
                End If

            Next i

            For i = 0 To 3
                If Strings.Right(CharSpace(i), 1) = "." Then
                    HexCharSpace(i) = "80"
                    CharSpace(i) = Strings.Left(CharSpace(i), Strings.Len(CharSpace(i)) - 1)
                End If

                For j = LBound(DisplayChar) To UBound(DisplayChar)
                    If CharSpace(i) = DisplayChar(j) Then
                        HexCharSpace(i) = hexadder(HexCharSpace(i), DisplayHexChar(j))
                    End If
                Next j
            Next i

            If lengthByte = 1 Then
                GrabPara_Display = HexCharSpace(0)
            ElseIf lengthByte = 2 Then
                GrabPara_Display = HexCharSpace(0) & " " & HexCharSpace(1)
            Else
                GrabPara_Display = HexCharSpace(0) & " " & HexCharSpace(1) & " " & HexCharSpace(2) & " " & HexCharSpace(3)
            End If

        Else
            If lengthByte = 1 Then
                GrabPara_Display = "00"
            ElseIf lengthByte = 2 Then
                GrabPara_Display = "00 00"
            Else
                GrabPara_Display = "00 00 00 00"
            End If
        End If
    End Function

    Private Function GrabPara_Clock(ByVal inputString As String) As String 'added from ver 5.2 for new data type $clock
        Dim LArray_GP() As String, YY As Integer, MM As Integer, DD As Integer, h As Integer, m As Integer, s As Integer
        Dim DArray() As String, TArray() As String, BinClock As String, ADate As Boolean
        YY = 0 : MM = 0 : DD = 0 : h = 0 : m = 0 : s = 0 : ADate = True
        GrabPara_Clock = "00000000"
        BinClock = ""
        LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
        If LArray_GP(0) = "clock" Then
            If Strings.InStr(LArray_GP(1), "-") > 1 Then
                'means you have date and time
                If isValidDate(Strings.Left(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") - 1)) And isValidTime(Strings.Mid(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") + 1)) Then
                    DArray = Strings.Split(Strings.Left(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") - 1), "/")
                    TArray = Strings.Split(Strings.Mid(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") + 1), ":")
                    YY = CInt(Strings.Right("00" & DArray(0), 2))
                    MM = CInt(DArray(1))
                    DD = CInt(DArray(2))
                    h = CInt(TArray(0))
                    m = CInt(TArray(1))
                    s = CInt(TArray(2))
                    BinClock = xlWorkFunc.Dec2Bin(YY, 6) & xlWorkFunc.Dec2Bin(MM, 4) & xlWorkFunc.Dec2Bin(DD, 5) & xlWorkFunc.Dec2Bin(h, 5) & xlWorkFunc.Dec2Bin(m, 6) & xlWorkFunc.Dec2Bin(s, 6)
                    GrabPara_Clock = xlWorkFunc.Bin2Hex(Strings.Left(BinClock, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(BinClock, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(BinClock, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(BinClock, 8), 2)
                End If
            Else
                'can be date only
                If isValidDate(LArray_GP(1)) Then
                    DArray = Strings.Split(LArray_GP(1), "/")
                    YY = CInt(Strings.Right("00" & DArray(0), 2))
                    MM = CInt(DArray(1))
                    DD = CInt(DArray(2))
                    BinClock = xlWorkFunc.Dec2Bin(YY, 6) & xlWorkFunc.Dec2Bin(MM, 4) & xlWorkFunc.Dec2Bin(DD, 5) & xlWorkFunc.Dec2Bin(h, 5) & xlWorkFunc.Dec2Bin(m, 6) & xlWorkFunc.Dec2Bin(s, 6)
                    GrabPara_Clock = xlWorkFunc.Bin2Hex(Strings.Left(BinClock, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(BinClock, 9, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Mid(BinClock, 17, 8), 2) & xlWorkFunc.Bin2Hex(Strings.Right(BinClock, 8), 2)
                End If
            End If
        End If
    End Function


    Function Grab_name(ByVal inputString As String) As String
        Dim LArray_v1() As String, i As Integer
        LArray_v1 = Strings.Split(Strings.Replace(Strings.Replace(inputString, "#", "$"), " ", ""), "$")

        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(LArray_v1(i), "name=") = 1 Then
                Grab_name = Strings.Replace(LArray_v1(i), "name=", "")
                If Len(Grab_name) < 1 Then
                    Grab_name = "NULL"
                End If
                Exit Function
            End If
        Next i
        Grab_name = "NULL"
    End Function

End Class
