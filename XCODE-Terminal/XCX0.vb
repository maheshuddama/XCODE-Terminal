Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
Imports System.Text

Public Class XCX0

    Dim unAttended As Boolean
    Dim RCUVarDictionary As New Dictionary(Of String, String)

    'constructors

    Public Sub New(Optional thisAttended As Boolean = False)
        unAttended = thisAttended
        UpdateVarDictionary()
    End Sub

    Private Sub UpdateVarDictionary()
        Dim RcuVarList As String() = "04,05,06,07,08,09,0A,0B,0C,0D,0E,0F,10,11,12,13,14,14,33,15,16,32,17,18,19,1A,1B,1C,1D,1E,1F,20,21,22,23,24,25,26,27,30,31".Split(","c)
        Dim RcuVarNameList As String() = "anti_ice,checkin,thermo1,thermo2,thermo3,thermo4,thermo5,thermo6,temp1,temp2,temp3,temp4,temp5,temp6,season,intervention,room_empty,room_empty_temp,room_empty_thermo,user,workflow,auto_dnd,gs1,gs2,gs3,gs4,gs5,gs6,gs7,gs8,gs9,gs10,ioexp1,ioexp2,ioexp3,ioexp4,ioexp5,ioexp6,tag_minmax,inter_fan_offset,inter_valve_offset".Split(","c)

        For i As Integer = LBound(RcuVarNameList) To UBound(RcuVarNameList)
            RCUVarDictionary(RcuVarNameList(i)) = RcuVarList(i)
        Next
    End Sub

    ' XC terminal elements

    Sub XC_Terminal_Header()
        If unAttended Then
            Console.WriteLine("XC-Terminal Ver(" & GetXCVer() & "." & GetXCbuildVer() & ")" & vbNewLine)
        Else
            Dim logo As String = "" & vbNewLine &
         "    ██╗  ██╗ ██████╗" & vbNewLine &
         "    ╚██╗██╔╝██╔════╝" & vbNewLine &
         "     ╚███╔╝ ██║     " & vbNewLine &
         "     ██╔██╗ ██║     " & vbNewLine &
         "    ██╔╝ ██╗╚██████╗" & vbNewLine &
         "    ╚═╝  ╚═╝ ╚═════╝"
            Console.WriteLine(logo & vbNewLine & "     XC-Terminal Ver(" & GetXCVer() & "." & GetXCbuildVer() & ")" & vbNewLine)
        End If
    End Sub

    Sub ConsoleProgress(StageDiscription As String, StageProgress As Double)
        Dim progressBarWidth As Integer = 20
        If unAttended Then
            'ConsoleEraseThisLine()
            'Console.Write("XC Progress:> " & StageDiscription & "> [{0}{1}]{2}%", New String("#"c, StageProgress * progressBarWidth / 100), New String(" "c, progressBarWidth - (StageProgress * progressBarWidth / 100)), CInt(StageProgress))
        Else
            ConsoleEraseThisLine()
            Console.Write("XC Progress:> " & StageDiscription & "> [{0}{1}]{2}%", New String("█"c, StageProgress * progressBarWidth / 100), New String(" "c, progressBarWidth - (StageProgress * progressBarWidth / 100)), CInt(StageProgress))
        End If
    End Sub

    Sub ConsoleMsg(thisMessage As String, Optional Overwrite As Boolean = True)
        If (Not unAttended) AndAlso Overwrite Then
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

    Function Get_memmap_name_from_words(inputWords As String(), Optional index As Integer = 1) As String
        Dim w As Integer
        Get_memmap_name_from_words = inputWords(index)

        If UBound(inputWords) > index Then
            For w = (index + 1) To UBound(inputWords)
                Get_memmap_name_from_words = Get_memmap_name_from_words & " " & inputWords(w)
                If Strings.Right(Strings.LCase(inputWords(w)), 5) = ".xlsx" Then
                    Exit For
                End If
            Next w
        End If
    End Function

    ' Mem map evaluation
    Function GetRCUList(ByRef xlSheetList As List(Of String)) As List(Of String)
        Dim L As Integer
        Dim thisRCUList As New List(Of String)

        For L = 1 To 64
            If IsSheetExist_From_List(xlSheetList, "Device" & L) AndAlso
                IsSheetExist_From_List(xlSheetList, "XCODE" & L) AndAlso
                IsSheetExist_From_List(xlSheetList, "Program" & L) AndAlso
                IsSheetExist_From_List(xlSheetList, "Settings" & L) AndAlso
                IsSheetExist_From_List(xlSheetList, "CustomVar" & L) Then
                thisRCUList.Add(L)
            End If
            ConsoleProgress("Listing RCUs", CInt((L / 64) * 100))
        Next L

        If thisRCUList.Count = 0 Then
            If IsSheetExist_From_List(xlSheetList, "Device") AndAlso
                IsSheetExist_From_List(xlSheetList, "XCODE") AndAlso
                IsSheetExist_From_List(xlSheetList, "Program") AndAlso
                IsSheetExist_From_List(xlSheetList, "Settings") AndAlso
                IsSheetExist_From_List(xlSheetList, "CustomVar") Then
                thisRCUList.Add("")
            Else
                thisRCUList.Add("NULL")
            End If
        End If

        GetRCUList = thisRCUList

    End Function

    Function NoVariableErrors(ByRef thisWorkbook As Excel.Workbook, ByRef xlSheetList As List(Of String), ByRef thisRCUList As List(Of String), ByRef ErrorWarnLog As String()) As Boolean
        NoVariableErrors = True
        Dim xlVarSheet As Excel.Worksheet
        Dim ce As Excel.Range
        Dim C_VarShort As String, C_VarLong As String
        Dim M_VarShort As String, M_VarLong As String

        For Each thisRCU As String In thisRCUList


            xlVarSheet = thisWorkbook.Sheets("CustomVar" & thisRCU) 'this sheet should exist because its being checked before
            C_VarShort = "" : C_VarLong = ","

            If IsVarSheetNotEmpty(xlVarSheet) Then
                For Each ce In xlVarSheet.Range(xlVarSheet.Range("B3").Offset(1, 0), xlVarSheet.Range("B3").End(Excel.XlDirection.xlDown))

                    If Not IsValidHexString(ce.Text) Then
                        ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Non-Hexadecimal value as variable short name]")
                        NoVariableErrors = True
                    End If

                    If Strings.InStr(Strings.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                        ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Custom Variable definition with spaces]")
                        NoVariableErrors = True
                    End If

                    If Strings.InStr(C_VarShort, Strings.Right("00" & Strings.LCase(ce.Text), 2) & ",") > 0 Then
                        ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Duplicate variable]")
                        NoVariableErrors = True
                    Else
                        C_VarShort = C_VarShort & Strings.Right("00" & Strings.LCase(ce.Text), 2) & ","
                    End If

                    If Strings.Len(ce.Offset(0, 1).Text) > 0 Then
                        If Strings.InStr(C_VarLong, "," & Strings.LCase(ce.Offset(0, 1).Text) & ",") > 0 Then
                            ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Duplicate variable definition]")
                            NoVariableErrors = True
                        Else
                            C_VarLong = C_VarLong & Strings.LCase(ce.Offset(0, 1).Text) & ","
                        End If
                    Else
                        ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Empty variable definition]")
                        NoVariableErrors = True
                    End If
                Next ce
            End If


            'optional check for Modbus var

            If IsSheetExist_From_List(xlSheetList, "ModbusVar" & thisRCU) Then

                xlVarSheet = thisWorkbook.Sheets("ModbusVar" & thisRCU)
                M_VarShort = "" : M_VarLong = ","

                If IsVarSheetNotEmpty(xlVarSheet) Then
                    For Each ce In xlVarSheet.Range(xlVarSheet.Range("B3").Offset(1, 0), xlVarSheet.Range("B3").End(Excel.XlDirection.xlDown))

                        If Not IsValidHexString(ce.Text) Then
                            ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Non-Hexadecimal value as variable short name]")
                            NoVariableErrors = False
                        End If

                        If Strings.InStr(Strings.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                            ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Modbus Variable definition with spaces]")
                            NoVariableErrors = False
                        End If

                        If Strings.InStr(M_VarShort, Strings.Right("00" & Strings.LCase(ce.Text), 2) & ",") > 0 Then
                            ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Duplicate variable]")
                            NoVariableErrors = False
                        Else
                            M_VarShort = M_VarShort & Strings.Right("00" & Strings.LCase(ce.Text), 2) & ","
                        End If

                        If Strings.Len(ce.Offset(0, 1).Text) > 0 Then
                            If Strings.InStr(M_VarLong, "," & Strings.LCase(ce.Offset(0, 1).Text) & ",") > 0 Then
                                ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Duplicate variable definition]")
                                NoVariableErrors = False
                            Else
                                M_VarLong = M_VarLong & Strings.LCase(ce.Offset(0, 1).Text) & ","
                            End If
                        Else
                            ErrorWarnLog(0) = AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Empty variable definition]")
                            NoVariableErrors = False
                        End If
                    Next ce
                End If


            End If

        Next

    End Function

    Function IsVarSheetNotEmpty(ByRef thisWorkSheet As Excel.Worksheet) As Boolean
        If Strings.Len(thisWorkSheet.Range("B3").Offset(1, 0).Text) <> 0 Then
            Return True
        End If
        Return False
    End Function

    'Function isSheetExist(ByRef thisWorkBook As Excel.Workbook, thisSheet As String) As Boolean
    '    isSheetExist = False

    '    Dim xlSheet As Excel.Worksheet

    '    For Each xlSheet In thisWorkBook.Sheets
    '        If xlSheet.Name = thisSheet Then
    '            isSheetExist = True
    '            Exit For
    '        End If
    '    Next xlSheet

    'End Function

    Function IsSheetExist_From_List(ByRef xlSheetList As List(Of String), thisSheet As String) As Boolean
        IsSheetExist_From_List = False
        If xlSheetList.Contains(thisSheet) Then
            IsSheetExist_From_List = True
        End If
    End Function

    'Function isSheetExist(ByRef thisWorkBook As Excel.Workbook, thisSheet As String) As Boolean
    '    Return thisWorkBook.Worksheets.Cast(Of Excel.Worksheet)().Any(Function(sheet) sheet.Name = thisSheet)
    'End Function

    Function IsValidHexString(inputString As String) As Boolean
        Return Regex.IsMatch(inputString, "^[0-9A-Fa-f]+$")
    End Function

    Function IsValidBinString(inputString As String) As Boolean
        Return Regex.IsMatch(inputString, "^[01]+$")
    End Function

    Function IsValidInteger(inputString As String, Optional lowerLimit As Integer = 0) As Boolean
        If IsNumeric(inputString) Then
            If (CInt(inputString) = CDbl(inputString)) AndAlso CInt(inputString) >= lowerLimit Then
                Return True
            End If
        End If

        Return False
    End Function

    Function Evaluate(thisExpression As String) As Double
        Dim table As New DataTable()
        Dim result As Object = table.Compute(thisExpression, Nothing)
        Return Convert.ToDouble(result)
    End Function

    Function GetFilePathFormat(filePath As String) As Integer 'added in XCODE 4.9.2, detect if local path accessible
        '0 other not accessible, 1 as local path, 2 as onedrive path,
        'removed letter c checking from path. in XCODE ver 5.2
        filePath = Strings.LCase(filePath)
        If InStr(filePath, ":\") = 2 Then
            GetFilePathFormat = 1
        ElseIf InStr(filePath, "http") = 1 Then
            Dim onedrivePath As String
            onedrivePath = LCase(Environ("onedrive"))
            If InStr(onedrivePath, ":\") = 2 Then
                GetFilePathFormat = 2
            Else
                GetFilePathFormat = 0
            End If
        Else
            GetFilePathFormat = 0
        End If
    End Function

    Function Getfilepath(filePath As String, Optional toLower As Boolean = True) As String 'added in XCODE 4.9.2, redirect to onedrive local path
        If toLower Then
            filePath = Strings.LCase(filePath)
        End If

        If InStr(filePath, "http") = 1 Then
            Dim onedrivePath As String
            onedrivePath = Environ("onedrive")
            Getfilepath = Strings.Replace(onedrivePath & Mid(filePath, InStr(filePath, "Documents") + 9), "/"c, "\"c)
        Else
            Getfilepath = filePath
        End If
    End Function

    Private Function IsValidDate(inputString As String) As Boolean
        Dim DArr() As String, Y1 As Integer, M1 As Integer, D1 As Integer
        IsValidDate = False
        If Strings.InStr(inputString, "/"c) > 2 Then
            DArr = inputString.Split("/"c)
            If UBound(DArr) = 2 Then
                If IsNumeric(DArr(0)) AndAlso IsNumeric(DArr(1)) AndAlso IsNumeric(DArr(2)) Then
                    Y1 = CInt(Strings.Right("00" & DArr(0), 2))
                    M1 = CInt(DArr(1))
                    D1 = CInt(DArr(2))

                    If Y1 > 10 AndAlso Y1 < 100 AndAlso M1 > 0 AndAlso M1 < 13 AndAlso D1 > 0 AndAlso D1 < 32 Then
                        IsValidDate = True
                    End If
                End If
            End If
        End If
    End Function

    Private Function IsValidTime(inputString As String) As Boolean
        Dim TArr() As String, H1 As Integer, M1 As Integer, S1 As Integer
        IsValidTime = False
        If Strings.InStr(inputString, ":"c) > 2 Then
            TArr = inputString.Split(":"c)
            If UBound(TArr) = 2 Then
                If IsNumeric(TArr(0)) AndAlso IsNumeric(TArr(1)) AndAlso IsNumeric(TArr(2)) Then
                    H1 = CInt(TArr(0))
                    M1 = CInt(TArr(1))
                    S1 = CInt(TArr(2))

                    If H1 < 24 AndAlso M1 < 60 AndAlso S1 < 60 Then
                        IsValidTime = True
                    End If
                End If
            End If
        End If
    End Function

    Function CommaExpand(inputString As String, lowerLimit As Integer, upperLimit As Integer) As String
        Dim relayvalues() As String, rangeSplit() As String
        Dim Lrelay As Integer, Urelay As Integer, i As Integer, k As Integer
        Dim ArrExp(upperLimit - lowerLimit) As Integer

        If inputString.Contains("."c) Then
            'addWarn("[" & inputString & "][A dot detected instead of a comma]")
            inputString = inputString.Replace("."c, ","c)
        End If

        If inputString.Contains("to") Then
            inputString = inputString.Replace("to", "-"c)
        End If

        relayvalues = inputString.Split(","c)
        CommaExpand = ""

        For i = LBound(relayvalues) To UBound(relayvalues)

            If relayvalues(i).Contains("-"c) Then

                rangeSplit = relayvalues(i).Split("-"c)

                If IsNumeric(rangeSplit(0)) AndAlso (Not IsNumeric(rangeSplit(1))) Then

                    If CInt(rangeSplit(0)) >= lowerLimit AndAlso CInt(rangeSplit(0)) <= upperLimit Then
                        ArrExp(CInt(rangeSplit(0)) - lowerLimit) = 1
                    End If

                ElseIf IsNumeric(rangeSplit(1)) AndAlso (Not IsNumeric(rangeSplit(0))) Then

                    If CInt(rangeSplit(1)) >= lowerLimit AndAlso CInt(rangeSplit(1)) <= upperLimit Then
                        ArrExp(CInt(rangeSplit(1)) - lowerLimit) = 1
                    End If

                ElseIf (Not IsNumeric(rangeSplit(1))) AndAlso (Not IsNumeric(rangeSplit(0))) Then

                ElseIf IsNumeric(rangeSplit(0)) AndAlso IsNumeric(rangeSplit(1)) Then
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
                If CInt(relayvalues(i)) >= lowerLimit AndAlso CInt(relayvalues(i)) <= upperLimit Then
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

    Function GetIndexFromBase(inputString As String, baseString As String) As Integer

        Dim pos As Integer = inputString.IndexOf(baseString, StringComparison.OrdinalIgnoreCase)

        Dim tail As String = inputString.Substring(pos + baseString.Length)

        Return CInt(Val(tail))
    End Function

    Function GetAllParameters(inputString As String) As String
        Dim dataArr() As String, i As Integer
        Dim output As New StringBuilder

        dataArr = inputString.Split(" "c)

        For i = 0 To dataArr.Length - 1
            If dataArr(i).Contains("$"c) Then
                output.Append(" " & dataArr(i))
            End If
        Next

        Return output.ToString()
    End Function

    Private Function Hexadder(hex1 As String, hex2 As String) As String
        ' Convert hex strings to integers and sum them
        Dim sum As Integer = Convert.ToInt64(hex1, 16) + Convert.ToInt64(hex2, 16)

        ' Convert sum back to Hex formatted as a 2-digit zero-padded string ("X2")
        Dim hexResult As String = sum.ToString("X2")

        ' Keep the rightmost 2 characters (handles potential sum overflow above 0xFF)
        If hexResult.Length > 2 Then
            Return hexResult.Substring(hexResult.Length - 2)
        Else
            Return hexResult
        End If
    End Function

    'encription/decription
    '53 70 84 85
    '84 101 115 116
    'T e s t

    Function Encript(inputstring As String) As String
        Dim i As Integer
        Encript = ""
        For i = 1 To Strings.Len(inputstring)
            Encript = Encript & Strings.Right("00" & (Strings.Asc(Strings.Mid(inputstring, i, 1)) - 31), 2)
        Next i
    End Function

    Function Decript(encripted As String) As String
        Dim x As Integer
        Decript = ""
        For x = 1 To Strings.Len(encripted) / 2
            Decript = Decript & Chr(Strings.Mid(encripted, 2 * x - 1, 2) + 31)
        Next x
    End Function

    'connection
    Function GetOpenAuth() As String
        GetOpenAuth = Decript("5451452873858581842716166980688415728080727770156880781684818370666984737070858416691618145070488739197279688559537667241418642181194948498977775685225117896456908558257081681670697485328684813084736683747972")
    End Function

    Function GetLogAuth() As String
        GetLogAuth = Decript("73858581842716166980688415728080727770156880781671808378841669167016183934428150455270505620815026668339583652646784672179855519377857495358395651568856408138785034724650429123883416718083785170848180798470")
    End Function

    'URL encode
    Function URLEncode(StringToEncode As String, Optional PlusSpace As Boolean = True) As String
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
    Function GetXCVer() As String
        GetXCVer = "5.7" 'major minor versions
    End Function

    Function GetXCbuildVer() As String
        GetXCbuildVer = "0" ' build version 
    End Function

    'utc functions
    Function GetTimeStamp() As String
        GetTimeStamp = Format(Now, "yyMMdd-HHmmss")
    End Function

    Function ExpandTimeStamp(thisTimeStamp As String) As String ' ver 5.0
        Dim Months() As String, thisH As Integer, AMPM As String
        Months = Split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ","c)

        thisH = CInt(Strings.Mid(thisTimeStamp, 8, 2))
        If thisH > 12 Then
            AMPM = " PM"
            thisH = thisH - 12
        Else
            AMPM = " AM"
        End If

        ExpandTimeStamp = thisH & ":" & Strings.Mid(thisTimeStamp, 10, 2) & ":" & Strings.Mid(thisTimeStamp, 12, 2) & AMPM & " - " & Months(CInt(Strings.Mid(thisTimeStamp, 3, 2)) - 1) & " " & Strings.Mid(thisTimeStamp, 5, 2) & ", 20" & Strings.Left(thisTimeStamp, 2)
    End Function

    Function BuildDate_hex() As String
        ' Capture a single DateTime snapshot
        Dim now As DateTime = DateTime.Now

        ' Bit field extractions:
        ' Year: 7 bits (Year Mod 1000)
        ' Month: 4 bits
        ' Day: 5 bits
        Dim yy As UShort = CUShort(now.Year Mod 1000)
        Dim mm As UShort = CUShort(now.Month)
        Dim dd As UShort = CUShort(now.Day)

        ' Pack into a 16-bit integer (7 bits + 4 bits + 5 bits = 16 bits total)
        ' Layout: [Year: 7 bits] [Month: 4 bits] [Day: 5 bits]
        Dim packed As UShort = CUShort((yy << 9) Or (mm << 5) Or dd)

        ' Convert to 4-character Hex string (2 bytes)
        Dim hex As String = packed.ToString("X4")

        ' Returns Byte 1 (Left 8 bits) and Byte 2 (Right 8 bits) separated by space
        Return $"{hex.Substring(0, 2)} {hex.Substring(2, 2)}"
    End Function

    'Function BuildDate_hex() As String
    '    BuildDate_hex = DateTime.Now.Year Mod 1000
    '    BuildDate_hex = xlWorkFunc.Dec2Bin(BuildDate_hex, 7)
    '    BuildDate_hex = BuildDate_hex & xlWorkFunc.Dec2Bin(DateTime.Now.Month, 4) & xlWorkFunc.Dec2Bin(DateTime.Now.Day, 5)
    '    BuildDate_hex = xlWorkFunc.Bin2Hex(Strings.Left(BuildDate_hex, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Right(BuildDate_hex, 8), 2)
    'End Function

    'Function Give_Time_1() As String
    '    'Dim YY As Integer, MM As Integer, DD As Integer, h As Integer, m As Integer, s As Integer

    '    Dim YY As String = DateTime.Now.ToString("yy")
    '    Dim MM As String = DateTime.Now.ToString("MM")
    '    Dim DD As String = DateTime.Now.ToString("dd")
    '    Dim h As String = DateTime.Now.ToString("HH")
    '    Dim m As String = DateTime.Now.ToString("mm")
    '    Dim s As String = DateTime.Now.ToString("ss")

    '    'YY = Year(Now)
    '    'MM = Month(Now)
    '    'DD = Day(Now)
    '    'h = Hour(Format(Now, "MM/dd/yyyy HH:mm:ss"))
    '    'm = Minute(Now)
    '    's = Second(Now)
    '    ' sec 60(64)6 : min 60(64)6 : hr 24(32)5 : day 31(32)5 : month 12(16)4 : year 64(64)6
    '    Give_Time_1 = xlWorkFunc.Dec2Bin(Strings.Right(YY, 2), 6) & xlWorkFunc.Dec2Bin(MM, 4) & xlWorkFunc.Dec2Bin(DD, 5) & xlWorkFunc.Dec2Bin(h, 5) & xlWorkFunc.Dec2Bin(m, 6) & xlWorkFunc.Dec2Bin(s, 6)
    '    Give_Time_1 = xlWorkFunc.Bin2Hex(Strings.Right(Give_Time_1, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Mid(Give_Time_1, 17, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Mid(Give_Time_1, 9, 8), 2) & " " & xlWorkFunc.Bin2Hex(Strings.Left(Give_Time_1, 8), 2)
    'End Function

    Function Give_Time_1() As String
        ' Capture current time in a single snapshot to prevent second-boundary drift
        Dim now As DateTime = DateTime.Now

        ' Extract numeric components
        Dim yy As UInteger = CUInt(now.Year Mod 100) ' 2-digit year (0-99)
        Dim mm As UInteger = CUInt(now.Month)        ' Month (1-12)
        Dim dd As UInteger = CUInt(now.Day)          ' Day (1-31)
        Dim h As UInteger = CUInt(now.Hour)         ' Hour (0-23)
        Dim m As UInteger = CUInt(now.Minute)       ' Minute (0-59)
        Dim s As UInteger = CUInt(now.Second)       ' Second (0-59)

        ' Pack into a 32-bit integer matching the exact bit layout:
        ' YY (6 bits) | MM (4 bits) | DD (5 bits) | h (5 bits) | m (6 bits) | s (6 bits)
        Dim packed As UInteger = (yy << 26) Or (mm << 22) Or (dd << 17) Or (h << 12) Or (m << 6) Or s

        ' Convert to 8-character hex string (e.g., "A1B2C3D4")
        Dim hex As String = packed.ToString("X8")

        ' Original return byte order: Byte 4, Byte 3, Byte 2, Byte 1
        ' Hex string indices:
        ' Byte 1 (Left 8)  = hex.Substring(0, 2)
        ' Byte 2 (Mid 9)   = hex.Substring(2, 2)
        ' Byte 3 (Mid 17)  = hex.Substring(4, 2)
        ' Byte 4 (Right 8) = hex.Substring(6, 2)
        Return $"{hex.Substring(6, 2)} {hex.Substring(4, 2)} {hex.Substring(2, 2)} {hex.Substring(0, 2)}"
    End Function



    'error and warn apend
    Function AddErrorOrWarn(OriginalString As String, thisErrorOrWarn As String) As String
        If Strings.InStr(OriginalString & " ", thisErrorOrWarn & " ") = 0 Then
            AddErrorOrWarn = OriginalString & vbNewLine & thisErrorOrWarn
        Else
            AddErrorOrWarn = OriginalString
        End If
    End Function

    'file functions
    Function IsFileExist(Fpath As String) As Boolean
        If File.Exists(Fpath) Then
            Return True
        End If
        Return False
    End Function

    Function IsFileInUse(FileName As String) As Boolean
        Try
            Using f As New IO.FileStream(FileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            End Using
        Catch Ex As Exception
            Return True
        End Try
        Return False
    End Function

    Function IsAnExcelFile(FileName As String) As Boolean
        If FileName.ToLower.EndsWith(".xlsx") Then
            Return True
        End If

        Return False
    End Function

    'is IP
    'ip checker
    Function IsAValidIP(inputip As String) As Boolean
        IsAValidIP = False
        If inputip.Contains("."c) Then

            Dim LArray_ip() As String, L As Integer
            LArray_ip = inputip.Split("."c)

            If UBound(LArray_ip) = 3 Then

                For L = LBound(LArray_ip) To UBound(LArray_ip)
                    If IsNumeric(LArray_ip(L)) Then
                        If CInt(LArray_ip(L)) < 0 OrElse CInt(LArray_ip(L)) > 255 Then
                            Exit Function
                        End If
                    Else
                        Exit Function 'if atleast one non numeric found exit function with false
                    End If
                Next L

                IsAValidIP = True
            End If
        End If
    End Function

    'name version
    Function Getsufixname(MemmapName As String) As String
        Return MemmapName.Replace("RCU mem map.", "").Replace(".xlsx", "").Replace(".XLSX", "")
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

    Sub SetFunctionSeparaters(ByRef thisProgramSheet As Excel.Worksheet)
        Dim ce As Excel.Range, i As Integer
        For Each ce In thisProgramSheet.Range(thisProgramSheet.Range("E5"), thisProgramSheet.Range("E5").End(Excel.XlDirection.xlDown))
            If ce.Text = "04 01 00 00 00 00 00 00;" OrElse (Strings.Left(ce.Text, 5) = "01 01" AndAlso Strings.Mid(ce.Text, 7, 5) <> "00 00") Then
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
    Sub Print_XCODE_build(ByRef thisXCODESheet As Excel.Worksheet)
        Dim XCODEConfig As String, XConfigArr() As String
        XCODEConfig = Strings.LCase(Strings.Trim(thisXCODESheet.Range("C1").Text))
        If Strings.InStr(XCODEConfig, "]") > 0 AndAlso Strings.InStr(XCODEConfig, "[") > 0 Then
            XCODEConfig = Strings.Mid(XCODEConfig, Strings.InStr(XCODEConfig, "[") + 1, Strings.InStr(XCODEConfig, "]") - Strings.InStr(XCODEConfig, "[") - 1)
            XConfigArr = XCODEConfig.Split(","c)
            If UBound(XConfigArr) > 0 Then
                thisXCODESheet.Range("C1").FormulaR1C1 = "XCODE [" & GetXCVer() & "," & XConfigArr(1) & "]"
            End If
        End If
    End Sub

    Sub Check_XCODE_format(MemmapName As String, ByRef thisXCODESheet As Excel.Worksheet, thisRCU As String)
        Dim L As Integer, m As Integer, L1 As Integer, L2 As Integer, L3 As Integer, L4 As Integer, thisLine As String, V As Integer
        L = 0 : L1 = 1 : L2 = 1 : L3 = 1 : L4 = 1 : m = 1 : V = 0 'get first 2 lines

        While (L < 2 AndAlso m < 100)
            If Strings.Len(thisXCODESheet.Range("C" & m + 2).Text) > 2 Then
                thisLine = Strings.Trim(Strings.LCase(thisXCODESheet.Range("C" & m + 2).Text))
                If Strings.InStr(thisLine, "//") <> 1 Then
                    If Strings.InStr(thisLine, "//") > 4 Then
                        thisLine = Strings.Trim(Strings.Left(thisLine, Strings.InStr(thisLine, "//") - 1))
                    End If

                    If thisLine = "debug" Then
                        L = L + 1
                        L1 = m + 2
                    ElseIf thisLine = "exit" AndAlso L1 > 1 Then
                        L = L + 1
                        L2 = m + 2
                    ElseIf thisLine.StartsWith("version") Then
                        L = L + 1
                        L3 = m + 2
                    ElseIf thisLine = "goto #debugging" AndAlso L3 > 1 Then
                        L = L + 1
                        L4 = m + 2
                    End If
                    V = V + 1
                End If
            End If
            m = m + 1
        End While


        If V = 2 AndAlso L1 > 1 AndAlso L2 > L1 Then
            thisXCODESheet.Range("C" & L1).FormulaR1C1 = "version [" & Encript(Getsufixname(MemmapName)) & ",-1]"
            thisXCODESheet.Range("C" & L2).ClearContents()
            thisXCODESheet.Range("C" & L1 + 1).FormulaR1C1 = "goto #debugging"
            ConsoleMsg("XC Progress:> Formatting XCODE" & thisRCU)
            Move_add_debugging(thisXCODESheet)
        ElseIf V = 2 AndAlso L3 > 1 AndAlso L4 > L3 Then
            Move_add_debugging(thisXCODESheet)
            Check_Version_Format(MemmapName, thisXCODESheet, L3)
        End If
    End Sub

    Private Sub Move_add_debugging(ByRef thisXCODESheet As Excel.Worksheet)
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

            While k < 200 AndAlso L2 = 1
                This_Line = Strings.LCase(Strings.Trim(thisXCODESheet.Range("C" & k).Text)) & " "
                If Strings.InStr(This_Line, "exit ") = 1 OrElse Strings.InStr(This_Line, "goto ") = 1 Then
                    L1 = k
                ElseIf Strings.InStr(This_Line, "function ") = 1 AndAlso L1 > 1 Then
                    L2 = k
                End If
                k = k + 1
            End While

            thisXCODESheet.Select()
            thisXCODESheet.Range("C" & L1).Select()

            Push_code_down(thisXCODESheet, "")
            Push_code_down(thisXCODESheet, "")
            Push_code_down(thisXCODESheet, "function debugging")
            Push_code_down(thisXCODESheet, "debug")
            Push_code_down(thisXCODESheet, "exit")
            Push_code_down(thisXCODESheet, "")
            Push_code_down(thisXCODESheet, "")
        End If
    End Sub

    Private Sub Check_Version_Format(MemmapName As String, ByRef thisXCODESheet As Excel.Worksheet, vLine As Integer)
        Dim thisLine As String, vData() As String
        thisLine = Strings.Trim(Strings.LCase(thisXCODESheet.Range("C" & vLine).Text))

        If Strings.InStr(thisLine, "[") < Strings.InStr(thisLine, "]") Then
            vData = Strings.Split(Strings.Mid(thisLine, Strings.InStr(thisLine, "[") + 1, Strings.InStr(thisLine, "]") - Strings.InStr(thisLine, "[") - 1), ",")
            If vData(0) <> Encript(Getsufixname(MemmapName)) Then
                thisXCODESheet.Range("C" & vLine).FormulaR1C1 = "version [" & Encript(Getsufixname(MemmapName)) & ",-1]"
            End If
        Else
            thisXCODESheet.Range("C" & vLine).FormulaR1C1 = "version [" & Encript(Getsufixname(MemmapName)) & ",-1]"
        End If

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

    'xcode core 1
    ' rcu var
    Function HasValidVar(inputString As String) As Boolean 'RCU var

        If Strings.InStr(Strings.LCase(inputString), " $var") > 1 Then
            Dim LArray_T2() As String, V As String
            V = ""
            LArray_T2 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If Strings.InStr(LArray_T2(i).Replace(" "c, ""), "var=") Then
                    V = LArray_T2(i).Replace("var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
                'For i = LBound(RcuVariableList) To UBound(RcuVariableList)
                '    If V = RcuVariableList(i) Or V = RcuNameList(i) Then
                '        hasValidVar = True
                '        Exit For
                '    End If
                'Next i
                If RCUVarDictionary.ContainsKey(V) OrElse RCUVarDictionary.ContainsValue(V) Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Function Grab_Var_1(inputString As String) As String 'RCU var

        Dim V As String, LArray_T1() As String, P As String, i As Integer

        V = "" : P = ""
        LArray_T1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If Strings.InStr(LArray_T1(i).Replace(" "c, ""), "var=") = 1 Then
                V = LArray_T1(i).Replace("var=", "")
                Exit For
            End If
        Next i
        'For i = LBound(RcuVariableList) To UBound(RcuVariableList)
        '    If V = RcuVariableList(i) Or V = RcuNameList(i) Then
        '        P = RcuVariableList(i)
        '        Exit For
        '    End If
        'Next i
        If RCUVarDictionary.ContainsKey(V) Then
            P = RCUVarDictionary(V)
        ElseIf RCUVarDictionary.ContainsValue(V) Then
            P = V
        End If

        Grab_Var_1 = P
    End Function

    ' is function 
    Function IsAFunction(inputString As String) As Boolean
        If Strings.InStr(inputString, "function ") = 1 Then
            Return True
        End If

        Return False
    End Function

    Function GetLineFromFunctionNameOtherRCU(inputString As String, inputLine As Integer, ByRef thisWorkbook As Excel.Workbook, otherRCU As String) As String
        'Dim thisXlProgramSheet As Excel.Worksheet
        Dim thisXlFUNCTIONSheet As Excel.Worksheet
        Dim ce As Excel.Range
        GetLineFromFunctionNameOtherRCU = "0000"

        If inputString = "exit" OrElse inputString = "quit" OrElse inputString = "end" Then

            'its useless to call exit line number in a another rcu
            'thisXlProgramSheet = thisWorkbook.Sheets("Program" & thisRCU)

            'If Strings.Len(thisXlProgramSheet.Range("F" & (inputLine + 1)).Text) = 0 Then
            '    'to remove this condition from else
            'ElseIf Strings.Len(thisXlProgramSheet.Range("F" & (inputLine + 2)).Text) = 0 Then
            '    If thisXlProgramSheet.Range("F" & (inputLine + 1)).Text = inputString Then
            '        getLineFromFunctionNameOtherRCU = xlWorkFunc.Dec2Hex(inputLine - 3, 4)
            '    End If
            'Else
            '    For Each ce In thisXlProgramSheet.Range(thisXlProgramSheet.Range("F" & (inputLine + 1)), thisXlProgramSheet.Range("F" & (inputLine + 1)).End(Excel.XlDirection.xlDown))
            '        If ce.Text = inputString Then
            '            getLineFromFunctionNameOtherRCU = xlWorkFunc.Dec2Hex(ce.Row - 3, 4)
            '            Exit Function
            '        End If
            '    Next ce
            'End If
        ElseIf IsNumeric(inputString) Then
            'impractical to call an offset in another rcu
            'If (inputLine + CInt(inputString) - 3) > 0 Then
            '    getLineFromFunctionNameOtherRCU = xlWorkFunc.Dec2Hex(inputLine + CInt(inputString) - 3, 4)
            'End If
        Else

            thisXlFUNCTIONSheet = thisWorkbook.Sheets("FUNCTION" & otherRCU)

            If Strings.Len(thisXlFUNCTIONSheet.Range("E3").Text) = 0 Then
                'to remove this condition from else
            ElseIf Strings.Len(thisXlFUNCTIONSheet.Range("E4").Text) = 0 Then
                If thisXlFUNCTIONSheet.Range("E3").Text = inputString Then
                    GetLineFromFunctionNameOtherRCU = thisXlFUNCTIONSheet.Range("F3").Text
                End If
            Else
                For Each ce In thisXlFUNCTIONSheet.Range(thisXlFUNCTIONSheet.Range("E3"), thisXlFUNCTIONSheet.Range("E3").End(Excel.XlDirection.xlDown))
                    If ce.Text = inputString Then
                        GetLineFromFunctionNameOtherRCU = ce.Offset(0, 1).Text
                        Exit Function
                    End If
                Next ce
            End If
        End If
    End Function


    'is other rcu header functions
    Function IsOtherRCUcall(inputString As String) As Boolean
        If Strings.InStr(inputString, " @rcu") > 1 AndAlso (Strings.InStr(inputString, "run ") = 1 OrElse Strings.InStr(inputString, "reuse ") = 1) Then
            Return True
        End If
        Return False
    End Function

    ' is ioexp direct
    Function IsAnIOEXPDirect(inputString As String) As Boolean
        inputString = inputString.Replace(" "c, "")
        If Strings.InStr(inputString, "ioexp") = 1 AndAlso (Strings.InStr(inputString, "."c) = 6 OrElse Strings.InStr(inputString, "."c) = 7 OrElse Strings.InStr(inputString, "."c) = 8) Then
            Return True
        End If
        Return False
    End Function

    ' is iodexp direct
    Function IsAnIODEXPDirect(inputString As String) As Boolean
        inputString = inputString.Replace(" "c, "")
        If Strings.InStr(inputString, "iodexp") = 1 AndAlso (Strings.InStr(inputString, "."c) = 7 OrElse Strings.InStr(inputString, "."c) = 8) Then
            Return True
        End If
        Return False
    End Function

    ' is rcu direct
    Function IsARCUDirect(inputString As String) As Boolean
        inputString = inputString.Replace(" "c, "")
        If inputString.StartsWith("rcu.") AndAlso
        (inputString.Contains(".set$relay=") OrElse inputString.Contains(".or$relay=") OrElse inputString.Contains(".unset$relay=") OrElse
        inputString.Contains(".and$relay=") OrElse inputString.Contains(".xor$relay=")) Then
            Return True
        End If
        Return False
    End Function

    ' is math direct
    Function IsAMathDirect(inputString As String) As Boolean
        If inputString.Replace(" "c, "").StartsWith("math.") Then
            Return True
        End If
        Return False
    End Function

    ' is save direct
    Function IsASaveDirect(inputString As String) As Boolean
        If inputString.Replace(" "c, "").StartsWith("save") Then
            Return True
        End If
        Return False
    End Function

    'extend modbus direct
    Function IsAModbusDirect(inputString As String) As Boolean
        If inputString.Replace(" "c, "").StartsWith("send.mod") Then
            Return True
        End If
        Return False
    End Function

    'expansion dimmer direct
    Function IsADimmerDirect(inputString As String) As Boolean
        If inputString.Replace(" "c, "").StartsWith("send.dim") Then
            Return True
        End If
        Return False
    End Function

    ' dimmer read direct
    Function IsADimmerReadDirect(inputString As String) As Boolean
        If inputString.Replace(" "c, "").StartsWith("read.dim") Then
            Return True
        End If
        Return False
    End Function

    ' dimmer load direct 
    Function IsADimmerLoad(inputString As String) As Boolean
        If inputString.StartsWith("load dimmer ") Then
            Return True
        End If
        Return False
    End Function

    Function LoadDimmerConvert(inputString As String) As String
        LoadDimmerConvert = "10 80 01 01 " & Grab_value_2(inputString) & " " & Grab_Time(inputString, 2) & ";"
    End Function

    'dali direct
    Function IsADaliDirect(inputString As String) As Boolean
        If inputString.Replace(" "c, "").StartsWith("dali") Then
            Return True
        End If
        Return False
    End Function

    'dmx
    Function IsADMXDirect(inputString As String) As Boolean
        If inputString.Replace(" ", "").StartsWith("dmx.") Then
            Return True
        End If
        Return False
    End Function

    'definition header
    Function IsADefinition(inputString As String) As Boolean
        If inputString.StartsWith("define ") Then
            Return True
        End If
        Return False
    End Function





    'xcode core - 2

    ' has function call
    Function HasAFunction(inputString As String) As Boolean
        If inputString.Contains(" #") Then
            Return True
        End If
        Return False
    End Function

    Function IsARCU(inputString As String) As Boolean
        If inputString.StartsWith("rcu.") Then
            Return True
        End If
        Return False
    End Function

    'dmx
    Function IsADMX(inputString As String) As Boolean
        If inputString.StartsWith("dmx.") Then
            Return True
        End If
        Return False
    End Function

    ' abs wait
    Function IsWaitABS(inputString As String) As Boolean   'XCODE 5.1 wait abs instruction check
        If inputString.StartsWith("wait abs ") Then
            Return True
        End If

        Return False
    End Function

    Function WaitConvert(inputString As String) As String
        Return "01 01 00 00 " & Grab_Time(inputString) & ";"
    End Function

    ' is an exit
    'exit handling
    Function IsAnExit(inputString As String) As Boolean
        If inputString.StartsWith("exit") OrElse inputString.StartsWith("quit") OrElse inputString.StartsWith("end") Then
            Return True
        End If
        Return False
    End Function

    Function ExitConvert() As String
        Return "04 01 00 00 00 00 00 00;"
    End Function

    'debug handling
    Function IsDebug(inputString As String) As Boolean
        If inputString.StartsWith("debug") Then
            Return True
        End If
        Return False
    End Function

    Function DebugConvert() As String
        Return "40 01 01 00 00 00 00 00;"
    End Function

    'version handling
    Function IsVersion(inputString As String) As Boolean
        If inputString.StartsWith("version") Then
            Return True
        End If
        Return False
    End Function

    Function VersionConvert(inputString As String) As String
        Dim buildVersion As Integer, vData() As String
        vData = Strings.Split(Strings.Mid(inputString, Strings.InStr(inputString, "[") + 1, Strings.InStr(inputString, "]") - Strings.InStr(inputString, "[") - 1), ",")
        buildVersion = CInt(vData(1))
        VersionConvert = "00 00 00 00 " & CInt(CDbl(GetXCVer()) * 10).ToString("X2") & " " & BuildDate_hex() & " " & buildVersion.ToString("X2") & ";"
    End Function

    ' load handling
    Function IsALoad(inputString As String) As Boolean
        If inputString.StartsWith("load ") Then
            Return True
        End If
        Return False
    End Function

    ' is send
    Function IsASend(inputString As String) As Boolean
        If inputString.StartsWith("send.") Then
            Return True
        End If
        Return False
    End Function

    'read handling
    Function IsARead(inputString As String) As Boolean
        If inputString.StartsWith("read.") Then
            IsARead = True
        Else
            IsARead = False
        End If
    End Function

    'save handling
    Function IsASave(inputString As String) As Boolean
        If inputString.StartsWith("save ") Then
            IsASave = True
        Else
            IsASave = False
        End If
    End Function

    ' math handling
    Function IsAMath(inputString As String) As Boolean
        If inputString.StartsWith("math.") Then
            IsAMath = True
        Else
            IsAMath = False
        End If
    End Function

    'dimflow handling
    Function IsADimFlow(inputString As String) As Boolean
        If inputString.StartsWith("dimflow.") Then
            IsADimFlow = True
        Else
            IsADimFlow = False
        End If
    End Function

    ' grabbers
    ' end line worker functions
    Function GrabHexString(inputString As String) As String
        Dim HTL1 As String
        HTL1 = Strings.Left(inputString.Replace("0x", "").Trim().Replace(" "c, ""), 16)
        If IsValidHexString(HTL1) Then
            Return Strings.Left(HTL1, 2) & " " & Strings.Mid(HTL1, 3, 2) & " " & Strings.Mid(HTL1, 5, 2) & " " & Strings.Mid(HTL1, 7, 2) & " " & Strings.Mid(HTL1, 9, 2) & " " & Strings.Mid(HTL1, 11, 2) & " " & Strings.Mid(HTL1, 13, 2) & " " & Strings.Right(HTL1, 2) & ";"
        Else
            Return "FALSE"
        End If
    End Function

    Function GrabComment(inputString As String) As String
        If inputString.Contains("//") Then
            Dim GC1 As String() = inputString.Split(New String() {"//"}, StringSplitOptions.None)
            Return "//" & GC1(1)
        Else
            Return "//Direct hex"
        End If
    End Function

    Function GrabFunctionName(inputString As String) As String
        Dim LArray() As String
        LArray = inputString.Split(" "c)
        GrabFunctionName = LArray(1)
    End Function

    'value grabbers
    'value for OR, relay wise

    Function Grab_value_0(inputString As String, index As Integer) As String 'return indexed value otherwise null
        Dim LArray_v1() As String, tempS As String, indexcount As Integer, i As Integer
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        indexcount = 0
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            tempS = LArray_v1(i).Replace(" "c, "")
            If Strings.InStr(tempS, "bin=") = 1 OrElse Strings.InStr(tempS, "dec=") = 1 OrElse Strings.InStr(tempS, "hex=") = 1 _
            OrElse Strings.InStr(tempS, "relay=") = 1 OrElse Strings.InStr(tempS, "bit=") = 1 OrElse Strings.InStr(tempS, "bin=") = 1 _
            OrElse Strings.InStr(tempS, "fan=") = 1 OrElse Strings.InStr(tempS, "display=") = 1 OrElse Strings.InStr(tempS, "var=") Then
                Grab_value_0 = tempS
                indexcount = indexcount + 1
                If indexcount = index Then
                    Exit Function
                End If
            End If
        Next i
        Grab_value_0 = "NULL"
    End Function

    Function Grab_value_1(inputString As String, Vmode As String) As String
        Dim P As String, Q As Integer, R As String, t As Integer, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = "" : t = 0
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bin=") = 1 Then
                P = GrabPara_BIN(LArray_v1(i))
                R = Convert.ToUInt32(P, 2).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 AndAlso (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "relay")
                R = Convert.ToUInt32(P, 2).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 AndAlso (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "relay")
                R = Convert.ToUInt32(P, 2).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 AndAlso (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "bit")
                R = Convert.ToUInt32(P, 2).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 AndAlso (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "bit")
                R = Convert.ToUInt32(P, 2).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Int(LArray_v1(i), "dec")
                R = Q.ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "hex=") = 1 Then
                R = GrabPara_HEX(LArray_v1(i))
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "fan=") = 1 Then
                R = GrabPara_FAN(LArray_v1(i))
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "display=") = 1 Then
                Return GrabPara_Display(LArray_v1(i))
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "len=") = 1 Then
                t = GrabPara_Int(LArray_v1(i), "len")
                R = t.ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "clock=") = 1 Then
                R = GrabPara_Clock(LArray_v1(i))
                Return R.Substring(6, 2) & " " & R.Substring(4, 2) & " " & R.Substring(2, 2) & " " & R.Substring(0, 2)
            End If
        Next i
        Return "00 00 00 00"
    End Function

    Private Function Grab_value_2(inputString As String) As String
        Dim P As String, Q As Double, R As String, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = ""
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = CLng(CDbl(P) * 100).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = CLng(Q * 100).ToString("X8")
                Return R.Substring(6, 2) & " " & R.Substring(4, 2)
            End If
        Next i
        Return "00 00"
    End Function

    Function Grab_value_3(inputString As String, Vmode As String) As String
        Dim P As String, Q As Integer, R As String, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = ""
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bin=") = 1 Then
                P = GrabPara_BIN(LArray_v1(i))
                R = Convert.ToByte(P.Substring(16, 8), 2).ToString("X2") & Convert.ToByte(P.Substring(24, 8), 2).ToString("X2")
                Return R.Substring(0, 2) & " " & R.Substring(2, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 AndAlso (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "relay")
                R = Convert.ToByte(P.Substring(16, 8), 2).ToString("X2") & Convert.ToByte(P.Substring(24, 8), 2).ToString("X2")
                Return R.Substring(0, 2) & " " & R.Substring(2, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "relay=") = 1 AndAlso (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "relay")
                R = Convert.ToByte(P.Substring(16, 8), 2).ToString("X2") & Convert.ToByte(P.Substring(24, 8), 2).ToString("X2")
                Return R.Substring(0, 2) & " " & R.Substring(2, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 AndAlso (Vmode = "OR") Then
                P = GrabPara_BitPOS_OR(LArray_v1(i), "bit")
                R = Convert.ToByte(P.Substring(16, 8), 2).ToString("X2") & Convert.ToByte(P.Substring(24, 8), 2).ToString("X2")
                Return R.Substring(0, 2) & " " & R.Substring(2, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "bit=") = 1 AndAlso (Vmode = "AND") Then
                P = GrabPara_BitPOS_AND(LArray_v1(i), "bit")
                R = Convert.ToByte(P.Substring(16, 8), 2).ToString("X2") & Convert.ToByte(P.Substring(24, 8), 2).ToString("X2")
                Return R.Substring(0, 2) & " " & R.Substring(2, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Int(LArray_v1(i), "dec")
                R = Q.ToString("X4")
                Return R.Substring(0, 2) & " " & R.Substring(2, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "hex=") = 1 Then
                R = GrabPara_HEX(LArray_v1(i))
                Return R.Substring(4, 2) & " " & R.Substring(6, 2)
            End If
        Next i
        Grab_value_3 = "00 00"
    End Function

    Function Grab_value_4(inputString As String) As String
        Dim P As String, Q As Double, R As String, i As Integer
        Dim LArray_v1() As String
        P = "" : Q = 0 : R = ""
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = CLng(CDbl(P)).ToString("X8")
                Return R.Substring(6, 2)
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                Q = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = CLng(Q).ToString("X8")
                Return R.Substring(6, 2)
            End If
        Next i
        Return "00"
    End Function

    Function Grab_value_5(inputString As String) As String 'for dali
        Dim P As Double, R As Double, i As Integer
        Dim LArray_v1() As String
        P = 0
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = P * 254 / 100
                Grab_value_5 = CLng(R).ToString("X2")
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = P * 254 / 100
                Grab_value_5 = CLng(R).ToString("X2")
                Exit Function
            End If
        Next i
        Grab_value_5 = "00"
    End Function

    Function Grab_value_6(inputString As String) As String 'for dmx
        Dim P As Double, R As Double, i As Integer
        Dim LArray_v1() As String
        P = 0
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = P * 255 / 100
                Grab_value_6 = CLng(R).ToString("X2")
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = P * 255 / 100
                Grab_value_6 = CLng(R).ToString("X2")
                Exit Function
            End If
        Next i
        Grab_value_6 = "00"
    End Function

    Function Grab_value_7(inputString As String) As String 'for dali dtr
        Dim P As Double, R As Double, i As Integer
        Dim LArray_v1() As String
        P = 0
        LArray_v1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)
        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "lum=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "lum")
                R = P * 254 / 100
                Grab_value_7 = CLng(R).ToString("X2")
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "dec=") = 1 Then
                P = GrabPara_Dbl(Strings.Replace(LArray_v1(i), "%", ""), "dec")
                R = P
                Grab_value_7 = CLng(R).ToString("X2")
                Exit Function
            ElseIf Strings.InStr(Strings.Replace(LArray_v1(i), " ", ""), "hex=") = 1 Then
                R = GrabPara_HEX(LArray_v1(i))
                Grab_value_7 = Strings.Right(R, 2)
                Exit Function
            End If
        Next i
        Grab_value_7 = "00"
    End Function

    'variant
    Function Grab_variant(inputString As String, dataType As String) As String
        Dim LArray_v1() As String, i As Integer
        Grab_variant = "NULL"
        LArray_v1 = Strings.Split(Strings.Replace(Strings.Replace(Strings.LCase(inputString), "#"c, "$"c), " "c, ""), "$"c)

        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(LArray_v1(i), dataType & "=") = 1 Then
                Grab_variant = Strings.Replace(LArray_v1(i), dataType & "="c, "")
                If Strings.Len(Grab_variant) < 1 Then
                    Grab_variant = "NULL"
                End If
                Exit Function
            End If
        Next
    End Function

    'grab rcu num
    Function Grab_RCU_Num(inputString As String) As String
        Dim LArray_v1() As String, i As Integer, RcuNum As String
        Grab_RCU_Num = "NULL"

        LArray_v1 = inputString.Trim().Split(" "c)

        For i = LBound(LArray_v1) To UBound(LArray_v1)
            If Strings.InStr(LArray_v1(i), "rcu") = 1 Then
                RcuNum = LArray_v1(i).Replace("rcu", "")
                If IsValidInteger(RcuNum, 1) Then
                    Grab_RCU_Num = RcuNum
                End If
            End If
        Next

    End Function

    'grab ip
    Function Grab_IP_addr(inputString As String) As String
        Dim LArray_ip() As String, R As String, k As Integer

        LArray_ip = inputString.Replace(" "c, "").Split("."c)
        R = ""

        For k = LBound(LArray_ip) To UBound(LArray_ip)
            R = R & CLng(CDbl(LArray_ip(k))).ToString("X2")
        Next k

        Grab_IP_addr = Strings.Right(R, 2) & " " & Strings.Mid(R, 5, 2) & " " & Strings.Mid(R, 3, 2) & " " & Strings.Left(R, 2)
    End Function

    ' grab var
    ' settings
    'added in v5.3 read write settings
    Function Grab_Var_4(inputString As String) As String 'settings
        Dim V As String, LArray_T1() As String, i As Integer

        V = "00000000"
        LArray_T1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If LArray_T1(i).Replace(" "c, "").Contains("var=") Then
                V = LArray_T1(i).Replace("var=", "")
                Exit For
            End If
        Next i

        If IsValidHexString(V) Then
            If Convert.ToInt64(V, 16) < 65536 Then
                V = V.PadLeft(8, "0"c)
            End If
        End If

        Grab_Var_4 = Strings.Right(V, 2) & " " & Strings.Mid(V, 5, 2) & " " & Strings.Mid(V, 3, 2) & " " & Strings.Left(V, 2)
    End Function

    ' grab elements
    Function Grab_element(inputString As String) As String
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
    Function Grab_Reg_1(inputString As String) As String
        Dim P As Integer, i As Integer, hex4 As String
        Dim LArray_T1() As String
        Dim LArray_GP() As String

        P = 0
        LArray_T1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)

        For i = LBound(LArray_T1) To UBound(LArray_T1)
            If LArray_T1(i).Replace(" "c, "").Contains("reg=") Then
                LArray_GP = LArray_T1(i).Replace(" "c, "").Split("="c)
                If IsNumeric(LArray_GP(1)) AndAlso LArray_GP(0) = "reg" Then
                    If CInt(LArray_GP(1)) > -1 AndAlso CInt(LArray_GP(1)) < 65536 Then
                        P = CInt(LArray_GP(1))
                        Exit For
                    End If
                End If
            End If
        Next i

        hex4 = P.ToString("X4")
        Return hex4.Substring(0, 2) & " " & hex4.Substring(2, 2)
    End Function

    ' grab dev type
    Function Grab_Dev_Type_1(inputString As String) As String
        Dim P As Integer, i As Integer
        Dim LArray_T1() As String
        Dim LArray_GP() As String

        P = 0
        LArray_T1 = inputString.Replace("#"c, " $").Split(New String() {" $"}, StringSplitOptions.None)

        For i = LBound(LArray_T1) To UBound(LArray_T1)

            If LArray_T1(i).Replace(" "c, "").Contains("devtype=") OrElse LArray_T1(i).Replace(" "c, "").Contains("type=") Then
                LArray_GP = LArray_T1(i).Replace(" "c, "").Split("="c)
                If IsNumeric(LArray_GP(1)) AndAlso (LArray_GP(0) = "devtype" OrElse LArray_GP(0) = "type") Then
                    If CInt(LArray_GP(1)) > 0 AndAlso CInt(LArray_GP(1)) < 11 Then
                        P = CInt(LArray_GP(1))
                        Exit For
                    End If
                ElseIf (LArray_GP(1) = "server" OrElse LArray_GP(1) = "db" OrElse LArray_GP(1) = "ff") AndAlso (LArray_GP(0) = "devtype" OrElse LArray_GP(0) = "type") Then
                    P = 255
                    Exit For
                ElseIf (Not IsNumeric(LArray_GP(1))) AndAlso (LArray_GP(0) = "devtype" OrElse LArray_GP(0) = "type") Then
                    P = Grab_Type_from_Name(LArray_GP(1))
                    Exit For
                End If
            End If
        Next i

        Return P.ToString("X2")
    End Function

    Private Function Grab_Type_from_Name(inputString As String) As Integer
        Dim RcuDevList() As String, RcuDevNameList() As String

        RcuDevList = "01,02,03,04,04,05,05,06,06,07,07,08,10".Split(","c)
        RcuDevNameList = "idpg,tig,tag,ioexp,ioe,iodexp,iod,gs,gsw,ioexp,ioe,bsp,dali".Split(","c)


        For i = 0 To RcuDevList.Length - 1
            If inputString.Contains(RcuDevNameList(i)) Then
                Return CInt(RcuDevList(i))
            End If
        Next i

        Return 0
    End Function


    'Grab_Time_1 > return time Ticks
    'Grab_Value_1 > retrun value input to functions
    'Grab_Dev_1 > return device for 03 03

    Function HasTime(inputString As String) As Boolean
        If inputString.Contains(" $hour") OrElse inputString.Contains(" $min") OrElse inputString.Contains(" $sec") Then
            Return True
        End If
        Return False
    End Function

    Function Grab_Time(inputString As String, Optional length As Byte = 4) As String
        Dim P As Double, Q As Double, R As Double, t As Double, s As String
        Dim LArray_T1() As String, i As Integer

        P = 0 : Q = 0 : R = 0
        LArray_T1 = inputString.Replace(" "c, "").Split("$"c)

        For i = 0 To LArray_T1.Length - 1
            If LArray_T1(i).StartsWith("hour=") Then
                P = GrabPara_Dbl(LArray_T1(i), "hour")
            ElseIf LArray_T1(i).StartsWith("min=") Then
                Q = GrabPara_Dbl(LArray_T1(i), "min")
            ElseIf LArray_T1(i).StartsWith("sec=") Then
                R = GrabPara_Dbl(LArray_T1(i), "sec")
            End If
        Next i

        t = ((P * 3600) + (Q * 60) + R) * 10000
        s = CUInt(t).ToString("X8")

        If length = 2 Then
            'time for dimmer load
            Return s.Substring(6, 2) & " " & s.Substring(4, 2)
        ElseIf length = 3 Then
            'time for RCU dimmer
            Return s.Substring(6, 2) & " " & s.Substring(4, 2) & " " & s.Substring(2, 2)
        Else
            'time for run wait
            Return s.Substring(6, 2) & " " & s.Substring(4, 2) & " " & s.Substring(2, 2) & " " & s.Substring(0, 2)
        End If

    End Function

    'time for run wait
    'Function Grab_Time_1(inputString As String) As String
    '    Dim P As Double, Q As Double, R As Double, t As Double, s As String
    '    Dim LArray_T1() As String, i As Integer

    '    P = 0 : Q = 0 : R = 0
    '    LArray_T1 = inputString.Replace(" "c, "").Split("$"c)

    '    For i = 0 To LArray_T1.Length - 1
    '        If Strings.InStr(LArray_T1(i), "hour=") = 1 Then
    '            P = GrabPara_Dbl(LArray_T1(i), "hour")
    '        ElseIf Strings.InStr(LArray_T1(i), "min=") = 1 Then
    '            Q = GrabPara_Dbl(LArray_T1(i), "min")
    '        ElseIf Strings.InStr(LArray_T1(i), "sec=") = 1 Then
    '            R = GrabPara_Dbl(LArray_T1(i), "sec")
    '        End If
    '    Next i

    '    t = ((P * 3600) + (Q * 60) + R) * 10000
    '    s = CLng(t).ToString("X8")

    '    Return s.Substring(6, 2) & " " & s.Substring(4, 2) & " " & s.Substring(2, 2) & " " & s.Substring(0, 2)
    'End Function

    'time for dimmer load
    'Private Function Grab_Time_2(inputString As String) As String
    '    Dim P As Double, Q As Double, R As Double, t As Double, s As String
    '    Dim LArray_T1() As String, i As Integer

    '    P = 0 : Q = 0 : R = 0
    '    LArray_T1 = inputString.Replace(" "c, "").Split("$"c)

    '    For i = 0 To LArray_T1.Length - 1
    '        If Strings.InStr(LArray_T1(i), "hour=") = 1 Then
    '            P = GrabPara_Dbl(LArray_T1(i), "hour")
    '        ElseIf Strings.InStr(LArray_T1(i), "min=") = 1 Then
    '            Q = GrabPara_Dbl(LArray_T1(i), "min")
    '        ElseIf Strings.InStr(LArray_T1(i), "sec=") = 1 Then
    '            R = GrabPara_Dbl(LArray_T1(i), "sec")
    '        End If
    '    Next i

    '    t = ((P * 3600) + (Q * 60) + R) * 100
    '    s = CLng(t).ToString("X8")

    '    Return s.Substring(6, 2) & " " & s.Substring(4, 2)
    'End Function

    'time for RCU dimmer
    'Function Grab_Time_3(inputString As String) As String
    '    Dim P As Double, Q As Double, R As Double, t As Double, s As String
    '    Dim LArray_T1() As String, i As Integer

    '    P = 0 : Q = 0 : R = 0
    '    LArray_T1 = inputString.Replace(" "c, "").Split("$"c)

    '    For i = 0 To LArray_T1.Length - 1
    '        If Strings.InStr(LArray_T1(i), "hour=") = 1 Then
    '            P = GrabPara_Dbl(LArray_T1(i), "hour")
    '        ElseIf Strings.InStr(LArray_T1(i), "min=") = 1 Then
    '            Q = GrabPara_Dbl(LArray_T1(i), "min")
    '        ElseIf Strings.InStr(LArray_T1(i), "sec=") = 1 Then
    '            R = GrabPara_Dbl(LArray_T1(i), "sec")
    '        End If
    '    Next i

    '    t = ((P * 3600) + (Q * 60) + R) * 10000
    '    s = CLng(t).ToString("X8")

    '    Return s.Substring(6, 2) & " " & s.Substring(4, 2) & " " & s.Substring(2, 2)
    'End Function

    ' level 0 grabbers
    Private Function GrabPara_BIN(inputString As String) As String
        Dim LArray_GP() As String

        LArray_GP = inputString.Replace(" "c, "").Split("="c)

        If IsValidBinString(LArray_GP(1)) AndAlso LArray_GP(0) = "bin" Then
            Return LArray_GP(1).PadLeft(32, "0"c)
        End If

        Return New String("0"c, 32)
    End Function

    Private Function GrabPara_HEX(inputString As String) As String
        Dim LArray_GP() As String

        LArray_GP = inputString.Replace(" "c, "").Split("="c)

        If IsValidHexString(LArray_GP(1)) AndAlso LArray_GP(0) = "hex" Then
            Return Strings.Right("00000000" & UCase(LArray_GP(1)), 8)
        Else
            Return "00000000"
        End If
    End Function

    'Private Function GrabPara_Lng(inputString As String, checkPara As String) As Long
    '    Dim LArray_GP() As String
    '    LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
    '    If IsNumeric(Evaluate(LArray_GP(1))) AndAlso LArray_GP(0) = checkPara Then
    '        If CLng(Evaluate(LArray_GP(1))) > 0 AndAlso CLng(Evaluate(LArray_GP(1))) < 4294967296.0# Then
    '            GrabPara_Lng = CLng(Evaluate(LArray_GP(1)))
    '        Else
    '            GrabPara_Lng = 0
    '        End If
    '    Else
    '        GrabPara_Lng = 0
    '    End If
    'End Function

    Private Function GrabPara_Int(inputString As String, checkPara As String) As UInteger
        Dim LArray_GP() As String
        LArray_GP = inputString.Replace(" "c, "").Split("="c)

        If LArray_GP.Length <> 2 OrElse LArray_GP(0) <> checkPara Then
            Return 0
        End If

        Try
            Dim val As Double = Evaluate(LArray_GP(1))
            If val > 0 AndAlso val < 4294967296.0# Then
                Return Convert.ToInt64(val)
            End If
        Catch ex As Exception
            'to report error later
        End Try

        Return 0

    End Function

    Private Function GrabPara_Dbl(inputString As String, checkPara As String) As Double
        Dim LArray_GP() As String
        LArray_GP = inputString.Replace(" "c, "").Split("="c)

        If LArray_GP.Length <> 2 OrElse LArray_GP(0) <> checkPara Then
            Return 0
        End If

        Try
            Dim val As Double = Evaluate(LArray_GP(1))
            If val > 0 AndAlso val < 4294967296.0# Then
                Return val
            End If
        Catch ex As Exception
            'to report error later
        End Try

        Return 0
    End Function

    'Private Function GrabPara_BitPOS_OR(inputString As String, checkPara As String) As String
    '    Dim LArray_GP() As String, RelayBits() As String, i As Integer
    '    LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
    '    RelayBits = Strings.Split("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
    '    GrabPara_BitPOS_OR = ""
    '    If LArray_GP(0) = checkPara Then
    '        Dim Sbits() As String
    '        LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
    '        Sbits = Strings.Split(LArray_GP(1), ",")

    '        For i = LBound(Sbits) To UBound(Sbits)
    '            If IsNumeric(Sbits(i)) Then
    '                If CInt(Sbits(i)) > 0 AndAlso CInt(Sbits(i)) < 33 Then
    '                    RelayBits(CInt(Sbits(i)) - 1) = "1"
    '                End If
    '            End If
    '        Next i
    '    End If
    '    For i = LBound(RelayBits) To UBound(RelayBits)
    '        GrabPara_BitPOS_OR = RelayBits(i) & GrabPara_BitPOS_OR
    '    Next i
    'End Function

    'Private Function GrabPara_BitPOS_AND(inputString As String, checkPara As String) As String
    '    Dim LArray_GP() As String, RelayBits() As String, i As Integer
    '    LArray_GP = Strings.Split(Strings.Replace(inputString, " ", ""), "=")
    '    RelayBits = Strings.Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1", ",")
    '    GrabPara_BitPOS_AND = ""
    '    If LArray_GP(0) = checkPara Then
    '        Dim Sbits() As String
    '        LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
    '        Sbits = Strings.Split(LArray_GP(1), ",")

    '        For i = LBound(Sbits) To UBound(Sbits)
    '            If IsNumeric(Sbits(i)) Then
    '                If CInt(Sbits(i)) > 0 AndAlso CInt(Sbits(i)) < 33 Then
    '                    RelayBits(CInt(Sbits(i)) - 1) = "0"
    '                End If
    '            End If
    '        Next i
    '    End If
    '    For i = LBound(RelayBits) To UBound(RelayBits)
    '        GrabPara_BitPOS_AND = RelayBits(i) & GrabPara_BitPOS_AND
    '    Next i
    'End Function

    Private Function GrabPara_BitPOS_OR(inputString As String, checkPara As String) As String
        Dim LArray_GP() As String, i As Integer
        Dim OutputString As New StringBuilder("00000000000000000000000000000000")

        LArray_GP = inputString.Replace(" "c, "").Split("="c)

        If LArray_GP(0) = checkPara Then
            Dim Sbits() As String
            LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
            Sbits = LArray_GP(1).Split(","c)

            For i = 0 To Sbits.Length - 1
                If IsNumeric(Sbits(i)) Then
                    If CInt(Sbits(i)) > 0 AndAlso CInt(Sbits(i)) < 33 Then
                        OutputString(31 - CInt(Sbits(i)) + 1) = "1"
                    End If
                End If
            Next i
        End If

        Return OutputString.ToString
    End Function

    Private Function GrabPara_BitPOS_AND(inputString As String, checkPara As String) As String
        Dim LArray_GP() As String, i As Integer
        Dim OutputString As New StringBuilder("11111111111111111111111111111111")

        LArray_GP = inputString.Replace(" "c, "").Split("="c)
        If LArray_GP(0) = checkPara Then
            Dim Sbits() As String
            LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
            Sbits = LArray_GP(1).Split(","c)

            For i = 0 To Sbits.Length - 1
                If IsNumeric(Sbits(i)) Then
                    If CInt(Sbits(i)) > 0 AndAlso CInt(Sbits(i)) < 33 Then
                        OutputString(31 - CInt(Sbits(i)) + 1) = "0"
                    End If
                End If
            Next i
        End If

        Return OutputString.ToString
    End Function


    Private Function GrabPara_FAN(inputString As String) As String
        Dim LArray_GP() As String
        LArray_GP = inputString.Replace(" "c, "").Split("="c)
        If LArray_GP(1) = "fan1" AndAlso LArray_GP(0) = "fan" Then
            Return "003F0655"
        ElseIf LArray_GP(1) = "fan2" AndAlso LArray_GP(0) = "fan" Then
            Return "003F0565"
        ElseIf LArray_GP(1) = "fan3" AndAlso LArray_GP(0) = "fan" Then
            Return "003F0556"
        ElseIf (LArray_GP(1) = "auto" OrElse LArray_GP(1) = "fan4") AndAlso LArray_GP(0) = "fan" Then
            Return "003F0955"
        ElseIf (LArray_GP(1) = "eco" OrElse LArray_GP(1) = "fan5") AndAlso LArray_GP(0) = "fan" Then
            Return "003F0595"
        ElseIf (LArray_GP(1) = "off" OrElse LArray_GP(1) = "fan6") AndAlso LArray_GP(0) = "fan" Then
            Return "003F0559"
        Else
            Return "00000000"
        End If
    End Function

    Private Function GrabPara_Display(inputString As String, Optional lengthByte As Integer = 4) As String
        Dim LArray_GP() As String, thischar As String, nextchar As String
        LArray_GP = inputString.Replace(" "c, "").Split("="c)

        If LArray_GP(0) = "display" Then

            Dim DisplayChar() As String, DisplayHexChar() As String, CharSpace() As String, HexCharSpace() As String, rPoint As Integer
            DisplayChar = "a,b,/c,\c,c,d,/e,e,f,g,/h,h,i,j,l,/n,n,/o,\o,o,p,q,r,s,t,/u,\u,u,y,0,1,2,3,4,5,6,7,8,9,.,-,_, ".Split(","c)
            DisplayHexChar = "77,1F,0E,43,4D,3E,6F,4F,47,5D,17,37,10,38,0D,16,75,1E,63,7D,67,73,06,5B,0F,1C,23,3D,3B,7D,30,6E,7A,33,5B,5F,70,7F,7B,80,02,08,00".Split(","c)
            CharSpace = ",,,".Split(","c)
            HexCharSpace = "00,00,00,00".Split(","c)

            If LArray_GP(1).StartsWith("["c) Then
                LArray_GP(1) = LArray_GP(1).Substring(1)
            End If
            If LArray_GP(1).EndsWith("]"c) Then
                LArray_GP(1) = LArray_GP(1).Substring(0, LArray_GP(1).Length - 1)
            End If
            LArray_GP(1) = "        " & LArray_GP(1) '8 spaces
            rPoint = -1

            For i = 0 To Strings.Len(LArray_GP(1)) - 4
                thischar = Strings.Mid(LArray_GP(1), Strings.Len(LArray_GP(1)) - i, 1)
                nextchar = Strings.Mid(LArray_GP(1), Strings.Len(LArray_GP(1)) - 1 - i, 1)

                If thischar <> "/" AndAlso thischar <> "\" Then
                    If thischar <> "." Then
                        rPoint = rPoint + 1
                        If rPoint < 4 Then
                            CharSpace(rPoint) = thischar & CharSpace(rPoint)
                        End If
                    Else
                        If thischar = "." AndAlso nextchar <> "." Then
                            If rPoint < 3 Then
                                CharSpace(rPoint + 1) = thischar & CharSpace(rPoint + 1)
                            End If
                        ElseIf thischar = "." AndAlso nextchar = "." Then
                            rPoint = rPoint + 1
                            If rPoint < 4 Then
                                CharSpace(rPoint) = thischar & CharSpace(rPoint)
                            End If
                        End If
                    End If
                Else
                    If rPoint > -1 AndAlso rPoint < 4 Then
                        CharSpace(rPoint) = thischar & CharSpace(rPoint)
                    End If
                End If

            Next i

            For i = 0 To 3
                If CharSpace(i).EndsWith("."c) Then
                    HexCharSpace(i) = "80"
                    CharSpace(i) = Strings.Left(CharSpace(i), Strings.Len(CharSpace(i)) - 1)
                End If

                For j = LBound(DisplayChar) To UBound(DisplayChar)
                    If CharSpace(i) = DisplayChar(j) Then
                        HexCharSpace(i) = Hexadder(HexCharSpace(i), DisplayHexChar(j))
                    End If
                Next j
            Next i

            If lengthByte = 1 Then
                Return HexCharSpace(0)
            ElseIf lengthByte = 2 Then
                Return HexCharSpace(0) & " " & HexCharSpace(1)
            Else
                Return HexCharSpace(0) & " " & HexCharSpace(1) & " " & HexCharSpace(2) & " " & HexCharSpace(3)
            End If

        Else
            If lengthByte = 1 Then
                Return "00"
            ElseIf lengthByte = 2 Then
                Return "00 00"
            Else
                Return "00 00 00 00"
            End If
        End If
    End Function

    Private Function GrabPara_Clock(inputString As String) As String 'added from ver 5.2 for new data type $clock
        Dim LArray_GP() As String, Year As Integer, Month As Integer, Day As Integer, Hour As Integer, Minute As Integer, Second As Integer
        Dim DArray() As String, TArray() As String, BinClock As String

        LArray_GP = inputString.Replace(" "c, "").Split("="c)
        If LArray_GP(0) = "clock" Then
            If LArray_GP(1).IndexOf("-"c) > 0 Then
                'means you have date and time
                If IsValidDate(Strings.Left(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") - 1)) AndAlso IsValidTime(Strings.Mid(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") + 1)) Then
                    DArray = Strings.Split(Strings.Left(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") - 1), "/"c)
                    TArray = Strings.Split(Strings.Mid(LArray_GP(1), Strings.InStr(LArray_GP(1), "-") + 1), ":"c)
                    Year = CInt(Strings.Right("00" & DArray(0), 2))
                    Month = CInt(DArray(1))
                    Day = CInt(DArray(2))
                    Hour = CInt(TArray(0))
                    Minute = CInt(TArray(1))
                    Second = CInt(TArray(2))

                    BinClock = Convert.ToString(Year, 2).PadLeft(6, "0"c) & Convert.ToString(Month, 2).PadLeft(4, "0"c) & Convert.ToString(Day, 2).PadLeft(5, "0"c) &
                         Convert.ToString(Hour, 2).PadLeft(5, "0"c) & Convert.ToString(Minute, 2).PadLeft(6, "0"c) & Convert.ToString(Second, 2).PadLeft(6, "0"c)
                    Return Convert.ToUInt32(BinClock, 2).ToString("X8")
                End If
            Else
                'can be date only
                If IsValidDate(LArray_GP(1)) Then
                    DArray = LArray_GP(1).Split("/"c)
                    Year = CInt(Strings.Right("00" & DArray(0), 2))
                    Month = CInt(DArray(1))
                    Day = CInt(DArray(2))

                    BinClock = Convert.ToString(Year, 2).PadLeft(6, "0"c) & Convert.ToString(Month, 2).PadLeft(4, "0"c) & Convert.ToString(Day, 2).PadLeft(5, "0"c) &
                         Convert.ToString(Hour, 2).PadLeft(5, "0"c) & Convert.ToString(Minute, 2).PadLeft(6, "0"c) & Convert.ToString(Second, 2).PadLeft(6, "0"c)
                    Return Convert.ToUInt32(BinClock, 2).ToString("X8")
                End If
            End If
        End If

        Return "00000000"
    End Function


    Function Grab_name(inputString As String) As String
        Dim LArray_v1() As String, i As Integer, output As String
        LArray_v1 = inputString.Replace("#"c, "$"c).Replace(" "c, "").Split("$"c)

        For i = 0 To LArray_v1.Length - 1
            If LArray_v1(i).StartsWith("name=") Then
                output = LArray_v1(i).Replace("name=", "")
                If output.Length < 1 Then
                    Return "NULL"
                Else
                    Return output
                End If

            End If
        Next

        Return "NULL"
    End Function

End Class
