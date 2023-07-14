VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XcodeX1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim X3 As New XcodeX3
Dim ErrorLog As String, WarnLog As String

Public Function getfilename(pp As Application) As String
   getfilename = Replace(Replace(Replace(pp.ActiveWorkbook.Name, ".xlsm", ""), ".xlsx", ""), ".XLSX", "")
End Function

Function getfilepath(pp As Application, Optional toLower As Boolean = True) As String 'added in XCODE 4.9.2, redirect to onedrive local path
    Dim tempPath As String

    tempPath = Replace(pp.ActiveWorkbook.FullName, pp.ActiveWorkbook.Name, "")
    
    If toLower Then
        tempPath = LCase(tempPath)
    End If
    
    If InStr(tempPath, "http") = 1 Then
        Dim onedrivePath As String
        onedrivePath = Environ("onedrive")
        getfilepath = Replace(onedrivePath & Mid(pp.ActiveWorkbook.Path, InStr(pp.ActiveWorkbook.Path, "Documents") + 9), "/", "\")
    Else
        getfilepath = tempPath
    End If
End Function

Public Function getFilePathFormat(pp As Application) As Integer 'added in XCODE 4.9.2, detect if local path accessible
    '0 other not accessible, 1 as local path, 2 as onedrive path,
    'removed letter c checking from path. in XCODE ver 5.2
    Dim tempPath As String
    tempPath = LCase(Replace(pp.ActiveWorkbook.FullName, pp.ActiveWorkbook.Name, ""))
    If InStr(tempPath, ":\") = 2 Then
        getFilePathFormat = 1
    ElseIf InStr(tempPath, "http") = 1 Then
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

Public Function getsufixname(pp As Application) As String
    getsufixname = Replace(Replace(Replace(Replace(pp.ActiveWorkbook.Name, "RCU mem map.", ""), ".xlsm", ""), ".xlsx", ""), ".XLSX", "")
End Function

' sheet check
Public Function checkSheet(n As String, pp As Application) As Boolean
  checkSheet = False
  For Each ws In pp.Worksheets
    If n = ws.Name Then
      checkSheet = True
      Exit Function
    End If
  Next ws
End Function

'bin checker
Public Function isValidBinString(strNumber As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .IgnoreCase = True
        .Global = True
        .Pattern = "[2-9a-z\W]+"
        isValidBinString = Not .test(strNumber)
    End With
End Function

Public Function isBitTrue(inputBin As String, inputPos As Integer) As Boolean ' not required to validate since it has validated inputs
    isBitTrue = False
    
    If Mid(inputBin, Len(inputBin) - inputPos + 1, 1) = "1" Then
        isBitTrue = True
    End If
End Function

'hex checker

Public Function isValidHexString(strNumber As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .IgnoreCase = True
        .Global = True
        .Pattern = "[g-z\W]+"
        isValidHexString = Not .test(strNumber)
    End With
End Function

'date validation, added from ver 5.2
Private Function isValidDate(inputString As String) As Boolean
    Dim DArr() As String, Y1 As Integer, M1 As Integer, D1 As Integer
    isValidDate = False
    If InStr(inputString, "/") > 2 Then
        DArr = Split(inputString, "/")
        If UBound(DArr) = 2 Then
            If IsNumeric(DArr(0)) And IsNumeric(DArr(1)) And IsNumeric(DArr(2)) Then
                Y1 = Int(Right("00" & DArr(0), 2))
                M1 = Int(DArr(1))
                D1 = Int(DArr(2))
                
                If Y1 > 10 And Y1 < 100 And M1 > 0 And M1 < 13 And D1 > 0 And D1 < 32 Then
                    isValidDate = True
                End If
            End If
        End If
    End If
End Function

'time validation, added from ver 5.2
Private Function isValidTime(inputString As String) As Boolean
    Dim TArr() As String, H1 As Integer, M1 As Integer, S1 As Integer
    isValidTime = False
    If InStr(inputString, ":") > 2 Then
        TArr = Split(inputString, ":")
        If UBound(TArr) = 2 Then
            If IsNumeric(TArr(0)) And IsNumeric(TArr(1)) And IsNumeric(TArr(2)) Then
                H1 = Int(TArr(0))
                M1 = Int(TArr(1))
                S1 = Int(TArr(2))
                
                If H1 < 24 And M1 < 60 And S1 < 60 Then
                    isValidTime = True
                End If
            End If
        End If
    End If
End Function


'ip checker
Public Function isAValidIP(inputip As String) As Boolean
    isAValidIP = False
    
    If InStr(inputip, ".") > 0 Then
    
        Dim LArray_ip() As String
        LArray_ip = Split(inputip, ".")
        
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

'var convert
Public Function varShort2LongFull(inputvar As String) As String
    Longvar = "e0084"
    Dim varInt As Integer
    varInt = WorksheetFunction.Hex2Dec(inputvar)
    Longvar = Longvar & WorksheetFunction.Dec2Hex(varInt * 4, 3)
    varShort2LongFull = Left(Longvar, 4) & " " & Right(Longvar, 4)
End Function

'crc16
Public Function modbusCRC16(inputString As String) As String
    Dim mask, i, j, nC, Crc As Integer
    Crc = &HFFFF
    For nC = 1 To Len(inputString) Step 2
        j = Val("&H" + Mid(inputString, nC, 2))
        Crc = Crc Xor j
        For j = 1 To 8
            mask = 0
            If Crc / 2 <> Int(Crc / 2) Then mask = &HA001
            Crc = Int(Crc / 2) And &H7FFF: Crc = Crc Xor mask
        Next j
    Next nC
    modbusCRC16 = Hex$(Crc)
End Function

'save files
' def_name for new name other than name generated from file ver 5.0
Private Function DeviceXF(pp As Application, rcuNum As String, Def_Name As String) As String
    Dim ce As Range, thisRCU As String, thisName As String
    Dim withCBS As Boolean, CBSline As String, CBSint As Integer
    
    withCBS = False
    
    If rcuNum = "" Then
        thisRCU = ""
    Else
        thisRCU = "-RCU" & rcuNum
    End If
    
    If Def_Name = "NULL" Then
        thisName = getsufixname(pp) & thisRCU
    Else
        thisName = Def_Name
    End If
    
    Open getfilepath(pp) & "\Devices." & thisName & ".hex" For Output As #1
    For Each ce In pp.Range("Device" & rcuNum & "!B4:B" & getDevLimit(pp, rcuNum))
        If ce.Text <> "" Then
            Print #1, ce.Cells(1, 1) & vbTab & ce.Cells(1, 2) & vbTab & ce.Cells(1, 3) & vbTab & ce.Cells(1, 4) & vbTab & ce.Cells(1, 5) & vbTab & ce.Cells(1, 6)
            
            If LCase(pp.WorksheetFunction.Trim(ce.Cells(1, 2))) = "0e" And (Not withCBS) Then
                withCBS = True
                CBSline = Replace(pp.WorksheetFunction.Trim(ce.Cells(1, 5)), " ", "")
                CBSline = Right(CBSline, 2) & Left(CBSline, 2)
                CBSint = pp.WorksheetFunction.Hex2Dec(CBSline) + 4
            End If
        End If
    Next ce
        
    If withCBS Then
        For Each ce In pp.Range("Device" & rcuNum & "!J4:J" & (CBSint - 1))
            If ce.Text <> "" Then
                Print #1, ce.Text
            End If
        Next ce
        For Each ce In pp.Range(pp.Range("Device" & rcuNum & "!G" & CBSint), pp.Range("Device" & rcuNum & "!G" & CBSint).End(xlDown))
            If ce.Text <> "" Then
                Print #1, ce.Cells(1, 1) & vbTab & ce.Cells(1, 2) & vbTab & ce.Cells(1, 3) & vbTab & ce.Cells(1, 4)
            End If
        Next ce
    Else
        For Each ce In pp.Range(pp.Range("Device" & rcuNum & "!J4"), pp.Range("Device" & rcuNum & "!J4").End(xlDown))
            If ce.Text <> "" Then
                Print #1, ce.Text
            End If
        Next ce
    End If
    
    Close #1
    
    DeviceXF = "Devices." & thisName & ".hex created"
End Function
  
Private Function ProgramXF(pp As Application, rcuNum As String, Def_Name As String) As String
    Dim ce As Range, thisRCU As String, thisName As String
    
    If rcuNum = "" Then
        thisRCU = ""
    Else
        thisRCU = "-RCU" & rcuNum
    End If
    
    If Def_Name = "NULL" Then
        thisName = getsufixname(pp) & thisRCU
    Else
        thisName = Def_Name
    End If
    
    Open getfilepath(pp) & "\Program." & thisName & ".hex" For Output As #1
    
    For Each ce In pp.Range(pp.Range("Program" & rcuNum & "!E1"), pp.Range("Program" & rcuNum & "!E1").End(xlDown))
        If ce.Text <> "" Then
            Print #1, ce.Text
        End If
    Next ce
    
    Close #1
    
    ProgramXF = "Program." & thisName & ".hex created"
End Function

Private Function ParamsXF(pp As Application, rcuNum As String, Def_Name As String) As String
    Dim ce As Range, thisRCU As String, thisName As String
    
    If rcuNum = "" Then
        thisRCU = ""
    Else
        thisRCU = "-RCU" & rcuNum
    End If
    
    If Def_Name = "NULL" Then
        thisName = getsufixname(pp) & thisRCU
    Else
        thisName = Def_Name
    End If
    
    Open getfilepath(pp) & "\Params." & thisName & ".hex" For Output As #1
    
    For Each ce In pp.Range(pp.Range("Settings" & rcuNum & "!B1"), pp.Range("Settings" & rcuNum & "!B1").End(xlDown))
        If ce.Text <> "" Then
            Print #1, ce.Text
        End If
    Next ce
    
    Close #1
    
    ParamsXF = "Params." & thisName & ".hex created"
End Function

Private Function getDefinedName(pp As Application, thisRCU As String) As String 'added in ver 5.0
    Dim ce As Range, ThisTempLine As String
    getDefinedName = "NULL"
    
    For Each ce In pp.Range(pp.Range("XCODE" & thisRCU & "!B2"), pp.Range("XCODE" & thisRCU & "!B2").End(xlDown))
        
        ThisTempLine = LCase(pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text))
                    
        If InStr(ThisTempLine, "//") > 2 And Not Left(ThisTempLine, 2) = "0x" Then 'comment detection and filter except direct hex
            ThisTempLine = pp.WorksheetFunction.Trim(Left(ThisTempLine, InStr(ThisTempLine, "//") - 1))
        End If
        
        If InStr(ThisTempLine, "define rcu ") = 1 Then
            getDefinedName = Grab_name(ThisTempLine)
            Exit Function
        End If
        
    Next ce
    
End Function

Public Sub SaveAllXcode(pp As Application)
    Dim thisRCUList As Variant, Notification As String, thisRCU As String, evalPath As Integer
    Dim Defined_Name As String, rcount As Integer, rfullcount As Integer
    
    thisRCUList = RCUList(pp)
    Notification = "": thisRCU = "": Defined_Name = "NULL"
    evalPath = getFilePathFormat(pp)
    
    If evalPath > 0 Then
    
        If evalPath = 2 Then
            Notification = "OneDrive synchronise pending" & vbNewLine & vbNewLine
        End If
        
        rcount = 0
        rfullcount = UBound(thisRCUList) - LBound(thisRCUList) + 1
        
        For L = LBound(thisRCUList) To UBound(thisRCUList)
        
            thisRCU = CStr(thisRCUList(L))
            Defined_Name = getDefinedName(pp, thisRCU)
            
            If checkSheet("Device" & thisRCU, pp) Then
                Notification = Notification & DeviceXF(pp, thisRCU, Defined_Name) & vbNewLine
            Else
                Notification = Notification & "Device" & thisRCU & " sheet not found in this mem map" & vbNewLine & "please rename sheet to Device" & thisRCU & vbNewLine
            End If
            
            If checkSheet("Program" & thisRCU, pp) Then
                Notification = Notification & ProgramXF(pp, thisRCU, Defined_Name) & vbNewLine
            Else
                Notification = Notification & "Program" & thisRCU & " sheet not found in this mem map" & vbNewLine & "please rename sheet to Program" & thisRCU & vbNewLine
            End If
            
            If checkSheet("Settings" & thisRCU, pp) Then
                Notification = Notification & ParamsXF(pp, thisRCU, Defined_Name) & vbNewLine
            Else
                Notification = Notification & "Settings" & thisRCU & " sheet not found in this mem map" & vbNewLine & "please rename sheet to Settings" & thisRCU & vbNewLine
            End If
            
            rcount = rcount + 1
            pp.StatusBar = "XCODE | Printer | Collecting pages  " & String(((rcount * 100 / rfullcount) * 0.25), ChrW(9607)) & String(25 - ((rcount * 100 / rfullcount) * 0.25), ChrW(9620)) & " Paging " & CInt(rcount * 100 / rfullcount) & "% - " & getfilename(pp)
                    
        Next L
        pp.StatusBar = "XCODE | Printer |  Writing files to [" & getfilepath(pp, False) & "]"
        MsgBox Notification, vbInformation, "XCODE | Printer"
        
    Else
    
        Notification = "Destination path could not access"
        pp.StatusBar = "XCODE | Printer |  Destination path could not access"
        MsgBox Notification, vbCritical, "XCODE | Printer"
        
    End If
    

End Sub

'progbase check XCODE ver 4.8
Public Sub checkProgbase(pp As Application)
    Dim thisRCUList As Variant, thisRCU As String, withProgbase As Boolean, pfullcount As Integer, parambase As Boolean, parambval As Integer
    thisRCUList = RCUList(pp)
    thisRCU = "": withProgbase = False: parambase = False
    
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        thisRCU = CStr(thisRCUList(L))
        pfullcount = pp.Range("Program" & thisRCU & "!E4").End(xlDown).Row
        pp.StatusBar = "XCODE | Formatter |  RCU" & thisRCU & " Checking ProgBase setting"
        If pfullcount > 2002 Then
            withProgbase = True
        End If
        
        parambval = pp.WorksheetFunction.Hex2Dec(Replace(pp.Range("Settings" & thisRCU & "!B26").Value, ";", ""))
        
        If parambval = 0 Then
            parambase = True
        End If
        
    Next L
    
    If withProgbase Then
        For L = LBound(thisRCUList) To UBound(thisRCUList)
            thisRCU = CStr(thisRCUList(L))
            pp.Range("Settings" & thisRCU & "!B26").FormulaR1C1 = "=DEC2HEX(RC[1],2)&"";"""
            pp.Range("Settings" & thisRCU & "!C26").FormulaR1C1 = "0"
        Next L
        pp.StatusBar = "XCODE | Formatter |  Params updated with higher order settings"
        MsgBox getfilename(pp) & " exceeds standard size" _
        & vbNewLine & "Params updated with higher order settings" _
        & vbNewLine & vbNewLine & "Make sure to use RCUmanager enabled with Progbase or" & vbNewLine & "Newer version of it that manages Progbase automatically", vbInformation, "XCODE | Formatter"
    Else
        If parambase Then
            MSG1 = MsgBox("Noticed Progbase still enabled in params" & vbNewLine & "Even though the program is in standard size" & vbNewLine & "Do you still need to enable Progbase for " & getfilename(pp) & "?", vbYesNo + vbQuestion, "Enable Progbase?")
            If MSG1 = vbYes Then
                For L = LBound(thisRCUList) To UBound(thisRCUList)
                    thisRCU = CStr(thisRCUList(L))
                    pp.Range("Settings" & thisRCU & "!B26").FormulaR1C1 = "=DEC2HEX(RC[1],2)&"";"""
                    pp.Range("Settings" & thisRCU & "!C26").FormulaR1C1 = "0"
                Next L
                pp.StatusBar = "XCODE | Formatter |  " & getfilename(pp) & " is in standard size, but Params updated with higher order settings as you requested"
                MsgBox getfilename(pp) & " is in standard size" _
                & vbNewLine & "But Params updated with higher order settings as you requested" _
                & vbNewLine & vbNewLine & "Make sure to use RCUmanager enabled with Progbase", vbInformation, "XCODE | Formatter"
            Else
                For L = LBound(thisRCUList) To UBound(thisRCUList)
                    thisRCU = CStr(thisRCUList(L))
                    pp.Range("Settings" & thisRCU & "!B26").FormulaR1C1 = "=DEC2HEX(RC[1],2)&"";"""
                    pp.Range("Settings" & thisRCU & "!C26").FormulaR1C1 = "128"
                Next L
                pp.StatusBar = "XCODE | Formatter |  " & getfilename(pp) & " is in standard size, Params updated with lower order settings"
                MsgBox getfilename(pp) & " is in standard size" _
                & vbNewLine & "Params updated with lower order settings" _
                & vbNewLine & vbNewLine & "Make sure to use RCUmanager without Progbase", vbInformation, "XCODE | Formatter"
            End If
        End If
    End If
    
End Sub


'variable check moved in ver 4.9

Public Function hasValidVar(inputString As String) As Boolean 'RCU var
    If InStr(LCase(inputString), " $var") > 1 Then
        Dim RcuVariableList() As String, RcuNameList() As String
        RcuVariableList = Split("04,05,06,07,08,09,0A,0B,0C,0D,0E,0F,10,11,12,13,14,14,33,15,16,17,18,19,1A,1B,1C,1D,1E,1F,20,21,22,23,24,25,26,27", ",")
        RcuNameList = Split("anti_ice,checkin,thermo1,thermo2,thermo3,thermo4,thermo5,thermo6,temp1,temp2,temp3,temp4,temp5,temp6,season,intervention,room_empty,room_empty_temp,room_empty_thermo,user,workflow,gs1,gs2,gs3,gs4,gs5,gs6,gs7,gs8,gs9,gs10,ioexp1,ioexp2,ioexp3,ioexp4,ioexp5,ioexp6,tag_minmax", ",")
        Dim LArray_T2() As String, V As String
        V = ""
        LArray_T2 = Split(Replace(inputString, "#", " $"), " $")
        For i = LBound(LArray_T2) To UBound(LArray_T2)
            If InStr(Replace(LArray_T2(i), " ", ""), "var=") Then
                V = Replace(LArray_T2(i), "var=", "")
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
        Else
            hasValidVar = False
        End If
    Else
        hasValidVar = False
    End If
End Function


Public Function hasValidC_Var(inputString As String, pp As Application, rr As String) As Boolean 'Custom var, moified in ver 5.0
    If C_VarListNotEmpty(pp, rr) Then
    'If checkSheet("CustomVar" & rr, pp) Then
        'If XcheckCustomVarList(pp, rr) Then
            Dim LArray_T2() As String, V As String
            V = ""
            LArray_T2 = Split(Replace(inputString, "#", " $"), " $")
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If InStr(Replace(LArray_T2(i), " ", ""), "var=") Then
                    V = Replace(LArray_T2(i), "var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
               Dim rngVar As Range
               For Each rngVar In pp.Range(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0), pp.Range("CustomVar" & rr & "!B3").End(xlDown))
                    If V = pp.WorksheetFunction.Trim(LCase(rngVar.Text)) Or V = pp.WorksheetFunction.Trim(LCase(rngVar.Offset(0, 1).Text)) Then
                        hasValidC_Var = True
                        Exit Function
                    End If
               Next rngVar
            Else
               hasValidC_Var = False
            End If
        'Else
            'hasValidC_Var = False
        'End If
    Else
        hasValidC_Var = False
    End If
End Function

Public Function C_VarListNotEmpty(pp As Application, rr As String) As Boolean 'ver 5.0
    C_VarListNotEmpty = False
    If checkSheet("CustomVar" & rr, pp) Then
        If Len(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0).Text) <> 0 Then
            C_VarListNotEmpty = True
        End If
    End If
End Function

Public Function NoErrorCustomVarList(pp As Application) As Boolean ' ver 5.0 if true no errors, if false with critical errors and erros saved to error log
    NoErrorCustomVarList = True
    Dim thisRCUList As Variant, thisRCU As String
    Dim ce As Range, C_VarShort As String, C_VarLong As String, outPutError As String
    thisRCUList = RCUList(pp)
    thisRCU = "": outPutError = ""
    
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        thisRCU = CStr(thisRCUList(L))
        C_VarShort = "": C_VarLong = ","
        
        If C_VarListNotEmpty(pp, thisRCU) Then
            For Each ce In pp.Range(pp.Range("CustomVar" & thisRCU & "!B3").Offset(1, 0), pp.Range("CustomVar" & thisRCU & "!B3").End(xlDown))
                    
                If Not isValidHexString(ce.Text) Then
                    outPutError = outPutError & "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Non-Hexadecimal value as variable short name]" & vbNewLine
                End If
                
                If InStr(pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                    outPutError = outPutError & "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Custom Variable definition with spaces]" & vbNewLine
                End If
                
                If InStr(C_VarShort, Right("00" & LCase(ce.Text), 2) & ",") > 0 Then
                    outPutError = outPutError & "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Duplicate variable]" & vbNewLine
                Else
                    C_VarShort = C_VarShort & Right("00" & LCase(ce.Text), 2) & ","
                End If
                
                If Len(ce.Offset(0, 1).Text) > 0 Then
                    If InStr(C_VarLong, "," & LCase(ce.Offset(0, 1).Text) & ",") > 0 Then
                        outPutError = outPutError & "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Duplicate variable definition]" & vbNewLine
                    Else
                        C_VarLong = C_VarLong & LCase(ce.Offset(0, 1).Text) & ","
                    End If
                Else
                    outPutError = outPutError & "[RCU" & thisRCU & " | CustomVar" & thisRCU & " | " & ce.Text & " ] [Empty variable definition]" & vbNewLine
                End If
            Next ce
        End If
    
    Next L
    
    If outPutError <> "" Then
        addError "Critical Error" & vbNewLine & outPutError
        NoErrorCustomVarList = False
    End If
End Function

'Function XcheckCustomVarList(pp As Application, rr As String) As Boolean 'replaced in ver 5.0 with C_VarListNotEmpty() and NoErrorCustomVarList()
    'XcheckCustomVarList = True
    'Dim ce As Range
    'If Len(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0).Text) <> 0 Then
        'For Each ce In pp.Range(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0), pp.Range("CustomVar" & rr & "!B3").End(xlDown))
                'If Not isValidHexString(ce.Text) Then
                    'XcheckCustomVarList = False
                    'MsgBox "Found '" & ce.Text & "' is not valid Custom Variable inside Table" _
                    '& vbNewLine & "Please use Hexadecimal variables"
                    'Exit Function
                'End If
                'If InStr(pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                    'XcheckCustomVarList = False
                    'MsgBox "'" & ce.Offset(0, 1).Text & "' is not valid. " _
                    '& vbNewLine & "Make Custom Variable definition without spaces"
                    'Exit Function
                'End If
        'Next ce
    'Else
        'XcheckCustomVarList = False
    'End If
'End Function


Public Function hasValidM_Var(inputString As String, pp As Application, rr As String) As Boolean 'Modbus var, moified in ver 5.0
    If M_VarListNotEmpty(pp, rr) Then
    'If checkSheet("ModbusVar" & rr, pp) Then
        'If XcheckModbusVarList(pp, rr) Then
            Dim LArray_T2() As String, V As String
            V = ""
            LArray_T2 = Split(Replace(inputString, "#", " $"), " $")
            For i = LBound(LArray_T2) To UBound(LArray_T2)
                If InStr(Replace(LArray_T2(i), " ", ""), "var=") Then
                    V = Replace(LArray_T2(i), "var=", "")
                    Exit For
                End If
            Next i
            If V <> "" Then
               Dim rngVar As Range
               For Each rngVar In pp.Range(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0), pp.Range("ModbusVar" & rr & "!B3").End(xlDown))
                    If V = pp.WorksheetFunction.Trim(LCase(rngVar.Text)) Or V = pp.WorksheetFunction.Trim(LCase(rngVar.Offset(0, 1).Text)) Then
                        hasValidM_Var = True
                        Exit Function
                    End If
               Next rngVar
            Else
               hasValidM_Var = False
            End If
        'Else
            'hasValidM_Var = False
        'End If
    Else
        hasValidM_Var = False
    End If
End Function

Public Function M_VarListNotEmpty(pp As Application, rr As String) As Boolean 'ver 5.0
    M_VarListNotEmpty = False
    If checkSheet("ModbusVar" & rr, pp) Then
        If Len(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0).Text) <> 0 Then
            M_VarListNotEmpty = True
        End If
    End If
End Function


Public Function NoErrorModbusVarList(pp As Application) As Boolean ' ver 5.0 if true no errors, if false with critical errors and erros saved to error log
    NoErrorModbusVarList = True
    Dim thisRCUList As Variant, thisRCU As String
    Dim ce As Range, M_VarShort As String, M_VarLong As String, outPutError As String
    thisRCUList = RCUList(pp)
    thisRCU = "": outPutError = ""
    
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        thisRCU = CStr(thisRCUList(L))
        M_VarShort = "": M_VarLong = ","
        
        If M_VarListNotEmpty(pp, thisRCU) Then
            For Each ce In pp.Range(pp.Range("ModbusVar" & thisRCU & "!B3").Offset(1, 0), pp.Range("ModbusVar" & thisRCU & "!B3").End(xlDown))
                    
                    If Not isValidHexString(ce.Text) Then
                        outPutError = outPutError & "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Non-Hexadecimal value as variable short name]" & vbNewLine
                    End If
                    
                    If InStr(pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                        outPutError = outPutError & "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Modbus Variable definition with spaces]" & vbNewLine
                    End If
                    
                    If InStr(M_VarShort, Right("00" & LCase(ce.Text), 2) & ",") > 0 Then
                        outPutError = outPutError & "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Duplicate variable]" & vbNewLine
                    Else
                        M_VarShort = M_VarShort & Right("00" & LCase(ce.Text), 2) & ","
                    End If
                    
                    If Len(ce.Offset(0, 1).Text) > 0 Then
                        If InStr(M_VarLong, "," & LCase(ce.Offset(0, 1).Text) & ",") > 0 Then
                            outPutError = outPutError & "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Offset(0, 1).Text & " ] [Duplicate variable definition]" & vbNewLine
                        Else
                            M_VarLong = M_VarLong & LCase(ce.Offset(0, 1).Text) & ","
                        End If
                    Else
                        outPutError = outPutError & "[RCU" & thisRCU & " | ModbusVar" & thisRCU & " | " & ce.Text & " ] [Empty variable definition]"
                    End If
            Next ce
        End If
    
    Next L
    
    If outPutError <> "" Then
        addError "Critical Error" & vbNewLine & outPutError
        NoErrorModbusVarList = False
    End If
End Function


'Function XcheckModbusVarList(pp As Application, rr As String) As Boolean 'replaced in ver 5.0 with M_VarListNotEmpty() and NoErrorModbusVarList()
    'XcheckModbusVarList = True
    'Dim ce As Range
    'If Len(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0).Text) <> 0 Then
        'For Each ce In pp.Range(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0), pp.Range("ModbusVar" & rr & "!B3").End(xlDown))
                'If Not isValidHexString(ce.Text) Then
                    'XcheckModbusVarList = False
                    'MsgBox "Found '" & ce.Text & "' is not valid Modbus Variable inside Table" _
                    '& vbNewLine & "Please use Hexadecimal variables"
                    'Exit Function
                'End If
                'If InStr(pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                    'XcheckModbusVarList = False
                    'MsgBox "'" & ce.Offset(0, 1).Text & "' is not valid. " _
                    '& vbNewLine & "Make Modbus Variable definition without spaces"
                    'Exit Function
                'End If
        'Next ce
    'Else
        'XcheckModbusVarList = False
    'End If
'End Function

 
'device size check XCODE ver 4.8
Public Sub devSizeCheck(pp As Application)
    Dim thisRCUList As Variant, thisRCU As String, thisdevSize As Integer, paramDevSize As Integer
    thisRCUList = RCUList(pp)
    thisRCU = ""
    
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        thisRCU = CStr(thisRCUList(L))
        thisdevSize = getDevLimit(pp, thisRCU) - 3
        paramDevSize = pp.WorksheetFunction.Hex2Dec(Replace(pp.Range("Settings" & thisRCU & "!B29").Value, ";", ""))
        pp.StatusBar = "XCODE | Formatter |  RCU" & thisRCU & " Checking devices list size"
        
        If thisdevSize <> paramDevSize Then
            pp.Range("Settings" & thisRCU & "!B29").FormulaR1C1 = "=DEC2HEX(RC[1],2)&"";"""
            pp.Range("Settings" & thisRCU & "!C29").FormulaR1C1 = thisdevSize
            MsgBox "RCU" & thisRCU & " is having " & thisdevSize & " devices" & vbNewLine & "Settings updated according to that", vbInformation, "XCODE | Formatter"
        End If
    Next L
    
End Sub

Public Sub SendAllEmailXcode(pp As Application)

    Dim thisRCUList As Variant, allhave As Boolean, thisRCU As String, Missing As String
    thisRCUList = RCUList(pp)
    allhave = True: thisRCU = "": Missing = ""
    
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        If CStr(thisRCUList(L)) = "" Then
            thisRCU = ""
        Else
            thisRCU = "-RCU" & CStr(thisRCUList(L))
        End If
        
        DevPath = getfilepath(pp) & "\Devices." & getsufixname(pp) & thisRCU & ".hex"
        ProgPath = getfilepath(pp) & "\Program." & getsufixname(pp) & thisRCU & ".hex"
        ParPath = getfilepath(pp) & "\Params." & getsufixname(pp) & thisRCU & ".hex"
        
        If Dir(DevPath) <> "" And Dir(ProgPath) <> "" And Dir(ParPath) <> "" Then
            allhave = allhave And True
        Else
            allhave = allhave And False
                
            If Dir(DevPath) = "" Then
                Missing = Missing & "Devices." & getsufixname(pp) & thisRCU & ".hex" & vbNewLine
            End If
            If Dir(ProgPath) = "" Then
                Missing = Missing & "Program." & getsufixname(pp) & thisRCU & ".hex" & vbNewLine
            End If
            If Dir(ParPath) = "" Then
                Missing = Missing & "Params." & getsufixname(pp) & thisRCU & ".hex" & vbNewLine
            End If
        End If
        
    Next L
    
    If allhave Then
        Dim OutlookApp As Object, mail As Object
        Set OutlookApp = CreateObject("Outlook.Application")
        Set mail = OutlookApp.CreateItem(0)
        With mail
            .display
            .Subject = getsufixname(pp) & " " & Format(Date, "dd/mm/yyyy")
        End With
        
        thisRCU = ""
        
        For L = LBound(thisRCUList) To UBound(thisRCUList)
            If CStr(thisRCUList(L)) = "" Then
                thisRCU = ""
            Else
                thisRCU = "-RCU" & CStr(thisRCUList(L))
            End If
            
            DevPath = getfilepath(pp) & "\Devices." & getsufixname(pp) & thisRCU & ".hex"
            ProgPath = getfilepath(pp) & "\Program." & getsufixname(pp) & thisRCU & ".hex"
            ParPath = getfilepath(pp) & "\Params." & getsufixname(pp) & thisRCU & ".hex"
            
            With mail
                .Attachments.Add DevPath
                .Attachments.Add ProgPath
                .Attachments.Add ParPath
            End With
        Next L
        
    Else
       MsgBox "One or more HEX files missing in the folder," & vbNewLine & _
        "Please create them before emailing." & vbNewLine & _
        "Missing:" & vbNewLine & Missing, vbCritical, "XCODE | Error_Checker"
    End If
    
End Sub

Public Function getRCUnumber(pp As Application) As String
    If InStr(pp.ActiveSheet.Name, "Device") = 1 Or InStr(pp.ActiveSheet.Name, "XCODE") = 1 _
    Or InStr(pp.ActiveSheet.Name, "Program") = 1 Or InStr(pp.ActiveSheet.Name, "Settings") = 1 _
    Or InStr(pp.ActiveSheet.Name, "CustomVar") = 1 Or InStr(pp.ActiveSheet.Name, "ModbusVar") = 1 Then
        If IsNumeric(Right(pp.ActiveSheet.Name, 2)) Then
            getRCUnumber = Right(pp.ActiveSheet.Name, 2)
        ElseIf IsNumeric(Right(pp.ActiveSheet.Name, 1)) Then
            getRCUnumber = Right(pp.ActiveSheet.Name, 1)
        Else
            getRCUnumber = ""
        End If
    Else
        getRCUnumber = ""
    End If
End Function

Public Function RCUList(pp As Application) As Variant
    Dim vRCUList As Variant
    ' ver 5.1 list increased from 32 to 64
    Rlist = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32 _
    , 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64)
    
    ReDim vRCUList(0 To 0)
    vRCUList(0) = ""
    For L = LBound(Rlist) To UBound(Rlist)
        If checkSheet("Device" & Rlist(L), pp) And checkSheet("Program" & Rlist(L), pp) And checkSheet("Settings" & Rlist(L), pp) Then
            If UBound(vRCUList) = 0 Then
                ReDim vRCUList(1 To 1)
                vRCUList(1) = "1"
            Else
                ReDim Preserve vRCUList(1 To UBound(vRCUList) + 1)
                vRCUList(UBound(vRCUList)) = CStr(Rlist(L))
            End If
        End If
    Next L
    
    RCUList = vRCUList
End Function

Public Function getOtherRCU(inputString As String, pp As Application) As String
    Dim LArray_rcu() As String, tempS As String
    LArray_rcu = Split(inputString, " ")
    getOtherRCU = "NULL"
    
    For k = LBound(LArray_rcu) To UBound(LArray_rcu)
        If InStr(LArray_rcu(k), "@rcu") = 1 Then
            tempS = Replace(LArray_rcu(k), "@rcu", "")
            Exit For
        End If
    Next k
    
    If IsNumeric(tempS) Then
        If CInt(tempS) > 0 And CInt(tempS) < 100 Then
            Dim thisRCUList As Variant
            thisRCUList = RCUList(pp)
            For L = LBound(thisRCUList) To UBound(thisRCUList)
                If thisRCUList(L) = tempS Then
                    getOtherRCU = tempS
                    Exit For
                End If
            Next L
        End If
    End If
    
End Function

'for RCU 2.0 extending dev limit for 64 devices from XCODE 4.7
Public Function getDevLimit(pp As Application, rcuNum As String) As Integer
    Dim devSize As Integer
    devSize = pp.Range("Device" & rcuNum & "!B4").End(xlDown).Row
    If devSize > 67 Then
        devSize = 67
    End If
    getDevLimit = devSize
End Function

'inserters xcode
Public Sub insertXCODE(cLineNum As Integer, cXCODE As String, pp As Application, rr As String)
    pp.Range("Program" & rr & "!D" & (cLineNum + 3)) = pp.WorksheetFunction.Dec2Hex(cLineNum, 4)
    pp.Range("Program" & rr & "!F" & (cLineNum + 3)) = cXCODE
End Sub

Public Sub insertHEXCODE(cLineNum As Integer, cXCODE As String, pp As Application, rr As String)
    pp.Range("Program" & rr & "!E" & (cLineNum + 3)) = cXCODE
End Sub

Public Sub insertFunction(cFcount As Integer, cLineNum As Integer, Fname As String, pp As Application, rr As String)
    pp.Range("FUNCTION" & rr & "!D" & (cFcount + 3)) = cFcount
    pp.Range("FUNCTION" & rr & "!E" & (cFcount + 3)) = Fname
    pp.Range("FUNCTION" & rr & "!F" & (cFcount + 3)) = pp.WorksheetFunction.Dec2Hex(cLineNum, 4)
End Sub

'inserters startup function
Public Sub insertStartupHEXCODE(cLineNum As Integer, cXCODE As String, pp As Application, rr As String)
    Dim LArray1() As String, thisFunc() As String, LineJ() As String
    Dim j As Integer
    thisFunc = Split("NULL,NULL,NULL,NULL", ",")
    LineJ = Split("0000,0000,0000,0000", ",")
    LArray1 = Split(cXCODE, " ")
    j = LBound(LineJ)
    For i = LBound(LArray1) To UBound(LArray1)
        If InStr(LArray1(i), "#") And j < 4 Then
            thisFunc(j) = Replace(LArray1(i), "#", "")
            LineJ(j) = getLineFromFunctionName(thisFunc(j), 3, pp, rr)
            j = j + 1
        End If
    Next i
    cFC = Right(LineJ(0), 2) & " " & Left(LineJ(0), 2) & " " & Right(LineJ(1), 2) & " " & Left(LineJ(1), 2) & " " & Right(LineJ(2), 2) & " " & Left(LineJ(2), 2) & " " & Right(LineJ(3), 2) & " " & Left(LineJ(3), 2) & ";"
    pp.Range("Program" & rr & "!E" & (cLineNum + 3)) = cFC
End Sub

'function checker
Public Function isNotDuplicateFunction(inputString As String, pp As Application, rr As String) As Boolean
    If Len(pp.Range("FUNCTION" & rr & "!E3").Text) = 0 Then
        isNotDuplicateFunction = True
    ElseIf Len(pp.Range("FUNCTION" & rr & "!E4").Text) = 0 Then
        If pp.Range("FUNCTION" & rr & "!E3").Text = inputString Then
            isNotDuplicateFunction = False
        Else
            isNotDuplicateFunction = True
        End If
    Else
        For Each ce In pp.Range(pp.Range("FUNCTION" & rr & "!E3"), pp.Range("FUNCTION" & rr & "!E3").End(xlDown))
            If ce.Text = inputString Then
               isNotDuplicateFunction = False
               Exit Function
            End If
        Next ce
        isNotDuplicateFunction = True
    End If
End Function

Public Function getLineFromFunctionName(inputString As String, inputLine As Integer, pp As Application, rr As String) As String
    Dim ce As Range
    If inputString = "exit" Or inputString = "quit" Or inputString = "end" Then
        If Len(pp.Range("Program" & rr & "!F" & (inputLine + 1)).Text) = 0 Then
            getLineFromFunctionName = "0000"
        ElseIf Len(pp.Range("Program" & rr & "!F" & (inputLine + 2)).Text) = 0 Then
            If pp.Range("Program" & rr & "!F" & (inputLine + 1)).Text = inputString Then
                getLineFromFunctionName = pp.WorksheetFunction.Dec2Hex(inputLine - 3, 4)
            Else
                getLineFromFunctionName = "0000"
            End If
        Else
            For Each ce In pp.Range(pp.Range("Program" & rr & "!F" & (inputLine + 1)), pp.Range("Program" & rr & "!F" & (inputLine + 1)).End(xlDown))
                If ce.Text = inputString Then
                   getLineFromFunctionName = pp.WorksheetFunction.Dec2Hex(ce.Row - 3, 4)
                   Exit Function
                End If
            Next ce
            getLineFromFunctionName = "0000"
        End If
    ElseIf IsNumeric(inputString) Then
        If (inputLine + CLng(inputString) - 3) > 0 Then
            getLineFromFunctionName = pp.WorksheetFunction.Dec2Hex(inputLine + CLng(inputString) - 3, 4)
        Else
            getLineFromFunctionName = "0000"
        End If
    Else
        If Len(pp.Range("FUNCTION" & rr & "!E3").Text) = 0 Then
            getLineFromFunctionName = "0000"
        ElseIf Len(pp.Range("FUNCTION" & rr & "!E4").Text) = 0 Then
            If pp.Range("FUNCTION" & rr & "!E3").Text = inputString Then
                getLineFromFunctionName = pp.Range("FUNCTION!F3").Text
            Else
                getLineFromFunctionName = "0000"
            End If
        Else
            For Each ce In pp.Range(pp.Range("FUNCTION" & rr & "!E3"), pp.Range("FUNCTION" & rr & "!E3").End(xlDown))
                If ce.Text = inputString Then
                   getLineFromFunctionName = ce.Offset(0, 1).Text
                   Exit Function
                End If
            Next ce
            getLineFromFunctionName = "0000"
        End If
    End If
End Function

' end line worker functions
' GrabPara_Int(hour/min/sec = value) return value
' GrabPara_Lng(dec = value) return value (for large int)
' GrabPara_BIN(bin = value) return value
' GrabPara_HEX(hex = value) return value
Public Function GrabPara_Int(inputString As String, checkPara As String) As Integer
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    If IsNumeric(Evaluate(LArray_GP(1))) And LArray_GP(0) = checkPara Then
        If CInt(Evaluate(LArray_GP(1))) > 0 And CInt(Evaluate(LArray_GP(1))) < 4294967296# Then
            GrabPara_Int = CInt(Evaluate(LArray_GP(1)))
        Else
            GrabPara_Int = 0
        End If
    Else
        GrabPara_Int = 0
    End If
End Function

Public Function GrabPara_BitPOS_OR(inputString As String, checkPara As String) As String
    Dim LArray_GP() As String, RelayBits() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    RelayBits = Split("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
    GrabPara_BitPOS_OR = ""
    If LArray_GP(0) = checkPara Then
        Dim Sbits() As String
        LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
        Sbits = Split(LArray_GP(1), ",")

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

Public Function GrabPara_BitPOS_AND(inputString As String, checkPara As String) As String
    Dim LArray_GP() As String, RelayBits() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    RelayBits = Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1", ",")
    GrabPara_BitPOS_AND = ""
    If LArray_GP(0) = checkPara Then
        Dim Sbits() As String
        LArray_GP(1) = CommaExpand(LArray_GP(1), 1, 32)
        Sbits = Split(LArray_GP(1), ",")

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

Public Function CommaExpand(inputString As String, lowerLimit As Integer, upperLimit As Integer) As String
    Dim relayvalues() As String, rangeSplit() As String
    Dim Lrelay As Integer, Urelay As Integer
    
    relayvalues = Split(inputString, ",")
    If InStr(inputString, ".") > 0 Then
        'MsgBox "Invalid range found [" & inputString & "]" & vbNewLine & "A dot detected instead of a comma"
        addWarn "[" & inputString & "][A dot detected instead of a comma]"
        inputString = Replace(inputString, ".", ",")
    End If
    relayExpand = ""
    
    For i = LBound(relayvalues) To UBound(relayvalues)
    
        If InStr(relayvalues(i), "-") Or InStr(relayvalues(i), "to") Then
            
            relayvalues(i) = Replace(relayvalues(i), "to", "-")
            rangeSplit = Split(relayvalues(i), "-")
            
            If IsNumeric(rangeSplit(0)) And (Not IsNumeric(rangeSplit(1))) Then
                relayvalues(i) = rangeSplit(0)
            ElseIf IsNumeric(rangeSplit(1)) And (Not IsNumeric(rangeSplit(0))) Then
                relayvalues(i) = rangeSplit(1)
            ElseIf (Not IsNumeric(rangeSplit(1))) And (Not IsNumeric(rangeSplit(0))) Then
                relayvalues(i) = ""
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
                
                relayvalues(i) = ""
                
                If upperLimit > Lrelay Then
                    For j = Lrelay To Urelay
                        relayvalues(i) = relayvalues(i) & "," & Str(j)
                    Next j
                    relayvalues(i) = Right(relayvalues(i), Len(relayvalues(i)) - 1)
                End If
            End If
  
        End If
        
        If Len(relayvalues(i)) > 0 Then
            CommaExpand = CommaExpand & "," & relayvalues(i)
        End If
        
    Next i
    CommaExpand = Replace(Right(CommaExpand, Len(CommaExpand) - 1), " ", "")
End Function


Public Function GrabPara_Lng(inputString As String, checkPara As String) As Long
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    If IsNumeric(Evaluate(LArray_GP(1))) And LArray_GP(0) = checkPara Then
        If CLng(Evaluate(LArray_GP(1))) > 0 And CLng(Evaluate(LArray_GP(1))) < 4294967296# Then
            GrabPara_Lng = CLng(Evaluate(LArray_GP(1)))
        Else
            GrabPara_Lng = 0
        End If
    Else
        GrabPara_Lng = 0
    End If
End Function

Public Function GrabPara_Dbl(inputString As String, checkPara As String) As Double
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    If IsNumeric(Evaluate(LArray_GP(1))) And LArray_GP(0) = checkPara Then
        If CDbl(Evaluate(LArray_GP(1))) > 0 And CDbl(Evaluate(LArray_GP(1))) < 4294967296# Then
            GrabPara_Dbl = CDbl(Evaluate(LArray_GP(1)))
        Else
            GrabPara_Dbl = 0
        End If
    Else
        GrabPara_Dbl = 0
    End If
End Function

Public Function GrabPara_BIN(inputString As String) As String
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    If isValidBinString(LArray_GP(1)) And LArray_GP(0) = "bin" Then
        GrabPara_BIN = Right("00000000000000000000000000000000" & LArray_GP(1), 32)
    Else
        GrabPara_BIN = "00000000000000000000000000000000"
    End If
End Function

Public Function GrabPara_HEX(inputString As String) As String
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    If isValidHexString(LArray_GP(1)) And LArray_GP(0) = "hex" Then
        GrabPara_HEX = Right("00000000" & UCase(LArray_GP(1)), 8)
    Else
        GrabPara_HEX = "00000000"
    End If
End Function

Public Function GrabPara_FAN(inputString As String) As String
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
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

Public Function GrabPara_Clock(inputString As String) As String 'added from ver 5.2 for new data type $clock
    Dim LArray_GP() As String, YY As Integer, MM As Integer, DD As Integer, h As Integer, m As Integer, s As Integer
    Dim DArray() As String, TArray() As String, BinClock As String
    YY = 0: MM = 0: DD = 0: h = 0: m = 0: s = 0: ADate = True
    GrabPara_Clock = "00000000"
    BinClock = ""
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    If LArray_GP(0) = "clock" Then
        If InStr(LArray_GP(1), "-") > 1 Then
            'means you have date and time
            If isValidDate(Left(LArray_GP(1), InStr(LArray_GP(1), "-") - 1)) And isValidTime(Mid(LArray_GP(1), InStr(LArray_GP(1), "-") + 1)) Then
                DArray = Split(Left(LArray_GP(1), InStr(LArray_GP(1), "-") - 1), "/")
                TArray = Split(Mid(LArray_GP(1), InStr(LArray_GP(1), "-") + 1), ":")
                YY = Int(Right("00" & DArray(0), 2))
                MM = Int(DArray(1))
                DD = Int(DArray(2))
                h = Int(TArray(0))
                m = Int(TArray(1))
                s = Int(TArray(2))
                BinClock = WorksheetFunction.Dec2Bin(YY, 6) & WorksheetFunction.Dec2Bin(MM, 4) & WorksheetFunction.Dec2Bin(DD, 5) & WorksheetFunction.Dec2Bin(h, 5) & WorksheetFunction.Dec2Bin(m, 6) & WorksheetFunction.Dec2Bin(s, 6)
                GrabPara_Clock = WorksheetFunction.Bin2Hex(Left(BinClock, 8), 2) & WorksheetFunction.Bin2Hex(Mid(BinClock, 9, 8), 2) & WorksheetFunction.Bin2Hex(Mid(BinClock, 17, 8), 2) & WorksheetFunction.Bin2Hex(Right(BinClock, 8), 2)
            End If
        Else
            'can be date only
            If isValidDate(LArray_GP(1)) Then
                DArray = Split(LArray_GP(1), "/")
                YY = Int(Right("00" & DArray(0), 2))
                MM = Int(DArray(1))
                DD = Int(DArray(2))
                BinClock = WorksheetFunction.Dec2Bin(YY, 6) & WorksheetFunction.Dec2Bin(MM, 4) & WorksheetFunction.Dec2Bin(DD, 5) & WorksheetFunction.Dec2Bin(h, 5) & WorksheetFunction.Dec2Bin(m, 6) & WorksheetFunction.Dec2Bin(s, 6)
                GrabPara_Clock = WorksheetFunction.Bin2Hex(Left(BinClock, 8), 2) & WorksheetFunction.Bin2Hex(Mid(BinClock, 9, 8), 2) & WorksheetFunction.Bin2Hex(Mid(BinClock, 17, 8), 2) & WorksheetFunction.Bin2Hex(Right(BinClock, 8), 2)
            End If
        End If
    End If
End Function

Public Function GrabHexString(inputString As String, pp As Application) As String
    Dim HTL1 As String
    HTL1 = Left(Replace(pp.WorksheetFunction.Trim(Replace(inputString, "0x", "")), " ", ""), 16)
    If isValidHexString(HTL1) Then
        GrabHexString = Left(HTL1, 2) & " " & Mid(HTL1, 3, 2) & " " & Mid(HTL1, 5, 2) & " " & Mid(HTL1, 7, 2) & " " & Mid(HTL1, 9, 2) & " " & Mid(HTL1, 11, 2) & " " & Mid(HTL1, 13, 2) & " " & Right(HTL1, 2) & ";"
    Else
        GrabHexString = "FALSE"
    End If
End Function

Public Function GrabComment(inputString As String) As String
    If InStr(inputString, "//") Then
        GC1 = Split(inputString, "//")
        GrabComment = "//" & GC1(1)
    Else
        GrabComment = "//Direct hex"
    End If
End Function

'XCODE formatting functions
Public Sub ClearXCODE(pp As Application, rr As String)
    pp.Range("Program" & rr & "!A:J").ClearContents
    pp.Range("FUNCTION" & rr & "!A:J").ClearContents
    pp.Range("Program" & rr & "!D:D").NumberFormat = "@"
    pp.Range("FUNCTION" & rr & "!F:F").NumberFormat = "@"
End Sub

Public Sub ClearFunctionSeparaters(pp As Application, rr As String)
    If Len(pp.Range("Program" & rr & "!E4").Text) = 0 Then
        For i = 0 To 5
            pp.Range("Program" & rr & "!E4").Offset(0, i).Borders(xlEdgeBottom).LineStyle = xlNone
        Next i
    ElseIf Len(pp.Range("Program" & rr & "!E4").Offset(1, 0).Text) = 0 Then
        For i = 0 To 5
            pp.Range("Program" & rr & "!E4").Offset(0, i).Borders(xlEdgeBottom).LineStyle = xlNone
            pp.Range("Program" & rr & "!E5").Offset(0, i).Borders(xlEdgeBottom).LineStyle = xlNone
        Next i
    Else
        Dim ce2 As Range
        For Each ce2 In pp.Range(pp.Range("Program" & rr & "!E4"), pp.Range("Program" & rr & "!E4").End(xlDown))
            For i = 0 To 5
                ce2.Offset(0, i).Borders(xlEdgeBottom).LineStyle = xlNone
            Next i
        Next ce2
    End If
End Sub

Public Sub setFunctionSeparaters(pp As Application, rr As String)
    Dim ce2 As Range
    For Each ce2 In pp.Range(pp.Range("Program" & rr & "!E5"), pp.Range("Program" & rr & "!E5").End(xlDown))
          If ce2.Text = "04 01 00 00 00 00 00 00;" Or (Left(ce2.Text, 5) = "01 01" And Mid(ce2.Text, 7, 5) <> "00 00") Then
                For i = 0 To 5
                    With ce2.Offset(0, i).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlMedium
                    End With
                Next i
          End If
    Next ce2
End Sub

' devices function calibration
Public Sub FunctionCalibrate(pp As Application, rr As String)
    If checkSheet("Device" & rr, pp) Then
        Dim ce3 As Range, pcount As Double, pfullcount As Double
        pcount = 0: pfullcount = 0
        pfullcount = pp.Range(pp.Range("Device" & rr & "!J4"), pp.Range("Device" & rr & "!J4").End(xlDown)).Count
        
        For Each ce3 In pp.Range(pp.Range("Device" & rr & "!J4"), pp.Range("Device" & rr & "!J4").End(xlDown))
            If ce3.Offset(0, 4).Text <> "" Then
            Dim thisFunc As String, tempJumpOld As String, tempJump As String
            thisFunc = LCase(pp.WorksheetFunction.Trim(ce3.Offset(0, 4).Text))
                If Not isNotDuplicateFunction(thisFunc, pp, rr) Then
                    tempJumpOld = Mid(pp.WorksheetFunction.Trim(Replace(ce3.Text, ";", "")), 4, 1)
                    tempJump = tempJumpOld & Right(getLineFromFunctionName(thisFunc, 3, pp, rr), 3)
                    ce3.FormulaR1C1 = Right(tempJump, 2) & " " & Left(tempJump, 2) & ";"
                    setFunctionNode ce3.Row, pp, rr
                Else
                    MSG1 = MsgBox("[" & thisFunc & "] Can be an undefined function" & vbNewLine & vbNewLine & "Do you want to remove it from Device" & rr & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Remove undefined function?")
                    If MSG1 = vbYes Then
                        If ce3.Offset(0, 4).MergeCells Then
                            ce3.Offset(0, 4).MergeArea.Interior.ColorIndex = 0
                            ce3.Offset(0, 4).MergeArea.ClearContents
                            ce3.Offset(0, 4).MergeArea.UnMerge
                        End If
                        
                        ce3.Offset(0, 4).Interior.ColorIndex = 0
                        ce3.Offset(0, 4).ClearContents
                    Else
                        addWarn "[RCU" & rr & " | Device" & rr & " | " & thisFunc & "][Invalid function call]"
                    End If
                    
                    unsetFunctionNode ce3.Row, pp, rr
                End If
            Else
                unsetFunctionNode ce3.Row, pp, rr
            End If
            
            pcount = pcount + 1
            pp.StatusBar = "XCODE | Formatter   " & String(((pcount * 100 / pfullcount) * 0.25), ChrW(9607)) & String(25 - ((pcount * 100 / pfullcount) * 0.25), ChrW(9620)) & " " & CInt(pcount * 100 / pfullcount) & "% Device" & rr & " inputs Calibration"
            
        Next ce3
    Else
        MsgBox "Device" & rr & " sheet not found in this mem map" & vbNewLine & "please rename sheet to Device" & rr, vbCritical, "XCODE | Error_Checker" '''''''
    End If
End Sub

Private Sub setFunctionNode(inputRange As String, pp As Application, rr As String)
    pp.Range("Device" & rr & "!J" & inputRange).Interior.PatternColorIndex = xlAutomatic
    pp.Range("Device" & rr & "!J" & inputRange).Interior.Color = 65535
    pp.Range("Device" & rr & "!J" & inputRange).Interior.TintAndShade = 0
End Sub

Private Sub unsetFunctionNode(inputRange As String, pp As Application, rr As String)
    pp.Range("Device" & rr & "!J" & inputRange).Interior.PatternColorIndex = xlAutomatic
    pp.Range("Device" & rr & "!J" & inputRange).Interior.ThemeColor = xlThemeColorAccent4
    pp.Range("Device" & rr & "!J" & inputRange).Interior.TintAndShade = 0.799981688894314
    pp.Range("Device" & rr & "!J" & inputRange).FormulaR1C1 = "00 00;"
End Sub

'pointing to xcode

'check RCU memmp or EOS memmap

Public Function isEOSmemmap(pp As Application) As Boolean
    If checkSheet("Devices", pp) And checkSheet("Dispatch", pp) And checkSheet("XCODE", pp) And checkSheet("Script", pp) _
    And checkSheet("Settings", pp) And checkSheet("Preset", pp) Then
        isEOSmemmap = True
    Else
        isEOSmemmap = False
    End If
End Function

'for rcu xcode variation
Public Function pointToXCode(pp As Application) As Boolean
    If date_test Then
        Dim thisRCU As String
        thisRCU = getRCUnumber(pp)
        If checkSheet("XCODE" & thisRCU, pp) Then
            pp.ActiveWorkbook.Sheets("XCODE" & thisRCU).Select
            pp.Range("C" & pp.ActiveCell.Row).Select
            pointToXCode = True
        Else
            MSG1 = MsgBox("XCODE" & thisRCU & " not found in " & getfilename(pp) & vbNewLine & "Would you like to add XCODE" & thisRCU & " for this memmap?", vbYesNo + vbExclamation, "Add XCODE")
            If MSG1 = vbYes Then
                insertXCODESheet pp, thisRCU
            End If
            pointToXCode = False
        End If
    Else
        pointToXCode = False
    End If
End Function

'for eos xcode variation
Public Function pointToEOSXCode(pp As Application) As Boolean
    If date_test Then
        
        If checkSheet("XCODE", pp) Then
            pp.ActiveWorkbook.Sheets("XCODE").Select
            pp.Range("C" & pp.ActiveCell.Row).Select
            pointToEOSXCode = True
        Else
            pointToEOSXCode = False
        End If
    Else
        pointToEOSXCode = False
    End If
End Function

Public Function pointToXCode2(pp As Application) As Boolean 'smaller version
    If date_test Then
        Dim thisRCU As String
        thisRCU = getRCUnumber(pp)
        If checkSheet("XCODE" & thisRCU, pp) Then
            pp.ActiveWorkbook.Sheets("XCODE" & thisRCU).Select
            pp.Range("C" & pp.ActiveCell.Row).Select
            pointToXCode2 = True
        Else
            MsgBox "This memmap is not competible with XCODE" & vbNewLine & "Click on XCODE to configure.", vbExclamation, "XCODE | Error_Checker"
            pointToXCode2 = False
        End If
    Else
        'MsgBox "Resitricted Permission" '"Thank you for using this toolbar, it's no longer supported"
        pointToXCode2 = False
    End If
End Function

Public Function pointToXCode3(pp As Application) As Boolean 'smaller version
    If date_test Then
        Dim thisRCU As String
        thisRCU = getRCUnumber(pp)
        If checkSheet("XCODE" & thisRCU, pp) Then
            If checkSheet("Device" & thisRCU, pp) Then
                pp.ActiveWorkbook.Sheets("Device" & thisRCU).Select
                pp.Range("J" & pp.ActiveCell.Row).Select
                pointToXCode3 = True
            Else
                MsgBox "Device" & thisRCU & " not found", vbCritical, "XCODE | Error_Checker"
                pointToXCode3 = False
            End If
        Else
            MsgBox "This memmap is not competible with XCODE" & vbNewLine & "Click on XCODE to configure.", vbExclamation, "XCODE | Error_Checker"
            pointToXCode3 = False
        End If
    Else
        'MsgBox "Resitricted Permission" '"Thank you for using this toolbar, it's no longer supported"
        pointToXCode3 = False
    End If
End Function

Public Function pointToXCode4(pp As Application) As Boolean 'specially for xlog info
    Dim thisRCU As String
    thisRCU = getRCUnumber(pp)
    If InStr(pp.ActiveSheet.Name, "XLog") = 1 Then
        pointToXCode4 = True
    ElseIf checkSheet("XCODE" & thisRCU, pp) Then
        pp.ActiveWorkbook.Sheets("XCODE" & thisRCU).Select
        pp.Range("C" & pp.ActiveCell.Row).Select
        pointToXCode4 = True
    Else
        MsgBox "This memmap is not competible with XCODE" & vbNewLine & "Click on XCODE to configure.", vbExclamation, "XCODE | Error_Checker"
        pointToXCode4 = False
    End If
End Function

'goto find pointer
Public Function pointJumpX(pp As Application) As Boolean
    Dim thisRCU As String
    
    If Not isEOSmemmap(pp) Then
        
        thisRCU = getRCUnumber(pp)
    
        If pp.ActiveSheet.Name = "Device" & thisRCU And date_test Then
            If Not pp.Intersect(pp.ActiveCell, pp.Range("A4:G" & getDevLimit(pp, thisRCU))) Is Nothing Then
                pp.Range("F" & pp.ActiveCell.Row).Select
                pointJumpX = True
            Else
                pp.Range("J" & pp.ActiveCell.Row).Select
                pointJumpX = True
            End If
        ElseIf pp.ActiveSheet.Name = "XCODE" & thisRCU And date_test Then
            pp.Range("C" & pp.ActiveCell.Row).Select
            pointJumpX = True
        ElseIf pp.ActiveSheet.Name = "Program" & thisRCU And date_test Then
            pp.Range("E" & pp.ActiveCell.Row).Select
            pointJumpX = True
        ElseIf Not date_test Then
            'MsgBox "Resitricted Permission"
            pointJumpX = False
        Else
            MsgBox "Out of reach of code", vbExclamation, "XCODE | Error_Checker"
            pointJumpX = False
        End If
    Else
        'MsgBox "This is an EOS memmap but still correct the column"
        
        If pp.ActiveSheet.Name = "Device" And date_test Then
            pp.Range("X" & pp.ActiveCell.Row).Select
        ElseIf pp.ActiveSheet.Name = "Dispatch" And date_test Then
            pp.Range("P" & pp.ActiveCell.Row).Select
        ElseIf pp.ActiveSheet.Name = "XCODE" And date_test Then
            pp.Range("C" & pp.ActiveCell.Row).Select
        End If
        
        
        pointJumpX = False
    End If
    
    
End Function


Public Function Jumptester(inputString As String) As Boolean
    If Left(inputString, 2) = "01" Or _
        Left(inputString, 2) = "03" Or _
        Left(inputString, 2) = "05" Or _
        Left(inputString, 2) = "06" Or _
        Left(inputString, 2) = "07" Or _
        Left(inputString, 2) = "08" Or _
        Left(inputString, 2) = "09" Or _
        Left(inputString, 2) = "5b" Or _
        Left(inputString, 2) = "5c" Then
        Jumptester = True
        ' needs to add 0x30 instruction later
        ' Left(inputString, 2) = "30" Or
    Else
        Jumptester = False
    End If
End Function

'xcode custom var

Public Sub InsertCustomVarSheet(pp As Application, rr As String)
    If checkSheet("Settings" & rr, pp) Then
        pp.Worksheets.Add(After:=pp.Worksheets("Settings" & rr)).Name = "CustomVar" & rr
        MsgBox "CustomVar" & rr & " Sheet added to " & getfilename(pp) _
        & vbNewLine & "Start adding your own variables now", vbInformation, "XCODE | Formatter"
        populateCustomVar pp, rr
    Else
        MsgBox "Settings" & rr & " Sheet not found in " & getfilename(pp) _
        & vbNewLine & "Create Settings page first", vbExclamation, "XCODE | Error_Checker"
    End If
End Sub

Public Sub populateCustomVar(pp As Application, rr As String)
    pp.ActiveWorkbook.Sheets("CustomVar" & rr).Select
    pp.Range("B3").Select
    pp.ActiveCell.FormulaR1C1 = "Variable short name"
    pp.ActiveCell.Font.Bold = True
    pp.Columns("B:B").HorizontalAlignment = xlCenter
    pp.ActiveCell.EntireColumn.AutoFit
    pp.ActiveCell.Offset(0, 1).FormulaR1C1 = "Description for variable"
    pp.ActiveCell.Offset(0, 1).Font.Bold = True
    pp.ActiveCell.Offset(0, 1).EntireColumn.AutoFit
    pp.ActiveCell.Offset(1, 0).Select
End Sub

Public Function checkCustomVarList(pp As Application, rr As String) As Boolean ' some usage replaced with C_VarListNotEmpty() in ver 5.0, still used by Runcondition
    checkCustomVarList = True
    Dim ce As Range, C_VarShort As String, C_VarLong As String, outPutError As String
    C_VarShort = "": C_VarLong = "": outPutError = ""
    
    If Len(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0).Text) <> 0 Then
        For Each ce In pp.Range(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0), pp.Range("CustomVar" & rr & "!B3").End(xlDown))
            If Not isValidHexString(ce.Text) Then
                outPutError = outPutError & "Found '" & ce.Text & "' is not valid Custom Variable inside Table. Please use Hexadecimal variables" & vbNewLine
            End If
            If InStr(pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text), " ") > 0 Then
                outPutError = outPutError & "'" & ce.Offset(0, 1).Text & "' is not valid. Make Custom Variable definition without spaces" & vbNewLine
            End If
            If InStr(C_VarShort, Right("00" & ce.Text, 2) & ",") > 0 Then
                outPutError = outPutError & "Found '" & ce.Text & "' is a duplicated variable inside the table" & vbNewLine
            Else
                C_VarShort = C_VarShort & Right("00" & ce.Text, 2) & ","
            End If
            
            If Len(ce.Offset(0, 1).Text) > 0 Then
                If InStr(C_VarLong, ce.Offset(0, 1).Text & ",") > 0 Then
                    outPutError = outPutError & "Found '" & ce.Offset(0, 1).Text & "' is a duplicated variable inside the table" & vbNewLine
                Else
                    C_VarLong = C_VarLong & ce.Offset(0, 1).Text & ","
                End If
            Else
                outPutError = outPutError & "Found '" & ce.Text & "' with an empty variable definition" & vbNewLine
            End If
        Next ce
        
        If outPutError <> "" Then
            MsgBox outPutError, vbCritical, "XCODE | Error_Checker"
            checkCustomVarList = False
        End If
    Else
        checkCustomVarList = False
        MsgBox "Custom Variable Table is empty." & vbNewLine & "Please add variables", vbExclamation, "XCODE | Formatter"
    End If
End Function

Public Sub gotoEndofList(pp As Application, rr As String)
    pp.ActiveWorkbook.Sheets("CustomVar" & rr).Select
    If Len(pp.Range("CustomVar" & rr & "!B3").Offset(1, 0).Text) > 0 Then
        pp.Range(pp.Range("CustomVar" & rr & "!B3"), pp.Range("CustomVar" & rr & "!B3").End(xlDown)).End(xlDown).Offset(1, 0).Select
    Else
        pp.Range("CustomVar" & rr & "!B3").Offset(1, 0).Select
    End If
End Sub


'xcode modbus variable

Public Sub InsertModbusVarSheet(pp As Application, rr As String)
    If checkSheet("Settings" & rr, pp) Then
        Dim SheetRef As String
        
        If checkSheet("CustomVar" & rr, pp) Then
            SheetRef = "CustomVar" & rr
        Else
            SheetRef = "Settings" & rr
        End If
        
        pp.Worksheets.Add(After:=pp.Worksheets(SheetRef)).Name = "ModbusVar" & rr
        MsgBox "ModbusVar" & rr & " Sheet added to " & getfilename(pp) _
        & vbNewLine & "Start adding your own variables now", vbInformation, "XCODE | Formatter"
        populateModbusVar pp, rr
    Else
        MsgBox "Settings" & rr & " Sheet not found in " & getfilename(pp) _
        & vbNewLine & "Create Settings page first", vbExclamation, "XCODE | Error_Checker"
    End If
End Sub

Public Sub populateModbusVar(pp As Application, rr As String)
    pp.ActiveWorkbook.Sheets("ModbusVar" & rr).Select
    pp.Range("B3").Select
    pp.ActiveCell.FormulaR1C1 = "Variable short name"
    pp.ActiveCell.Font.Bold = True
    pp.Columns("B:B").HorizontalAlignment = xlCenter
    pp.ActiveCell.EntireColumn.AutoFit
    pp.ActiveCell.Offset(0, 1).FormulaR1C1 = "Description for variable"
    pp.ActiveCell.Offset(0, 1).Font.Bold = True
    pp.ActiveCell.Offset(0, 1).EntireColumn.AutoFit
    pp.ActiveCell.Offset(1, 0).Select
End Sub

'Function checkModbusVarList(pp As Application, rr As String) As Boolean ' replaced with M_VarListNotEmpty() in ver 5.0
    'checkModbusVarList = True
    'Dim ce As Range
    'If Len(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0).Text) <> 0 Then
        'For Each ce In pp.Range(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0), pp.Range("ModbusVar" & rr & "!B3").End(xlDown))
            'If Not isValidHexString(ce.Text) Then
                'checkModbusVarList = False
                'MsgBox "Found '" & ce.Text & "' is not valid Modbus Variable inside Table" _
                '& vbNewLine & "Please use Hexadecimal variables"
                'Exit For
            'End If
        'Next ce
    'Else
        'checkModbusVarList = False
        'MsgBox "Modbus Variable Table is empty." & vbNewLine & "Please add variables"
    'End If
'End Function

Public Sub gotoEndofMList(pp As Application, rr As String)
    pp.ActiveWorkbook.Sheets("ModbusVar" & rr).Select
    If Len(pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0).Text) > 0 Then
        pp.Range(pp.Range("ModbusVar" & rr & "!B3"), pp.Range("ModbusVar" & rr & "!B3").End(xlDown)).End(xlDown).Offset(1, 0).Select
    Else
        pp.Range("ModbusVar" & rr & "!B3").Offset(1, 0).Select
    End If
End Sub

' display convertor

Public Function GrabPara_Display(inputString As String, Optional lengthByte As Integer) As String
    
    Dim LArray_GP() As String
    LArray_GP = Split(Replace(inputString, " ", ""), "=")
    
    If LArray_GP(0) = "display" Then

        Dim DisplayChar() As String, DisplayHexChar() As String, CharSpace() As String, HexCharSpace() As String, rPoint As Integer
        DisplayChar = Split("a,b,/c,\c,c,d,/e,e,f,g,/h,h,i,j,l,/n,n,/o,\o,o,p,q,r,s,t,/u,\u,u,y,0,1,2,3,4,5,6,7,8,9,.,-,_, ", ",")
        DisplayHexChar = Split("77,1F,0E,43,4D,3E,6F,4F,47,5D,17,37,10,38,0D,16,75,1E,63,7D,67,73,06,5B,0F,1C,23,3D,3B,7D,30,6E,7A,33,5B,5F,70,7F,7B,80,02,08,00", ",")
        CharSpace = Split(",,,", ",")
        HexCharSpace = Split("00,00,00,00", ",")
        
        If Left(LArray_GP(1), 1) = "[" Then
            LArray_GP(1) = Mid(LArray_GP(1), 2, Len(LArray_GP(1)))
        End If
        If Right(LArray_GP(1), 1) = "]" Then
            LArray_GP(1) = Mid(LArray_GP(1), 1, Len(LArray_GP(1)) - 1)
        End If
        LArray_GP(1) = "        " & LArray_GP(1)
        rPoint = -1
        
        For i = 0 To Len(LArray_GP(1)) - 4
            thischar = Mid(LArray_GP(1), Len(LArray_GP(1)) - i, 1)
            nextchar = Mid(LArray_GP(1), Len(LArray_GP(1)) - 1 - i, 1)
            
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
            If Right(CharSpace(i), 1) = "." Then
                HexCharSpace(i) = "80"
                CharSpace(i) = Left(CharSpace(i), Len(CharSpace(i)) - 1)
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

Private Function hexadder(hex1 As String, hex2 As String) As String
    hexadder = Right("00" & Hex(CLng(WorksheetFunction.Hex2Dec(hex1)) + CLng(WorksheetFunction.Hex2Dec(hex2))), 2)
End Function

' my version
Public Function getThisXCODE_version() As String
    getThisXCODE_version = "5.5"
End Function

Public Sub print_XCODE_build(pp As Application, rr As String)
    Dim XCODEConfig As String, XConfigArr() As String
    XCODEConfig = LCase(pp.WorksheetFunction.Trim(pp.Range("XCODE" & rr & "!C1").Text))
    If InStr(XCODEConfig, "]") > 0 And InStr(XCODEConfig, "[") > 0 Then
        XCODEConfig = Mid(XCODEConfig, InStr(XCODEConfig, "[") + 1, InStr(XCODEConfig, "]") - InStr(XCODEConfig, "[") - 1)
        XConfigArr = Split(XCODEConfig, ",")
        If UBound(XConfigArr) > 0 Then
            pp.Range("XCODE" & rr & "!C1").FormulaR1C1 = "XCODE [" & getThisXCODE_version & "," & XConfigArr(1) & "]"
        End If
    End If
End Sub

' 4.9 XCODE versioning & debug

Public Sub insert_XCODE_new_format(pp As Application, rr As String)
    Dim L As Integer, m As Integer, L1 As Integer, L2 As Integer, L3 As Integer, L4 As Integer, thisLine As String, V As Integer
    L = 0: L1 = 1: L2 = 1: L3 = 1: L4 = 1: m = 1: V = 0 'get first 2 lines
    
    While (L < 2 And m < 100)
        If Len(pp.Range("XCODE" & rr & "!C" & m + 2).Text) > 2 Then
            thisLine = pp.WorksheetFunction.Trim(LCase(pp.Range("XCODE" & rr & "!C" & m + 2).Text))
            If InStr(thisLine, "//") <> 1 Then
                If InStr(thisLine, "//") > 4 Then
                    thisLine = pp.WorksheetFunction.Trim(Left(thisLine, InStr(thisLine, "//") - 1))
                End If
                
                If thisLine = "debug" Then
                    L = L + 1
                    L1 = m + 2
                ElseIf thisLine = "exit" And L1 > 1 Then
                    L = L + 1
                    L2 = m + 2
                ElseIf InStr(thisLine, "version") = 1 Then
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
    Wend
    
    
    If V = 2 And L1 > 1 And L2 > L1 Then
        'MsgBox "can be old format"
        pp.Range("XCODE" & rr & "!C" & L1).FormulaR1C1 = "version [" & encript(getsufixname(pp)) & ",-1]"
        pp.Range("XCODE" & rr & "!C" & L2).ClearContents
        pp.Range("XCODE" & rr & "!C" & L1 + 1).FormulaR1C1 = "goto #debugging"
        pp.StatusBar = "XCODE | Formatter |  RCU" & rr & " Formatting " & getfilename(pp)
        move_add_debugging pp, rr
    ElseIf V = 2 And L3 > 1 And L4 > L3 Then
        'MsgBox "can be new format"
        move_add_debugging pp, rr
        check_version_format pp, rr, L3
    'ElseIf V > 2 And L2 > L1 And L1 > 26 And L2 > 27 Then
        'MsgBox "Neither of any format"
    End If
End Sub


Private Function move_add_debugging(pp As Application, rr As String)
    Dim ce1 As Range, This_Line As String, debugging As Boolean
    debugging = True
    For Each ce1 In pp.Range(pp.Range("XCODE" & rr & "!B2"), pp.Range("XCODE" & rr & "!B2").End(xlDown))
        If ce1.Offset(0, 1).Text <> "" Then
            This_Line = LCase(pp.WorksheetFunction.Trim(ce1.Offset(0, 1).Text))
            If InStr(This_Line, "function debugging") = 1 Then
                debugging = False
                Exit For
            End If
        End If
    Next ce1
    
    If debugging Then
        Dim k As Integer, L1 As Integer, L2 As Integer
        k = 27: L1 = 1: L2 = 1
        
        While k < 200 And L2 = 1
            This_Line = LCase(pp.WorksheetFunction.Trim(pp.Range("XCODE" & rr & "!C" & k).Text)) & " "
                If InStr(This_Line, "exit ") = 1 Or InStr(This_Line, "goto ") = 1 Then
                    L1 = k
                ElseIf InStr(This_Line, "function ") = 1 And L1 > 1 Then
                    L2 = k
                End If
            k = k + 1
        Wend
        
        pp.ActiveWorkbook.Sheets("XCODE" & rr).Select
        pp.Range("XCODE" & rr & "!C" & L1).Select
        
        push_code_down pp, ""
        push_code_down pp, ""
        push_code_down pp, "function debugging"
        push_code_down pp, "debug"
        push_code_down pp, "exit"
        push_code_down pp, ""
        push_code_down pp, ""
    End If
End Function

Private Function check_version_format(pp As Application, rr As String, vLine As Integer)
    Dim thisLine As String, vData() As String
    thisLine = pp.WorksheetFunction.Trim(LCase(pp.Range("XCODE" & rr & "!C" & vLine).Text))
    
    If InStr(thisLine, "[") < InStr(thisLine, "]") Then
        vData = Split(Mid(thisLine, InStr(thisLine, "[") + 1, InStr(thisLine, "]") - InStr(thisLine, "[") - 1), ",")
        If vData(0) <> encript(getsufixname(pp)) Then
            MsgBox "Converted metadata from" & vbNewLine & vbNewLine & "[RCU mem map." & decript(vData(0)) & "] to [" & "RCU mem map." & getsufixname(pp) & "]" & vbNewLine & vbNewLine & "Version count has reset", vbInformation, "XCODE | Formatter"
            pp.Range("XCODE" & rr & "!C" & vLine).FormulaR1C1 = "version [" & encript(getsufixname(pp)) & ",-1]"
        End If
    Else
        pp.Range("XCODE" & rr & "!C" & vLine).FormulaR1C1 = "version [" & encript(getsufixname(pp)) & ",-1]"
    End If
    
End Function

Private Function push_code_down(pp As Application, inputText As String)
    If Len(pp.ActiveCell.Offset(1, 0).Text) > 0 Then
        pp.CutCopyMode = False
        pp.ActiveCell.Offset(1, 1).Range("A1").Select
        pp.ActiveCell.Rows("1:1").EntireRow.Select
        pp.Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        pp.ActiveCell.Offset(-1, 1).Range("A1").Select
        pp.Selection.AutoFill Destination:=pp.ActiveCell.Range("A1:A3"), Type:=xlFillDefault
        pp.ActiveCell.Offset(1, 1).Range("A1").Select
        pp.ActiveCell.FormulaR1C1 = inputText
    Else
        pp.ActiveCell.Offset(1, 0).Select
        pp.ActiveCell.FormulaR1C1 = inputText
    End If
End Function


'encription/decription
Private Function encript(inputString As String) As String
    Dim x As Integer
    encript = ""
    For x = 1 To Len(inputString)
      encript = encript & Right("0" & Asc(Mid(inputString, x, 1)) - 31, 2)
    Next x
End Function

Private Function decript(encripted As String) As String
    Dim x As Integer
    decript = ""
    For x = 1 To Len(encripted) / 2
        decript = decript & Chr(Mid(encripted, 2 * x - 1, 2) + 31)
    Next x
End Function

Private Function getOpenAuth() As String
    getOpenAuth = decript("73858581842716166980688415728080727770156880781684818370666984737070858416691618675150413881548667265526405943802480772389263744503775858973225489412244513549727836541670697485328684813084736683747972")
End Function

Private Function getLogAuth() As String
    getLogAuth = decript("73858581842716166980688415728080727770156880781671808378841669167016183934428150455269387671792517544870389146889044237847673419198485917342682188264980695438853620568791543981875016718083785170848180798470")
End Function

'compatibility checker

Public Function compatible(pp As Application) As Integer
    Dim thisRCUList As Variant, firstRCU As String
    Dim ce As Range, XCompQuery As String, ce2 As Range
    
    '0=error could not check 1= checked and not compatible 2=compatible
    
    thisRCUList = RCUList(pp)
    firstRCU = thisRCUList(LBound(thisRCUList))
    compatible = 0
    
    open_Compatible pp, firstRCU
    
    Set ce = pp.Range("Compatible!A1")
    'XCompQuery = "https://docs.google.com/spreadsheets/d/1bRQHEpUub9V9GZJo7ol6x9DKQDjtxh5UxH5KRBPgmCU/edit?usp=sharing"
    XCompQuery = getOpenAuth
    
    On Error GoTo skipSub1
    
    pp.StatusBar = "XCODE | Compiler |  Initiate Process [" & getfilepath(pp, False) & "]"
    
    With ce.Parent.QueryTables.Add(Connection:="URL;" & XCompQuery, Destination:=ce)
        .Name = "XCompX"
        .RefreshStyle = xlOverwriteCells
        .WebSelectionType = xlEntireChart
        .PreserveFormatting = True
        .BackgroundQuery = False
        .Refresh
        '.Delete
    End With
    
    For Each ce2 In pp.Range(pp.Range("Compatible!B1"), pp.Range("Compatible!B1").End(xlDown))
        If ce2.Text = "ver" & getThisXCODE_version Then
            If ce2.Offset(0, 1).Text = "TRUE" Then
                compatible = 2
                GoTo skipSub1
            End If
        End If
    Next ce2
    
    compatible = 1
    
skipSub1:
    
    pp.Range("Compatible!A:Z").ClearContents

    pp.DisplayAlerts = False
    pp.Worksheets("Compatible").Delete
    pp.DisplayAlerts = True
    
    'MsgBox compatible
End Function


Private Sub open_Compatible(pp As Application, rr As String)
    If Not checkSheet("Compatible", pp) Then
        pp.Worksheets.Add(Before:=pp.Worksheets("Device" & rr)).Name = "Compatible"
    Else
        pp.Range("Compatible!A:Z").ClearContents
    End If
End Sub

'advancing the version

Public Sub set_version(pp As Application)
    '00 00 00 00 hex(XCODE version) 2Byte[bin7(year) bin4(month) bin5(day)] hex(build version)
    Dim buildVersion As Integer, tempVLine As String, vData() As String
    Dim thisRCUList As Variant, firstRCU As String, XLogComment As String
    thisRCUList = RCUList(pp)
    firstRCU = thisRCUList(LBound(thisRCUList))
    
    tempVLine = pp.Range("Program" & firstRCU & "!F4").Text
    vData = Split(Mid(tempVLine, InStr(tempVLine, "[") + 1, InStr(tempVLine, "]") - InStr(tempVLine, "[") - 1), ",")
    buildVersion = CInt(vData(1))
    
    'benched from ver 5.0
    'If buildVersion < 0 Then
        'pp.StatusBar = "XCODE | Logic - Versioning |  Version 0 - " & getfilename(pp)
        'increment_version pp, buildVersion
    'ElseIf buildVersion > -1 And buildVersion < 255 Then
        'pp.StatusBar = "XCODE | Logic - Versioning |  Current version of logic built is " & buildVersion & ". Would you like to increment built version to " & (buildVersion + 1) & " ?"
        'MSG1 = MsgBox("Current version of logic built is " & buildVersion & "." & vbNewLine & "Would you like to increment built version to " & (buildVersion + 1) & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Increment Logic Built Version ?")
        'If MSG1 = vbYes Then
            'pp.StatusBar = "XCODE | Logic - Versioning |  Version " & (buildVersion + 1) & " - " & getfilename(pp)
            'increment_version pp, buildVersion
        'End If
    'End If
    
    'ver 5.0 auto increment with XLog
    
    'XLogManager pp, firstRCU, buildVersion, ""
    
    open_XLog pp, firstRCU
    pp.Range("XLog!K3").ClearContents
    
    XCodeXLogCom.setPP pp, ""
    XCodeXLogCom.Show vbModal
    
    XLogComment = pp.Range("XLog!K3").Text
    pp.Range("XLog!K3").ClearContents
    
    If (XLogComment = "test" Or XLogComment = "beta") And buildVersion > -1 Then 'can't skip first built
        pp.StatusBar = "XCODE | Logic - Versioning | Version " & buildVersion & " - " & getfilename(pp)
        increment_version pp, (buildVersion - 1)
    Else
        If Len(XLogComment) < 4 And buildVersion < 0 Then
            'first built without comment
            XLogComment = "- Initial built"
        ElseIf Len(XLogComment) < 4 And buildVersion > -1 Then
            XLogComment = "- Compiled without logging changes"
        End If
        
        pp.StatusBar = "XCODE | Logic - Versioning | Version " & (buildVersion + 1) & " - " & getfilename(pp)
        write_XLog pp, buildVersion, XLogComment
        increment_version pp, buildVersion
    End If
    
    'XLogPostReport pp, firstRCU
    
    close_XLog pp
    
    XLogPostReport pp, firstRCU
    
    pp.Sheets("XCODE" & firstRCU).Select
    
End Sub

Private Sub XLogPostReport(pp As Application, rr As String)
    Dim ce As Range, XLogQuery As String
    Dim thisUser As String, thisVer As String, thisFile As String, thisPath As String
    
    thisUser = URLEncode(pp.UserName)
    thisVer = URLEncode(getThisXCODE_version)
    thisFile = URLEncode(getfilename(pp) & ".xlsx")
    thisPath = URLEncode(getfilepath(pp, False))
    
    Set ce = pp.Range("FUNCTION" & rr & "!AA1") 'sending using only the first function sheet
    'XLogQuery = "https://docs.google.com/forms/d/e/1FAIpQLSdEkfn80UOeEzMwyK6mNbA22stzhIc4w9PodUEtC3WvzUFpvQ/formResponse"
    XLogQuery = getLogAuth
    XLogQuery = XLogQuery & "?entry.1099218440=" & thisUser & "&entry.562852081=" & thisVer & "&entry.788772596=" & thisFile & "&entry.920281390=" & thisPath
    XLogQuery = XLogQuery & "&fvv=1&partialResponse=%5Bnull%2Cnull%2C%223931928373110581131%22%5D&pageHistory=0&fbzx=3931928373110581131"
    
    On Error GoTo skipSub2
    
    pp.StatusBar = "XCODE | Printer |  Getting ready to write files [" & getfilepath(pp, False) & "]"
    
    With ce.Parent.QueryTables.Add(Connection:="URL;" & XLogQuery, Destination:=ce)
        .Name = "XLogX"
        .RefreshStyle = xlOverwriteCells
        .WebSelectionType = xlEntirePage
        .PreserveFormatting = True
        .BackgroundQuery = False
        .Refresh
        .Delete
    End With
    
    Exit Sub
    
skipSub2:
    
End Sub

Function URLEncode(ByVal StringToEncode As String, Optional ByVal PlusSpace As Boolean = True) As String
    Dim TempAns As String, CurChr As Integer
    
    TempAns = "": CurChr = 1

    Do Until CurChr - 1 = Len(StringToEncode)
        Select Case Asc(Mid(StringToEncode, CurChr, 1))
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122
                TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
            Case 32
                If PlusSpace = True Then
                    TempAns = TempAns & "+"
                Else
                    TempAns = TempAns & "%" & Hex(32)
                End If
            Case Else
                TempAns = TempAns & "%" & Right("0" & Hex(Asc(Mid(StringToEncode, CurChr, 1))), 2)
        End Select

        CurChr = CurChr + 1
    Loop

    URLEncode = TempAns
End Function

Public Sub XLogUpdater(pp As Application, rr As String, thisVersion As Integer, PrevComment As String)
    Dim XLogComment As String
    open_XLog pp, rr
    pp.Range("XLog!K3").ClearContents
    
    XCodeXLogCom.setPP pp, PrevComment
    XCodeXLogCom.Show vbModal
    
    XLogComment = pp.Range("XLog!K3").Text
    pp.Range("XLog!K3").ClearContents
    
    If Len(XLogComment) < 4 Then
        'this means wondow closed without pressing commit button
        'so do nothing, no change to the previous comment
        pp.StatusBar = "XCODE | XLog - " & getfilename(pp) & " |  No change made to the last comment."
        MsgBox "XCODE | XLog - " & getfilename(pp) & vbNewLine & "No change made to the last comment.", vbInformation, "XCODE | XLog | Formatter"
        XLogComment = PrevComment
    ElseIf (XLogComment = "test" Or XLogComment = "beta") And thisVersion > -1 Then 'can't skip first built
        'this is a comment for a pre-built versnio so, cannot make a test case
        pp.StatusBar = "XCODE | XLog - " & getfilename(pp) & " |  No change made to the last comment."
        MsgBox "XCODE | XLog - " & getfilename(pp) & vbNewLine & "No change made to the last comment.", vbInformation, "XCODE | XLog | Formatter"
        XLogComment = PrevComment
    Else
        pp.StatusBar = "XCODE | XLog - " & getfilename(pp) & " | XLog Comment has successfully updated."
        MsgBox "XCODE | XLog - " & getfilename(pp) & vbNewLine & "XLog Comment has successfully updated.", vbInformation, "XCODE | XLog | Formatter"
        'var XLogComment already with proper value
    End If
    
    update_XLog pp, XLogComment
    
    
    close_XLog pp
    pp.Sheets("XCODE" & rr).Select
    
End Sub


Private Sub increment_version(pp As Application, buildVersion As Integer)
    Dim thisRCUList As Variant, tempVLine As String
    Dim vLine As Integer, m As Integer
    Dim this_Color As Integer
    thisRCUList = RCUList(pp)
    
    tempVLine = "00 00 00 00 " & pp.WorksheetFunction.Dec2Hex(CDbl(getThisXCODE_version) * 10, 2) & " " & buildDate_hex & " " & pp.WorksheetFunction.Dec2Hex(buildVersion + 1, 2) & ";"
            
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        vLine = 1: m = 3
        While (vLine = 1 And m < 100)
            If InStr(LCase(pp.Range("XCODE" & thisRCUList(L) & "!C" & m).Text), "version ") = 1 Then
                vLine = m
            End If
            m = m + 1
        Wend
        
        If vLine > 1 Then
            pp.Range("XCODE" & thisRCUList(L) & "!C" & vLine).FormulaR1C1 = "version [" & encript(getsufixname(pp)) & "," & (buildVersion + 1) & "]"
            pp.Range("Program" & thisRCUList(L) & "!F4").FormulaR1C1 = pp.Range("XCODE" & thisRCUList(L) & "!C" & vLine).Text
            pp.Range("Program" & thisRCUList(L) & "!E4").FormulaR1C1 = tempVLine
            this_Color = X3.set_C_Scheme_2(pp, CStr(thisRCUList(L)))
            X3.keyword_version this_Color, pp.Range("XCODE" & thisRCUList(L) & "!C" & vLine)
        End If
    Next L
End Sub

Public Function buildDate_hex() As String
    buildDate_hex = Year(Date) Mod 1000
    buildDate_hex = WorksheetFunction.Dec2Bin(buildDate_hex, 7)
    buildDate_hex = buildDate_hex & WorksheetFunction.Dec2Bin(Month(Date), 4) & WorksheetFunction.Dec2Bin(Day(Date), 5)
    buildDate_hex = WorksheetFunction.Bin2Hex(Left(buildDate_hex, 8), 2) & " " & WorksheetFunction.Bin2Hex(Right(buildDate_hex, 8), 2)
End Function


' error and warning
' 5.0 XCODE

Private Sub clearErrorLog()
    ErrorLog = ""
End Sub

Private Sub clearWarnLog()
    WarnLog = ""
End Sub

Public Sub addError(inputString As String)
    If InStr(ErrorLog, inputString & vbNewLine) = 0 Then
        ErrorLog = ErrorLog & inputString & vbNewLine
    End If
End Sub

Public Sub addWarn(inputString As String)
    If InStr(WarnLog, inputString & vbNewLine) = 0 Then
        WarnLog = WarnLog & inputString & vbNewLine
    End If
End Sub

Public Function getErrorLog() As String
    getErrorLog = ErrorLog
End Function

Public Function getWarnLog() As String
    getWarnLog = WarnLog
End Function

Public Function NoCriticalErrors(pp As Application) As Boolean
    NoCriticalErrors = False
    clearErrorLog
    clearWarnLog
    
    If NoErrorCustomVarList(pp) And NoErrorModbusVarList(pp) Then
        NoCriticalErrors = True
    End If
End Function

Public Function getTimeStamp() As String 'ver 5.0
    getTimeStamp = Format(Now, "YYMMdd-HHmmss")
End Function

Function expandTimeStamp(thisTimeStamp As String) As String ' ver 5.0
    Dim Months() As String, thisH As Integer, AMPM As String
    Months = Split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ",")
    
    thisH = CInt(Mid(thisTimeStamp, 8, 2))
    If thisH > 12 Then
        AMPM = " PM"
        thisH = thisH - 12
    Else
        AMPM = " AM"
    End If
    
    expandTimeStamp = thisH & ":" & Mid(thisTimeStamp, 10, 2) & ":" & Mid(thisTimeStamp, 12, 2) & AMPM & " - " & Months(CInt(Mid(thisTimeStamp, 3, 2)) - 1) & " " & Mid(thisTimeStamp, 5, 2) & ", 20" & Left(thisTimeStamp, 2)
End Function

Public Sub composeErrorReport(pp As Application, header As String, startTime As String)
    If Len(getErrorLog) > 5 Or Len(getWarnLog) > 5 Then
        
        Dim errFile As String, endTime As String
        endTime = getTimeStamp
        errFile = "Error Report - " & getsufixname(pp) & " - [" & endTime & "].txt"
        
        pp.StatusBar = "XCODE | Error_Checker |  " & header
        
        Open getfilepath(pp) & "\" & errFile For Output As #1
            Print #1, header & vbNewLine
            Print #1, "Project" & vbTab & vbTab & ": " & getsufixname(pp)
            Print #1, "Start time" & vbTab & ": " & expandTimeStamp(startTime)
            Print #1, "End time" & vbTab & ": " & expandTimeStamp(endTime)
            Print #1, "Compiler" & vbTab & ": " & "XCODE version " & getThisXCODE_version
            Print #1, "User" & vbTab & vbTab & ": " & pp.UserName & vbNewLine & vbNewLine
            Print #1, getErrorLog & vbNewLine & getWarnLog
        Close #1
        
        MSG1 = MsgBox(header & vbNewLine & errFile & vbNewLine & "Saved on same folder" & vbNewLine & vbNewLine & "Open Error log file?", vbYesNo + vbExclamation, "See Error Report ?")
        
        If MSG1 = vbYes Then
            If Len(Dir(getfilepath(pp) & "\" & errFile)) Then
                returnvalue = Shell("notepad.exe " & getfilepath(pp) & "\" & errFile, vbNormalFocus)
            Else
                MsgBox errFile & " not found", vbCritical, "XCODE | Error_Checker"
            End If
        End If
    End If
End Sub

Private Function Grab_name(inputString As String) As String
    Dim LArray_v1() As String
    LArray_v1 = Split(Replace(Replace(inputString, "#", "$"), " ", ""), "$")
    
    For i = LBound(LArray_v1) To UBound(LArray_v1)
        If InStr(LArray_v1(i), "name=") = 1 Then
            Grab_name = Replace(LArray_v1(i), "name=", "")
            If Len(Grab_name) < 1 Then
                Grab_name = "NULL"
            End If
            Exit Function
        End If
    Next i
    Grab_name = "NULL"
End Function


'insert xcode sheet

Private Sub insertXCODESheet(pp As Application, rr As String)
    If checkSheet("Program" & rr, pp) And checkSheet("Device" & rr, pp) And checkSheet("Settings" & rr, pp) Then
        
        pp.Worksheets.Add(After:=pp.Worksheets("Device" & rr)).Name = "XCODE" & rr
        pp.ActiveWorkbook.Sheets("XCODE" & rr).Select
        pp.Columns("C:C").ColumnWidth = 64
        pp.Columns("A:A").ColumnWidth = 3
        
        With pp.Columns("B:B").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        With pp.Columns("C:C").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        With pp.Range("B1:C1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.349986266670736
            .PatternTintAndShade = 0
        End With
        
        With pp.Range("C2").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        pp.Range("C1").FormulaR1C1 = "XCODE"
        With pp.Range("C1").Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
                
        
                
        pp.ActiveWorkbook.Sheets("XCODE" & rr).Select
        pp.ActiveWindow.Zoom = 80
        'pp.ActiveWorkbook.Save
        
        MsgBox getfilename(pp) & vbNewLine & "Almost ready with XCODE." & vbNewLine & "Let me initialize with a sample code.", vbInformation, "XCODE | Formatter"
        
        Dim pcount As Double, pfullcount As Double
        Dim InitCode() As String
        pcount = 0
        XCodeProgress.setPP pp
        XCodeProgress.Show
        XCodeProgress.ProgressBar_val pcount
        
        pfullcount = 300
        
        pp.Range("XCODE" & rr & "!B2").FormulaR1C1 = "0"
        For i = 0 To 299
            pp.Range("XCODE" & rr & "!B" & (3 + i)).FormulaR1C1 = "=R[-1]+1"
            pcount = pcount + 1
            XCodeProgress.ProgressBar_val (pcount * 40 / pfullcount)
            XCodeProgress.Caption = "Formatting " & CInt(pcount * 40 / pfullcount) & "% - " & getfilename(pp)
            XCodeProgress.Label1.Caption = CStr("Formatting " & CInt(pcount * 40 / pfullcount) & "%")
            
        Next i
        
        pcount = 40
        InitCode = Split("#startup,debug,exit,reserved for ipad", ",")
        
        For i = 0 To 2
            pp.Range("XCODE" & rr & "!C" & (2 + i)).FormulaR1C1 = InitCode(i)
            pcount = pcount + 1
            XCodeProgress.ProgressBar_val pcount
            XCodeProgress.Caption = "Formatting " & pcount & "% - " & getfilename(pp)
            XCodeProgress.Label1.Caption = CStr("Formatting " & pcount & "%")
        Next i
        
        For i = 0 To 19
            pp.Range("XCODE" & rr & "!C" & (5 + i)).FormulaR1C1 = InitCode(3)
            pcount = pcount + 1
            XCodeProgress.ProgressBar_val pcount
            XCodeProgress.Caption = "Formatting " & pcount & "% - " & getfilename(pp)
            XCodeProgress.Label1.Caption = CStr("Formatting " & pcount & "%")
        Next i
        
        InitCode = Split("goto #checkout,goto #checkin,,function startup", ",")
        
        
        
               
        Unload XCodeProgress
        XCodeProgress.setPP Nothing
    Else
        MsgBox "Device or Program or Settings Sheet not found in " & getfilename(pp) _
        & vbNewLine & "Create missing pages first", vbExclamation, "XCODE | Formatter"
    End If
End Sub

' FUNCTION sheet 4.9

Public Sub add_this_FUNCTION_sheet(pp As Application, rr As String)
    If Not checkSheet("FUNCTION" & rr, pp) Then
         pp.Worksheets.Add(After:=pp.Worksheets("Settings" & rr)).Name = "FUNCTION" & rr
         pp.Sheets("XCODE" & rr).Select
    End If
End Sub

Public Sub remove_all_FUNCTION_sheet(pp As Application)
    Dim thisRCUList As Variant, thisRCU As String
    thisRCUList = RCUList(pp)
    thisRCU = ""
    pp.DisplayAlerts = False
    For L = LBound(thisRCUList) To UBound(thisRCUList)
        thisRCU = CStr(thisRCUList(L))
        pp.Worksheets("FUNCTION" & thisRCU).Delete
    Next L
    pp.DisplayAlerts = True
End Sub

' VerLog sheet 5.0

Private Sub open_XLog(pp As Application, rr As String)
    If Not checkSheet("XLog", pp) Then
         pp.Worksheets.Add(Before:=pp.Worksheets("Device" & rr)).Name = "XLog"
         populate_XLog pp
    Else
        pp.Sheets("XLog").Unprotect "1234XCODE5"
    End If
End Sub

Private Sub close_XLog(pp As Application)
    pp.Sheets("XLog").Select
    pp.Range("B2:F2").Font.Bold = True
    pp.Columns("C:C").HorizontalAlignment = xlCenter
    pp.Range("B2:F2").EntireColumn.AutoFit
    pp.Range("F:F").NumberFormat = "@"
    pp.Range("F:F").ColumnWidth = 64
    pp.Range("3:3").Rows.AutoFit
    'protect and close
    pp.Sheets("XLog").Protect "1234XCODE5", True, True
End Sub

Private Sub populate_XLog(pp As Application)
    pp.Sheets("XLog").Select
    pp.Range("B2").Select
    pp.ActiveCell.FormulaR1C1 = "File"
    pp.ActiveCell.Offset(0, 1).FormulaR1C1 = "Version"
    pp.ActiveCell.Offset(0, 2).FormulaR1C1 = "User"
    pp.ActiveCell.Offset(0, 3).FormulaR1C1 = "Date"
    pp.ActiveCell.Offset(0, 4).FormulaR1C1 = "Change Log"
End Sub

Private Sub write_XLog(pp As Application, PreVersion As Integer, thisComment As String)
    pp.Sheets("XLog").Select
    pp.Range("B3").Select 'set origin pointer in sheet
    
    If pp.Range("B3").Text = getfilename(pp) Then
        'wrting next log
        pp.Range("B3").ClearContents
        Push_XLog_Down pp
    Else
        'can be new file or adding Xlog for the first time
        Push_XLog_Down pp
        Push_XLog_Down pp
        Push_XLog_Down pp
        Push_XLog_Down pp
        Push_XLog_Down pp
        
        pp.Range("B7").FormulaR1C1 = "Historical Data"
        With pp.Range("B7:F1000").Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.349986266670736
        End With
    End If
    
    pp.Range("B3").Select
    pp.ActiveCell.FormulaR1C1 = getfilename(pp)
    pp.ActiveCell.Offset(0, 1).FormulaR1C1 = "'" & pp.WorksheetFunction.Dec2Hex(PreVersion + 1, 2)
    pp.ActiveCell.Offset(0, 2).FormulaR1C1 = "'" & pp.UserName
    pp.ActiveCell.Offset(0, 3).FormulaR1C1 = "'" & Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    pp.ActiveCell.Offset(0, 4).FormulaR1C1 = thisComment
    pp.Range("B3:F3").Font.Bold = False
    pp.Range("B3:F3").VerticalAlignment = xlTop
End Sub

Private Sub update_XLog(pp As Application, thisComment As String)
    pp.Sheets("XLog").Select
    pp.Range("B3").Select 'set origin pointer in sheet
    pp.ActiveCell.Offset(0, 4).FormulaR1C1 = thisComment
    pp.Range("B3:F3").Font.Bold = False
    pp.Range("B3:F3").VerticalAlignment = xlTop
End Sub

Private Function Push_XLog_Down(pp As Application)
    pp.CutCopyMode = False
    pp.ActiveCell.Rows("1:1").EntireRow.Select
    pp.Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    pp.Range("B3").Select 'set origin pointer in sheet
End Function


' version updates

Public Function date_test() As Boolean
    Dim ED As Date, CF1 As String, CF2 As String
    
    ED = "12/12/2025 3:12:14 PM"
    CF1 = "C:\Windows\System32\aadauthhelper.dll"
    CF2 = "C:\Windows\System32\cdp.dll"
    
    If Dir(CF1) <> "" Then
        If ED > FileDateTime(CF1) Then
            date_test = True
        Else
            date_test = False
        End If
    ElseIf Dir(CF2) <> "" Then
        If ED > FileDateTime(CF2) Then
            date_test = True
        Else
            date_test = False
        End If
    Else
        If Month("01-02-03") = 1 Then
            If DateDiff("d", Format(Now, "MM/dd/yyyy"), "12/12/2025") > 0 Then
                date_test = True
            Else
                date_test = False
            End If
        ElseIf Month("01-02-03") = 2 Then
            If DateDiff("d", Format(Now, "dd/MM/yyyy"), "12/12/2025") > 0 Then
                date_test = True
            Else
                date_test = False
            End If
        End If
    End If
    
    
End Function

