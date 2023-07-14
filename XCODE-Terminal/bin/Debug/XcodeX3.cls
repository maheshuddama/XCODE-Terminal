VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XcodeX3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim X1 As New XcodeX1

Public Sub Change_Color_ThisXcode_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        XCodeColorMode.setPP pp
        XCodeColorMode.Show
    End If
End Sub

Public Sub Color_ThisXcode_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim this_Color As Integer
        this_Color = set_C_Scheme(pp)
        
        If this_Color = 1 Or this_Color = 2 Then
           Dim ce As Range, tempString As String, pcount As Double, pfullcount As Double
           pp.Range("C2").Font.FontStyle = "Bold"
           
            XCodeProgress.setPP pp
            XCodeProgress.Show
            XCodeProgress.ProgressBar_val pcount
            pfullcount = pp.Range(pp.Range("B3"), pp.Range("B2").End(xlDown)).Count
           
           For Each ce In pp.Range(pp.Range("B3"), pp.Range("B2").End(xlDown))
               tempString = pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text)
               If Len(tempString) > 0 Then
                   If tempString <> ce.Offset(0, 1).Text Then
                       ce.Offset(0, 1).FormulaR1C1 = tempString
                   End If
               
                   If InStr(ce.Offset(0, 1).Text, "//") = 1 Then
                       has_full_comment this_Color, ce.Offset(0, 1)
                   Else
                        If InStr(LCase(ce.Offset(0, 1).Text), "function ") = 1 Then
                           keyword_function this_Color, ce.Offset(0, 1)
                           bold_function this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "exit") = 1 Then
                           keyword_exit this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "0x") = 1 Then
                           direct_hex this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "debug") = 1 Then
                           debugMode this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "version") = 1 Then
                           keyword_version this_Color, ce.Offset(0, 1)
                        Else
                           If InStr(LCase(ce.Offset(0, 1).Text), "#exit ") Then
                               keyword_exit this_Color, ce.Offset(0, 1)
                           End If
                           
                           If InStr(LCase(ce.Offset(0, 1).Text), "goto ") = 1 Then
                               keyword_goto this_Color, ce.Offset(0, 1)
                           ElseIf InStr(LCase(ce.Offset(0, 1).Text), "if ") = 1 Then
                               keyword_if_condition this_Color, ce.Offset(0, 1)
                           ElseIf InStr(LCase(ce.Offset(0, 1).Text), "define ") = 1 Then
                               keyword_define this_Color, ce.Offset(0, 1)
                           End If
                           
                           If InStr(ce.Offset(0, 1).Text, " #") > 1 And (Not InStr(LCase(ce.Offset(0, 1).Text), " #exit")) Then
                               bold_usage_function this_Color, ce.Offset(0, 1)
                               other_rcu_call this_Color, ce.Offset(0, 1)
                           End If
                           
                           Color_Parameters this_Color, ce.Offset(0, 1)
                           keyword_other this_Color, ce.Offset(0, 1)
                        End If
            
                       If InStr(ce.Offset(0, 1).Text, "//") > 2 Then
                           has_half_comment this_Color, ce.Offset(0, 1)
                       End If
                   End If
               End If
                pcount = pcount + 1
                XCodeProgress.ProgressBar_val (pcount * 100 / pfullcount)
                XCodeProgress.Caption = "Formatting " & CInt(pcount * 100 / pfullcount) & "%"
                XCodeProgress.Label1.Caption = ce.Offset(0, 1).Text
           Next ce
           
            Unload XCodeProgress
            XCodeProgress.setPP Nothing
        End If
    End If
End Sub


'same as <Color_ThisXcode_Function> without progressbar
Public Sub Refresh_ThisXcode_Function(pp As Application)
    If X1.pointToXCode2(pp) Then
        Dim this_Color As Integer
        this_Color = set_C_Scheme(pp)
        pp.ScreenUpdating = False
        If this_Color = 1 Or this_Color = 2 Then
           Dim ce As Range, tempString As String
           pp.Range("C2").Font.FontStyle = "Bold"
           
           For Each ce In pp.Range(pp.Range("B3"), pp.Range("B2").End(xlDown))
                'Color_ThisLine_Function pp, ce.Offset(0, 1)
               tempString = pp.WorksheetFunction.Trim(ce.Offset(0, 1).Text)
               If Len(tempString) > 0 Then
                   If tempString <> ce.Offset(0, 1).Text Then
                       ce.Offset(0, 1).FormulaR1C1 = tempString
                   End If
               
                   If InStr(ce.Offset(0, 1).Text, "//") = 1 Then
                       has_full_comment this_Color, ce.Offset(0, 1)
                   Else
                        If InStr(LCase(ce.Offset(0, 1).Text), "function ") = 1 Then
                           keyword_function this_Color, ce.Offset(0, 1)
                           bold_function this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "exit") = 1 Then
                           keyword_exit this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "0x") = 1 Then
                           direct_hex this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "debug") = 1 Then
                           debugMode this_Color, ce.Offset(0, 1)
                        ElseIf InStr(LCase(ce.Offset(0, 1).Text), "version") = 1 Then
                           keyword_version this_Color, ce.Offset(0, 1)
                        Else
                           If InStr(LCase(ce.Offset(0, 1).Text), "#exit ") Then
                               keyword_exit this_Color, ce.Offset(0, 1)
                           End If
                           
                           If InStr(LCase(ce.Offset(0, 1).Text), "goto ") = 1 Then
                               keyword_goto this_Color, ce.Offset(0, 1)
                           ElseIf InStr(LCase(ce.Offset(0, 1).Text), "if ") = 1 Then
                               keyword_if_condition this_Color, ce.Offset(0, 1)
                           ElseIf InStr(LCase(ce.Offset(0, 1).Text), "define ") = 1 Then
                               keyword_define this_Color, ce.Offset(0, 1)
                           End If
                           
                           If InStr(ce.Offset(0, 1).Text, " #") > 1 And (Not InStr(LCase(ce.Offset(0, 1).Text), " #exit")) Then
                               bold_usage_function this_Color, ce.Offset(0, 1)
                               other_rcu_call this_Color, ce.Offset(0, 1)
                           End If
                           
                           Color_Parameters this_Color, ce.Offset(0, 1)
                           keyword_other this_Color, ce.Offset(0, 1)
                        End If
            
                       If InStr(ce.Offset(0, 1).Text, "//") > 2 Then
                           has_half_comment this_Color, ce.Offset(0, 1)
                       End If
                       
                   End If
               End If
           Next ce
           
        End If
        pp.ScreenUpdating = True
    End If
End Sub


Public Sub Color_ThisRange(pp As Application, rng As Range)
    Dim rowCount As Integer, thisRow As Integer
    
    rowCount = rng.Rows.Count
    
    For thisRow = 1 To rowCount
        Color_ThisLine_Function pp, Range(rng.Cells(thisRow, 1).Address)
    Next thisRow
    
End Sub



Public Sub Color_ThisLine_Function(pp As Application, rng As Range)
    If X1.pointToXCode2(pp) Then
        Dim this_Color As Integer
        this_Color = set_C_Scheme_1(pp)
        
        If this_Color = 1 Or this_Color = 2 Then
           Dim tempString As String
           
            tempString = pp.WorksheetFunction.Trim(rng.Text)
            If Len(tempString) > 0 Then
                If tempString <> rng.Text Then
                    rng.FormulaR1C1 = tempString
                End If
            
                If InStr(rng.Text, "//") = 1 Then
                    has_full_comment this_Color, rng
                Else
                     If InStr(LCase(rng.Text), "function ") = 1 Then
                        keyword_function this_Color, rng
                        bold_function this_Color, rng
                     ElseIf InStr(LCase(rng.Text), "exit") = 1 Then
                        keyword_exit this_Color, rng
                     ElseIf InStr(LCase(rng.Text), "0x") = 1 Then
                        direct_hex this_Color, rng
                     ElseIf InStr(LCase(rng.Text), "debug") = 1 Then
                        debugMode this_Color, rng
                     ElseIf InStr(LCase(rng.Text), "version") = 1 Then
                        keyword_version this_Color, rng
                     Else
                        If InStr(LCase(rng.Text), "#exit ") Then
                            keyword_exit this_Color, rng
                        End If
                        
                        If InStr(LCase(rng.Text), "goto ") = 1 Then
                            keyword_goto this_Color, rng
                        ElseIf InStr(LCase(rng.Text), "if ") = 1 Then
                            keyword_if_condition this_Color, rng
                        ElseIf InStr(LCase(rng.Text), "define ") = 1 Then
                            keyword_define this_Color, rng
                        End If
                        
                        If InStr(rng.Text, " #") > 1 And (Not InStr(LCase(rng.Text), " #exit")) Then
                            bold_usage_function this_Color, rng
                            other_rcu_call this_Color, rng
                        End If
                        
                        Color_Parameters this_Color, rng
                        keyword_other this_Color, rng
                     End If
         
                    If InStr(rng.Text, "//") > 2 Then
                        has_half_comment this_Color, rng
                    End If
                End If
            End If
            
        End If
    End If
End Sub


Private Function set_C_Scheme(pp As Application) As Integer
    Dim XCODEConfig As String, XConfigArr() As String
    set_C_Scheme = 0
    XCODEConfig = LCase(pp.WorksheetFunction.Trim(pp.Range("C1").Text))
    If InStr(XCODEConfig, "]") > 0 And InStr(XCODEConfig, "[") > 0 Then
        XCODEConfig = Mid(XCODEConfig, InStr(XCODEConfig, "[") + 1, InStr(XCODEConfig, "]") - InStr(XCODEConfig, "[") - 1)
        XConfigArr = Split(XCODEConfig, ",")
        If UBound(XConfigArr) > 0 Then
            If XConfigArr(1) = "1" Then
                Color_Scheme_Jedai pp
                set_C_Scheme = 1
            ElseIf XConfigArr(1) = "2" Then
                Color_Scheme_dath pp
                set_C_Scheme = 2
            End If
        End If
    End If
End Function

Private Function set_C_Scheme_1(pp As Application) As Integer
    Dim XCODEConfig As String, XConfigArr() As String
    set_C_Scheme_1 = 0
    XCODEConfig = LCase(pp.WorksheetFunction.Trim(pp.Range("C1").Text))
    If InStr(XCODEConfig, "]") > 0 And InStr(XCODEConfig, "[") > 0 Then
        XCODEConfig = Mid(XCODEConfig, InStr(XCODEConfig, "[") + 1, InStr(XCODEConfig, "]") - InStr(XCODEConfig, "[") - 1)
        XConfigArr = Split(XCODEConfig, ",")
        If UBound(XConfigArr) > 0 Then
            If XConfigArr(1) = "1" Then
                Color_Scheme_Jedai_1 pp
                set_C_Scheme_1 = 1
            ElseIf XConfigArr(1) = "2" Then
                Color_Scheme_dath_1 pp
                set_C_Scheme_1 = 2
            End If
        End If
    End If
End Function

Function set_C_Scheme_2(pp As Application, thisRCU As String) As Integer
    Dim XCODEConfig As String, XConfigArr() As String
    set_C_Scheme_2 = 0
    XCODEConfig = LCase(pp.WorksheetFunction.Trim(pp.Range("XCODE" & thisRCU & "!C1").Text))
    If InStr(XCODEConfig, "]") > 0 And InStr(XCODEConfig, "[") > 0 Then
        XCODEConfig = Mid(XCODEConfig, InStr(XCODEConfig, "[") + 1, InStr(XCODEConfig, "]") - InStr(XCODEConfig, "[") - 1)
        XConfigArr = Split(XCODEConfig, ",")
        If UBound(XConfigArr) > 0 Then
            If XConfigArr(1) = "1" Then
                set_C_Scheme_2 = 1
            ElseIf XConfigArr(1) = "2" Then
                set_C_Scheme_2 = 2
            End If
        End If
    End If
End Function


Private Sub Color_Scheme_dath(pp As Application)
    With pp.Range("C:C").Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
   
    With pp.Range("A:AZ").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With pp.Range("B:B").Font
        .Color = -6908266
        .TintAndShade = 0
    End With
    
    With pp.Range("B:B")
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With pp.Range("B:B").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    
End Sub

Private Sub Color_Scheme_dath_1(pp As Application)
    With pp.ActiveCell.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
   
    With pp.ActiveCell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Private Sub Color_Scheme_Jedai(pp As Application)
    With pp.Range("C:C").Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
        .Name = "Calibri"
    End With
   
    With pp.Range("A:AZ").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With pp.Range("A:B").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    With pp.Range("B:B")
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    With pp.Range("B:B").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub

Private Sub Color_Scheme_Jedai_1(pp As Application)
    With pp.ActiveCell.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Bold = False
        .Name = "Calibri"
    End With
   
    With pp.ActiveCell.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Private Sub keyword_function(CC As Integer, inputRng As Range)
    C_func_mains CC, inputRng, InStr(LCase(inputRng.Text), "function "), 8
End Sub

Private Sub keyword_exit(CC As Integer, inputRng As Range)
    C_func_mains CC, inputRng, InStr(LCase(inputRng.Text), "exit"), 4
End Sub

Private Sub keyword_goto(CC As Integer, inputRng As Range)
    C_func_mains CC, inputRng, 1, 4
End Sub

Private Sub direct_hex(CC As Integer, inputRng As Range)
    C_func_mains CC, inputRng, 1, 2
End Sub

Private Sub debugMode(CC As Integer, inputRng As Range)
    C_func_mains CC, inputRng, 1, 5
End Sub

Public Sub keyword_version(CC As Integer, inputRng As Range) ' not private to access 4.9
    C_func_mains CC, inputRng, 1, 7
    C_func_hides CC, inputRng, 8, Len(inputRng.Text)
End Sub

Private Sub keyword_define(CC As Integer, inputRng As Range)
    C_func_mains CC, inputRng, 1, 6
End Sub

Private Sub bold_function(CC As Integer, inputRng As Range)
    Dim thisLine() As String, funcLength As Integer
    thisLine = Split(inputRng.Text, " ")
    If UBound(thisLine) > 0 Then
        funcLength = Len(thisLine(1))
    End If
    C_User_func CC, inputRng, 10, funcLength
End Sub

Private Sub bold_usage_function(CC As Integer, inputRng As Range)
    Dim useFuncPos As Integer, nextHalf As String, endUseFunc As Integer
    useFuncPos = InStr(inputRng.Text, "#") + 1
    nextHalf = Mid(inputRng.Text, useFuncPos, Len(inputRng.Text))
    endUseFunc = InStr(nextHalf, " ") - 1
    C_User_func CC, inputRng, useFuncPos, endUseFunc
End Sub

Private Sub keyword_other(CC As Integer, inputRng As Range)
    If InStr(LCase(inputRng.Text), "rcu.") = 1 Or InStr(LCase(inputRng.Text), "math.") = 1 _
    Or InStr(LCase(inputRng.Text), "dmx.") = 1 Or InStr(LCase(inputRng.Text), "ioexp") = 1 _
    Or InStr(LCase(inputRng.Text), "send.") = 1 Or InStr(LCase(inputRng.Text), "read.") = 1 _
    Or InStr(LCase(inputRng.Text), "dali") = 1 Or InStr(LCase(inputRng.Text), "dimflow.") = 1 _
    Or InStr(LCase(inputRng.Text), "iodexp") = 1 Then
        C_other_key_bold CC, inputRng, 1, InStr(inputRng.Text, ".") - 1
    ElseIf InStr(LCase(inputRng.Text), "load ") = 1 Or InStr(LCase(inputRng.Text), "save ") = 1 _
    Or InStr(LCase(inputRng.Text), "run ") = 1 Or InStr(LCase(inputRng.Text), "reuse ") = 1 _
    Or InStr(LCase(inputRng.Text), "wait ") = 1 Then
        C_other_key_bold CC, inputRng, 1, InStr(inputRng.Text, " ") - 1
    End If
    
    If InStr(LCase(inputRng.Text), ".") > 1 Then
        Dim dotLocation As Integer
        dotLocation = InStr(LCase(inputRng.Text), ".")
        If InStr(LCase(inputRng.Text), "=") > 1 Then
            dotLocation = InStr(LCase(inputRng.Text), "=") - dotLocation
        End If
        If dotLocation > 1 Then
            C_other_key CC, inputRng, InStr(inputRng.Text, "."), InStr(inputRng.Text, " ") - InStr(inputRng.Text, ".")
        End If
    End If
    If InStr(LCase(inputRng.Text), "load dimmer ") = 1 Then
        C_other_key CC, inputRng, InStr(inputRng.Text, " dimmer ") + 1, 6
    ElseIf InStr(LCase(inputRng.Text), "load time ") = 1 Then
        C_other_key CC, inputRng, InStr(inputRng.Text, " time ") + 1, 4
    End If
End Sub

Private Sub other_rcu_call(CC As Integer, inputRng As Range)
    If InStr(inputRng.Text, " @rcu") > 1 Then
        C_Other_RCU CC, inputRng, InStr(inputRng.Text, " @rcu") + 1, 6
    End If
End Sub

Private Sub keyword_if_condition(CC As Integer, inputRng As Range)
    Dim end_cond As Integer
    If InStr(LCase(inputRng.Text), "if equal ") = 1 Then
        end_cond = 8
    ElseIf InStr(LCase(inputRng.Text), "if greater than ") = 1 Then
        end_cond = 15
    ElseIf InStr(LCase(inputRng.Text), "if greater ") = 1 Then
        end_cond = 10
    ElseIf InStr(LCase(inputRng.Text), "if less than ") = 1 Then
        end_cond = 12
    ElseIf InStr(LCase(inputRng.Text), "if not equal ") = 1 Then
        end_cond = 12
    ElseIf InStr(LCase(inputRng.Text), "if bit true ") = 1 Then
        end_cond = 11
    ElseIf InStr(LCase(inputRng.Text), "if bit run ") = 1 Then
        end_cond = 10
    End If
    
    C_Condition CC, inputRng, 1, end_cond
End Sub


Private Sub Color_Parameters(CC As Integer, inputRng As Range)
    Dim i As Integer
    
    If InStr(LCase(inputRng.Text), " $dec") Then
        i = InStr(LCase(inputRng.Text), "$dec")
        While i > 1
            C_Parameter CC, inputRng, i, 4
            i = InStr(i + 1, LCase(inputRng.Text), "$dec")
        Wend
    End If
    If InStr(LCase(inputRng.Text), " $hex") Then
        i = InStr(LCase(inputRng.Text), "$hex")
        While i > 1
            C_Parameter CC, inputRng, i, 4
            i = InStr(i + 1, LCase(inputRng.Text), "$hex")
        Wend
    End If
    If InStr(LCase(inputRng.Text), " $bin") Then
        i = InStr(LCase(inputRng.Text), "$bin")
        While i > 1
            C_Parameter CC, inputRng, i, 4
            i = InStr(i + 1, LCase(inputRng.Text), "$bin")
        Wend
    End If
    If InStr(LCase(inputRng.Text), " $relay") Then
        i = InStr(LCase(inputRng.Text), "$relay")
        While i > 1
            C_Parameter CC, inputRng, i, 6
            i = InStr(i + 1, LCase(inputRng.Text), "$relay")
        Wend
    End If
    If InStr(LCase(inputRng.Text), " $bit") Then
        i = InStr(LCase(inputRng.Text), "$bit")
        While i > 1
            C_Parameter CC, inputRng, i, 4
            i = InStr(i + 1, LCase(inputRng.Text), "$bit")
        Wend
    End If
    If InStr(LCase(inputRng.Text), " $fan") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$fan"), 4
    ElseIf InStr(LCase(inputRng.Text), " $display") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$display"), 8
    End If
    
    If InStr(LCase(inputRng.Text), " $type") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$type"), 5
    ElseIf InStr(LCase(inputRng.Text), " $dev") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$dev"), 4
    End If
    
    If InStr(LCase(inputRng.Text), " $sec") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$sec"), 4
    End If
    If InStr(LCase(inputRng.Text), " $min") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$min"), 4
    End If
    If InStr(LCase(inputRng.Text), " $hour") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$hour"), 5
    End If
    If InStr(LCase(inputRng.Text), " $lum") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$lum"), 4
    End If
    If InStr(LCase(inputRng.Text), " $len") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$len"), 4
    End If
    If InStr(LCase(inputRng.Text), " $reg") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$reg"), 4
    End If
    
    If InStr(LCase(inputRng.Text), " $ch") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$ch"), 3
    End If
    If InStr(LCase(inputRng.Text), " $ballast") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$ballast"), 8
    End If
    If InStr(LCase(inputRng.Text), " $group") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$group"), 6
    End If
    If InStr(LCase(inputRng.Text), " $bus") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$bus"), 4
    End If
    If InStr(LCase(inputRng.Text), " $milisec") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$milisec"), 8
    End If
    If InStr(LCase(inputRng.Text), " $ip") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$ip"), 3
    End If
    If InStr(LCase(inputRng.Text), " $name") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$name"), 5
    End If
    If InStr(LCase(inputRng.Text), " $clock") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$clock"), 6
    End If
    
    If InStr(LCase(inputRng.Text), " $var") Then
        i = InStr(LCase(inputRng.Text), " $var")
        While i > 1
            C_Parameter CC, inputRng, i, 4
            i = InStr(i + 1, LCase(inputRng.Text), "$var")
        Wend
    ElseIf InStr(LCase(inputRng.Text), " $element") Then
        C_Parameter CC, inputRng, InStr(LCase(inputRng.Text), "$element"), 8
    End If
End Sub

Private Sub has_full_comment(CC As Integer, inputRng As Range)
    C_Comment CC, inputRng, 1, Len(inputRng.Text)
End Sub

Private Sub has_half_comment(CC As Integer, inputRng As Range)
    C_Comment CC, inputRng, InStr(inputRng.Text, "//"), Len(inputRng.Text)
End Sub

'colors

Private Sub C_Comment(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -10526881
            .TintAndShade = 0
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -4144960
            .TintAndShade = 0
        End With
    End If
End Sub

Private Sub C_Parameter(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -6684928
            .TintAndShade = 0
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -16744448
            .TintAndShade = 0
        End With
    End If
End Sub

Private Sub C_Condition(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Bold"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = -26113
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Bold"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = -65281
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    End If
End Sub

Private Sub C_Other_RCU(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Normal"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = -16737793
            .TintAndShade = 0
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Normal"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = -16737793
            .TintAndShade = 0
        End With
    End If
End Sub

Private Sub C_func_mains(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Normal"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = -13261
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Normal"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = -3394816
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    End If
End Sub

Private Sub C_func_hides(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Normal"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = RGB(0, 0, 0)
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Normal"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .Color = RGB(255, 255, 255)
            .TintAndShade = 0
            .ThemeFont = xlThemeFontMinor
        End With
    End If
End Sub

Private Sub C_User_func(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Bold"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .FontStyle = "Bold"
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
        End With
    End If
End Sub

Private Sub C_other_key(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -10040065
            .TintAndShade = 0
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -16763956
            .TintAndShade = 0
        End With
    End If
End Sub

Private Sub C_other_key_bold(CC As Integer, inputRng As Range, StartLocation As Integer, ofLength As Integer)
    If CC = 2 Then
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -10040065
            .TintAndShade = 0
            .FontStyle = "Bold"
        End With
    Else
        With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
            .Color = -16763956
            .TintAndShade = 0
            .FontStyle = "Bold"
        End With
    End If
End Sub




