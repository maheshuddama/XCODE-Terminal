Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX5

    Function set_C_Scheme_2(ByVal thisConfigLine As String) As Integer
        Dim XCODEConfig As String, XConfigArr() As String
        set_C_Scheme_2 = 0
        XCODEConfig = Strings.LCase(Strings.Trim(thisConfigLine))
        If Strings.InStr(XCODEConfig, "]") > 0 AndAlso Strings.InStr(XCODEConfig, "[") > 0 Then
            XCODEConfig = Strings.Mid(XCODEConfig, Strings.InStr(XCODEConfig, "[") + 1, Strings.InStr(XCODEConfig, "]") - Strings.InStr(XCODEConfig, "[") - 1)
            XConfigArr = Strings.Split(XCODEConfig, ",")
            If UBound(XConfigArr) > 0 Then
                If XConfigArr(1) = "1" Then
                    set_C_Scheme_2 = 1
                ElseIf XConfigArr(1) = "2" Then
                    set_C_Scheme_2 = 2
                End If
            End If
        End If
    End Function

    Sub keyword_version(ByVal CC As Integer, ByRef inputRng As Excel.Range)
        C_func_mains(CC, inputRng, 1, 7)
        C_func_hides(CC, inputRng, 8, Strings.Len(inputRng.Text))
    End Sub

    Private Sub C_func_mains(ByVal CC As Integer, ByRef inputRng As Excel.Range, ByVal StartLocation As Integer, ByVal ofLength As Integer)
        If CC = 2 Then
            With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
                .FontStyle = "Normal"
                .OutlineFont = False
                .Shadow = False
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .Color = -13261
                .TintAndShade = 0
                .ThemeFont = Excel.XlThemeFont.xlThemeFontMinor
            End With
        Else
            With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
                .FontStyle = "Normal"
                .OutlineFont = False
                .Shadow = False
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .Color = -3394816
                .TintAndShade = 0
                .ThemeFont = Excel.XlThemeFont.xlThemeFontMinor
            End With
        End If
    End Sub

    Private Sub C_func_hides(ByVal CC As Integer, ByRef inputRng As Excel.Range, ByVal StartLocation As Integer, ByVal ofLength As Integer)
        If CC = 2 Then
            With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
                .FontStyle = "Normal"
                .OutlineFont = False
                .Shadow = False
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .Color = RGB(0, 0, 0)
                .TintAndShade = 0
                .ThemeFont = Excel.XlThemeFont.xlThemeFontMinor
            End With
        Else
            With inputRng.Characters(Start:=StartLocation, Length:=ofLength).Font
                .FontStyle = "Normal"
                .OutlineFont = False
                .Shadow = False
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
                .Color = RGB(255, 255, 255)
                .TintAndShade = 0
                .ThemeFont = Excel.XlThemeFont.xlThemeFontMinor
            End With
        End If
    End Sub

End Class
