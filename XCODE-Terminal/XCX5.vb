Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX5

    Function Set_C_Scheme_2(thisConfigLine As String) As Integer
        Dim XCODEConfig As String, XConfigArr() As String


        XCODEConfig = thisConfigLine.Trim().ToLower()
        If XCODEConfig.Contains("]") AndAlso XCODEConfig.Contains("[") Then
            XCODEConfig = Strings.Mid(XCODEConfig, Strings.InStr(XCODEConfig, "[") + 1, Strings.InStr(XCODEConfig, "]") - Strings.InStr(XCODEConfig, "[") - 1)
            XConfigArr = XCODEConfig.Split(","c)
            If XConfigArr.Length > 1 Then
                If XConfigArr(1) = "1" Then
                    Return 1
                ElseIf XConfigArr(1) = "2" Then
                    Return 2
                End If
            End If
        End If

        Return 0
    End Function

    Sub Keyword_version(CC As Integer, ByRef inputRng As Excel.Range)
        C_func_mains(CC, inputRng, 1, 7)
        C_func_hides(CC, inputRng, 8, Strings.Len(inputRng.Text))
    End Sub

    Private Sub C_func_mains(CC As Integer, ByRef inputRng As Excel.Range, StartLocation As Integer, ofLength As Integer)
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

    Private Sub C_func_hides(CC As Integer, ByRef inputRng As Excel.Range, StartLocation As Integer, ofLength As Integer)
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
