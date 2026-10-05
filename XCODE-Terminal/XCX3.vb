Imports Excel = Microsoft.Office.Interop.Excel

Public Class XCX3
    Dim xlXCODESheet As Excel.Worksheet
    Dim xlProgramSheet As Excel.Worksheet

    Dim X0 As XCX0

    Sub XCODE_Core3(ByRef thisWorkBook As Excel.Workbook, ByRef thisRCUList As List(Of String), ByRef ErrorWarnLog As String(), unAttended As Boolean, ByRef IPDictionary As Dictionary(Of String, String))

        'initiate if there are more than 1 rcu in the memmap
        If thisRCUList.Count > 1 Then

            Dim RCUname As String, pcount As Integer, pfullcount As Integer
            Dim ThisTempLine As String

            X0 = New XCX0(unAttended)

            For Each thisRCU As String In thisRCUList

                RCUname = "RCU" & thisRCU & " > "
                X0.ConsoleMsg("XC Progress:> " & RCUname & "XCODE > Core - 3 > Mapping.......")
                xlProgramSheet = thisWorkBook.Sheets("Program" & thisRCU)

                pcount = 0
                pfullcount = xlProgramSheet.Range(xlProgramSheet.Range("D4"), xlProgramSheet.Range("D4").End(Excel.XlDirection.xlDown)).Count

                Dim ScripttargetRange As Excel.Range = xlProgramSheet.Range("F4:F" & (pfullcount + 3))
                Dim ScriptValues As Object(,) = DirectCast(ScripttargetRange.Value, Object(,))
                Dim rowCount As Integer = ScriptValues.GetLength(0)
                Dim thisRow As Integer

                'For Each ce3 In xlProgramSheet.Range(xlProgramSheet.Range("D4"), xlProgramSheet.Range("D4").End(Excel.XlDirection.xlDown))
                For r As Integer = 1 To rowCount

                    ThisTempLine = If(ScriptValues(r, 1)?.ToString(), String.Empty).Trim() 'already in lcase
                    thisRow = r + 3 'excel row

                    If X0.IsOtherRCUcall(ThisTempLine) Then
                        Dim otherRCU As String
                        otherRCU = GetOtherRCU(ThisTempLine, thisRCUList)

                        If otherRCU <> "NULL" Then

                            'Dim ThisTempLine2 As String
                            Dim otherRCUIP As String, otherRCUF As String, thisFunction As String, LArrayF() As String
                            'Dim ce4 As Excel.Range
                            Dim f As Integer
                            otherRCUIP = "00 00 00 00" : otherRCUF = "00 00"

                            'xlXCODESheet = thisWorkBook.Sheets("XCODE" & thisRCU)

                            'For Each ce4 In xlXCODESheet.Range(xlXCODESheet.Range("B2"), xlXCODESheet.Range("B2").End(Excel.XlDirection.xlDown))
                            '    ThisTempLine2 = ce4.Offset(0, 1).Text

                            '    If X0.isADefinition(ThisTempLine2) Then

                            '        If InStr(ThisTempLine2, "//") > 20 Then 'filter comment with minimum distance
                            '            ThisTempLine2 = Strings.Left(ThisTempLine2, Strings.InStr(ThisTempLine2, "//") - 1)
                            '        End If

                            '        If isAnIPdefinition(ThisTempLine2, otherRCU, thisRCU, ErrorWarnLog) Then
                            '            otherRCUIP = X0.Grab_IP_addr(X0.Grab_variant(ThisTempLine2, "ip"))
                            '            Exit For
                            '        End If
                            '    End If
                            'Next ce4

                            If IPDictionary.ContainsKey("XCODE" & otherRCU) Then
                                otherRCUIP = X0.Grab_IP_addr(IPDictionary("XCODE" & otherRCU))
                            End If


                            LArrayF = ThisTempLine.Split(" "c)

                            For f = LBound(LArrayF) To UBound(LArrayF)
                                If LArrayF(f).StartsWith("#"c) Then
                                    thisFunction = LArrayF(f).Replace("#"c, "")
                                    otherRCUF = X0.GetLineFromFunctionNameOtherRCU(thisFunction, 3, thisWorkBook, otherRCU) '3 input line for nearest exit
                                    otherRCUF = otherRCUF.Substring(2, 2) & " " & otherRCUF.Substring(0, 2)
                                    Exit For 'skip checking other elements
                                End If
                            Next f

                            If otherRCUF <> "00 00" AndAlso otherRCUIP <> "00 00 00 00" Then
                                xlProgramSheet.Range("E" & thisRow).FormulaR1C1 = "30 " & otherRCUF & " 00 " & otherRCUIP & ";"
                            Else
                                If otherRCUF = "00 00" Then
                                    ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | " & ThisTempLine & "][Invalid function call at RCU" & otherRCU & "]")
                                End If
                                If otherRCUIP = "00 00 00 00" Then
                                    ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Request functions from RCU" & otherRCU & "][Missing RCU" & otherRCU & " definition]")
                                End If
                            End If

                        Else
                            ErrorWarnLog(0) = X0.AddErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | " & ThisTempLine & "][Invalid request of RCU]")
                        End If

                    End If

                    pcount += 1
                    X0.ConsoleProgress(RCUname & "XCODE > Core - 3 > Mapping", (pcount * 100 / pfullcount))

                Next

                Next


            ReleaseObject(xlXCODESheet)
            ReleaseObject(xlProgramSheet)
            ReleaseObject(X0)

        End If
    End Sub

    ' other rcu
    Function GetOtherRCU(inputString As String, ByRef thisRCUList As List(Of String)) As String
        Dim LArray_rcu() As String, tempS As String, k As Integer

        LArray_rcu = inputString.Split(" "c)
        GetOtherRCU = "NULL" : tempS = ""

        For k = LBound(LArray_rcu) To UBound(LArray_rcu)
            If LArray_rcu(k).StartsWith("@rcu") Then
                tempS = LArray_rcu(k).Replace("@rcu", "")
                Exit For
            End If
        Next k

        If IsNumeric(tempS) Then
            If CInt(tempS) > 0 AndAlso CInt(tempS) <= 64 Then
                If thisRCUList.Contains(tempS) Then
                    GetOtherRCU = CStr(CInt(tempS))
                End If
            End If
        End If

    End Function

    'ip definition
    'Private Function IsAnIPdefinition(inputString As String, otherRCU As String, thisRCU As String, ByRef ErrorWarnLog As String()) As Boolean
    '    Dim otherRCUIP As String
    '    otherRCUIP = X0.Grab_variant(inputString, "ip")

    '    If otherRCUIP <> "NULL" AndAlso InStr(inputString, "define ") = 1 AndAlso InStr(inputString, "rcu" & otherRCU & " ") > 1 Then
    '        If X0.IsAValidIP(otherRCUIP) Then
    '            Return True
    '        Else
    '            'not valid IP address
    '            ErrorWarnLog(0) = X0.addErrorOrWarn(ErrorWarnLog(0), "[RCU" & thisRCU & " | XCODE" & thisRCU & " | Definition of RCU" & otherRCU & " $ip=" & otherRCUIP & "][Invalid IP address]")
    '        End If
    '    End If
    '    Return False
    'End Function


    'garbage management
    Private Sub ReleaseObject(ByRef thisObject As Object)
        Try
            Runtime.InteropServices.Marshal.FinalReleaseComObject(thisObject)
            thisObject = Nothing
        Catch ex As Exception
            thisObject = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


End Class
