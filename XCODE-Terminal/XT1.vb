Imports System.ComponentModel

Module XT1
    Dim X0 As XCX0
    Dim ErrorWarnLog As String()
    Dim Compiling As Boolean

    Sub Main(ByVal args As String())
        X0 = New XCX0()
        Console.Title = "XC-Terminal Ver(" & X0.getXCVer & "." & X0.getXCbuildVer & ") - (For Preview Only)"
        Compiling = False
        Dim Input As String
        Dim xlPath = My.Computer.FileSystem.CurrentDirectory

        Console.Clear()
        X0.XC_Terminal_Header()

        If args.Length > 0 Then
            Input = "xc " & X0.get_memmap_name_from_words(args, 0)
            Console.WriteLine("XC:>" & Input)
        Else
            Input = ""
        End If


        While (Input.ToLower() <> "exit" And Not Compiling)

            If Input.Length = 0 Then
                Console.Write("XC:>")
                Input = Console.ReadLine()
            End If

            Dim words As String() = Strings.Split(Input, " ")

            If words(0).ToLower = "xc" And Not Compiling Then
                Input = ""
                ErrorWarnLog = {"", ""}
                Dim X1 As New XCX1
                Dim MemmapFileName As String

                If UBound(words) > 0 AndAlso words(1) <> "" Then
                    MemmapFileName = X0.get_memmap_name_from_words(words)

                    If X0.isAnExcelFile(MemmapFileName) Then
                        If X0.isFileExist(xlPath & "\" & MemmapFileName) Then
                            If Not X0.IsFileInUse(xlPath & "\" & MemmapFileName) Then
                                Compiling = True
                                X1.XCM(MemmapFileName, ErrorWarnLog)
                                Compiling = False
                            Else
                                Console.WriteLine("'" & MemmapFileName & "', file is in use or does not have access.")
                                Console.WriteLine("")
                            End If
                        Else
                            Console.WriteLine("'" & MemmapFileName & "', file does not exist.")
                            Console.WriteLine("")
                        End If
                    Else
                        Console.WriteLine("'" & MemmapFileName & "', does not match the required file extention of '.xlsx'.")
                        Console.WriteLine("")
                    End If
                Else
                    Console.WriteLine("Instruction incomplete! type 'help' on Console to know the format.")
                    Console.WriteLine("")
                End If
                releaseObject(X1)
            ElseIf words(0).ToLower = "help" Then
                Input = ""
                Console.WriteLine("")
                Console.WriteLine("XC [Filename]" & vbTab & vbTab & "-Compile the given file.")
                Console.WriteLine("Path" & vbTab & vbTab & vbTab & "-Path for Solution or Error Report.")
                Console.WriteLine("Version" & vbTab & vbTab & vbTab & "-Show my version. (You see it already)")
                Console.WriteLine("Help" & vbTab & vbTab & vbTab & "-Help menu. (How to see me again)")
                Console.WriteLine("CLS" & vbTab & vbTab & vbTab & "-Clear Screen.")
                Console.WriteLine("Exit" & vbTab & vbTab & vbTab & "-Exit XC-Terminal.")
                Console.WriteLine("")
            ElseIf words(0).ToLower = "path" Then
                Input = ""
                Console.WriteLine("Root folder is : [" & xlPath & "]")
                Console.WriteLine("")
            ElseIf words(0).ToLower = "version" Then
                Input = ""
                Console.WriteLine("My Version is XC-Terminal (" & X0.getXCVer & "." & X0.getXCbuildVer & ")")
                Console.WriteLine("")
            ElseIf words(0).ToLower = "cls" Then
                Input = ""
                Console.Clear()
                X0.XC_Terminal_Header()
            Else
                If Input = "" Then
                    Input = ""
                ElseIf Input <> "exit" Then 'this to capture when it runs from cmd and exit to cmd
                    Console.WriteLine("'" & Input & "' is not a recognized XC instruction.")
                    Console.WriteLine("")
                    Input = ""
                End If
            End If

        End While

    End Sub

    'garbage management

    Private Sub releaseObject(ByRef thisObject As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(thisObject)
            thisObject = Nothing
        Catch ex As Exception
            thisObject = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module
