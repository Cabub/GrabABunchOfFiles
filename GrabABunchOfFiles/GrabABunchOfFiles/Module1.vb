Imports System.Text.RegularExpressions
Imports System.IO
Module Module1

    Sub Main(ByVal Args() As String)
        Dim ArgsOffset As Integer = 0
        Dim bQuiet As Boolean = True
        Dim bPanelFile As Boolean = False
        Dim bSourceDirFile As Boolean = False
        Dim bNoCriteria As Boolean = False
        Dim bCopy As Boolean = True
        'For i As Integer = 0 To Args.Count - 1
        If Args.Count < 3 Then
            ShowHelp()
            Return
        End If
        If Args(0).StartsWith("-") Then
            ArgsOffset = 1
            For i As Integer = 1 To Args(0).Count - 1
                Select Case (Args(0).Chars(i))
                    Case "h"
                        ShowHelp()
                        Return
                    Case "n"
                        bNoCriteria = True
                        ArgsOffset = ArgsOffset - 1
                    Case "v"
                        bQuiet = False
                    Case "r"
                        bPanelFile = True
                    Case "s"
                        bSourceDirFile = True
                    Case "d"
                        bCopy = False
                    Case Else
                        ShowHelp()
                        Return
                End Select
            Next
        End If
        'Next

        Dim RegXFileCriteria As String
        If bNoCriteria Then
            RegXFileCriteria = "{r}"
        Else
            RegXFileCriteria = Args(0 + ArgsOffset)
        End If
        '= {5571, 5379, 8047, 1340, 8075, 5562, 2210, 8024, 8562, 8044, 8533, 8027,
        '5565, 5310, 8033, 8434, 5566, 8042, 7970, 8369, 8532, 8045, 8028, 2220,
        '8025, 8031, 8005, 6694, 8366, 8043, 8026, 8032, 7854, 5570, 5378, 8046,
        '8378, 5344, 8029}
        Dim DestDir As String = Args(1 + ArgsOffset)

        Dim SourceDir As New List(Of String) '= {"\\rptscheduler\archives\s2005\20151230", "\\rptscheduler\archives\s2005\20151231"}
        Dim Panel As New List(Of String)

        If Not Directory.Exists(DestDir) Then
            Console.Error.Write("ERROR: Destination directory " & DestDir & " cannot be found, quitting...")
            Return
        End If

        If bSourceDirFile Then
            If Not File.Exists(Args(2 + ArgsOffset)) Then
                Console.Error.Write("ERROR: Source directory file " & Args(2 + ArgsOffset) & " cannot be found, quitting...")
                Return
            Else
                Using fr As New StreamReader(Args(2 + ArgsOffset))
                    While Not fr.EndOfStream
                        SourceDir.Add(fr.ReadLine)
                    End While
                End Using
            End If
        Else
            SourceDir.Add(Args(2 + ArgsOffset))
        End If


        If Not bNoCriteria AndAlso bPanelFile Then
            If Not File.Exists(Args(3 + ArgsOffset)) Then
                Console.Error.Write("ERROR: Replacement file " & Args(3 + ArgsOffset) & " cannot be found, quitting...")
                Return
            Else
                Using fr As New StreamReader(Args(3 + ArgsOffset))
                    While Not fr.EndOfStream
                        Panel.Add(fr.ReadLine)
                    End While
                End Using
            End If
        ElseIf Not bNoCriteria Then
            For i As Integer = 3 + ArgsOffset To Args.Count - 1
                Panel.Add(Args(i))
            Next
        End If

        If Not DestDir.EndsWith("\") Then
            DestDir = DestDir & "\"
        End If
        If Not Directory.Exists(DestDir) Then
            Console.Error.Write("ERROR: Destination file " & DestDir & " cannot be found, quitting...")
            Return
        End If

        For Each direct As String In SourceDir
            If Not direct.EndsWith("\") Then
                direct = direct & "\"
            End If
            If Not Directory.Exists(direct) Then
                Console.Error.Write("ERROR: Source file " & direct & " cannot be found, quitting...")
                Return
            End If
        Next

        If Not bNoCriteria And RegXFileCriteria.Length > 0 And Panel.Count > 0 And Not RegXFileCriteria.Contains("{r}") Then
            Console.Error.Write("ERROR: Found replacement string(s), but criteria does not contain {r}. Why?")
        End If

        Try
            If bNoCriteria Then
                For Each srcDir As String In SourceDir
                    Dim dir As New DirectoryInfo(srcDir)
                    Dim files() As FileInfo = dir.GetFiles()
                    For Each file In files
                        If bCopy Then
                            If Not bQuiet Then
                                Console.WriteLine("COPYING " & file.FullName & vbTab & " to " & DestDir & file.Name)
                            End If
                            FileCopy(file.FullName, DestDir & file.Name)
                        Else
                            Console.WriteLine("FOUND " & file.FullName)
                        End If
                    Next
                Next
            ElseIf Panel.Count > 0 Then
                For Each p As String In Panel
                    Dim RegX As New Regex(RegXFileCriteria.Replace("{r}", p.ToString()), RegexOptions.IgnoreCase)
                    For Each srcDir As String In SourceDir
                        Dim dir As New DirectoryInfo(srcDir)
                        Dim files() As FileInfo = dir.GetFiles()
                        For Each file In files
                            If RegX.IsMatch(file.Name) Then
                                If bCopy Then
                                    If Not bQuiet Then
                                        Console.WriteLine("COPYING " & file.FullName & vbTab & " to " & DestDir & file.Name)
                                    End If
                                    FileCopy(file.FullName, DestDir & file.Name)
                                Else
                                    Console.WriteLine("FOUND " & file.FullName)
                                End If
                            End If
                        Next
                    Next
                Next
            Else
                Dim RegX As New Regex(RegXFileCriteria, RegexOptions.IgnoreCase)
                For Each srcDir As String In SourceDir
                    Dim dir As New DirectoryInfo(srcDir)
                    Dim files() As FileInfo = dir.GetFiles()
                    For Each file In files
                        If RegX.IsMatch(file.Name) Then
                            If bCopy Then
                                If Not bQuiet Then
                                    Console.WriteLine("COPYING " & file.FullName & vbTab & " to " & DestDir & file.Name)
                                End If
                                FileCopy(file.FullName, DestDir & file.Name)
                            Else
                                Console.WriteLine("FOUND " & file.FullName)
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            Console.Error.WriteLine("ERROR: " & ex.Message & vbNewLine & ex.StackTrace)
        End Try
    End Sub

    Private Sub ShowHelp()
        Console.WriteLine("GrabABunchOfFiles [<-options>] <criteria> <dest> <source> [<replace>]")
        Console.WriteLine("            replaces {r} in criteria with whatever you put in replace")
        Console.WriteLine("            can be as many arguments as you want/need                ")
        Console.WriteLine("            -h shows this help                                       ")
        Console.WriteLine("            -n means copy all files in every source directory        ")
        Console.WriteLine("            -s means read a list of line seperated source directories")
        Console.WriteLine("             just put the name of the line seperated file instead    ")
        Console.WriteLine("            -r means read a list of line seperated replacement values")
        Console.WriteLine("            -v means run verbose                                     ")
        Console.WriteLine("            -d means dont copy, just report files found (for debug)  ")
    End Sub

End Module
