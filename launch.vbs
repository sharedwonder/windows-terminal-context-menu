If Wscript.Arguments.Count > 1 Then
    Set shell = WScript.CreateObject("Shell.Application")
    folder = WScript.Arguments(0)
    profile = WScript.Arguments(1)
    If Wscript.Arguments.Count = 2 Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & folder & """", "", "", 1
    ElseIf WScript.Arguments(2) = "-elevated" Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & folder & """", "", "RunAs", 1
    End If
End If
