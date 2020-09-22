Attribute VB_Name = "ModFileRegisterBatCreator"
'  .======================================.
' /         DeepLook Project Scanner       \
' |       By Dean Camera, 2003 - 2005      |
' \   Completly re-written from scratch    /
'  '======================================'
' /  For more FREE software, please visit  \
' |         the En-Tech Website at:        |
' \            www.en-tech.i8.com          /
'  '======================================'
' / Most of this project is now commented  \
' \           to help developers.          /
'  '======================================'

Option Explicit

Sub CreateHeadder(FileName As String)
    Open FileName For Output As #55

    Print #55, "@echo off" & vbCrLf & "echo                    En-Tech DeepLook Project Scanner" & vbCrLf & "echo             *** Automatic file register batch script file ***" & vbCrLf & "echo -------------------------------------------------------------------------------" & vbCrLf & "echo You must be using WinME/98/95 and have the RegSvr32.exe in your windows folder." & vbNewLine & "echo." & vbCrLf & "pause" & vbCrLf & "Cls"
End Sub


Sub AddRegAndCopyFile(FileName As String, Findex As Long, Fmax As Long)
    Print #55, "echo *** Copying File #" & Findex & " of " & Fmax & " (" & FileName & ")..."
    Print #55, "echo." & vbCrLf & "copy """ & FileName & """, """ & "%WINDIR%\System\" & FileName & """"
    Print #55, "echo *** Registering File #" & Findex & " of " & Fmax & " (" & FileName & ")..."
    Print #55, "%WINDIR%\System\Regsvr32.exe ""%WINDIR%\System\" & FileName & """ /s"
    Print #55, "wait 1" & vbCrLf & "cls"
End Sub

Sub AddFooter(FileName As String)
    Print #55, "echo." & vbCrLf & "echo." & vbCrLf & "echo File copy/registration complete." & vbCrLf & "pause"
    Close #55
End Sub
