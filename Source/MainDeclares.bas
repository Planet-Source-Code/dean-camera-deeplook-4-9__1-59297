Attribute VB_Name = "ModMainDeclares"
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

' -----------------------------------------------------------------------------------------------
        Public Const ShowItemKey As Boolean = False
       ' Change to TRUE for DEBUG purposes
' -----------------------------------------------------------------------------------------------

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SETREDRAW = &HB

Public PROJECTMODE As PMode
Public ShowCurrItemPic As Long
Public IsExit As Boolean
Public TotalLines As Long
Public ProjectPath As String
Public FilesRootDirectory As String

Public EXENewOrOld As String * 2

Public GlobalVars As Collection
Public GlobalVarsLoc As Collection

Public DefaultStatText As String

Public IncludeConsts As Boolean

Public Enum PMode
    VB6 = 0
    NET = 1
End Enum

#If False Then ' Learnt from Roger Gilchrist's Code Fixer - preserves the case of
    Private VB6, NET ' Enum statements when typing them into the IDE
#End If

' -----------------------------------------------------------------------------------------------

Public Function FixRelPath(ByVal Path As String) As String
    Dim Temp As String
    Dim C As Long
    Dim i As Long
    
    If Left$(Path, 1) = "." Then
        Do
            C = C + 1
            Path = Mid$(Path, 4)
        Loop Until Left$(Path, 1) <> "."
        
        Temp = Left$(ProjectPath, InStrRev(ProjectPath, "\", , vbTextCompare) - 1)
        For i = 1 To C
            Temp = Left$(Temp, InStrRev(Temp, "\", , vbTextCompare) - 1)
        Next
        
        FixRelPath = Temp & "\" & Path
    Else
        FixRelPath = Path
    End If
End Function

Public Function FileExists(ByVal Path As String) As Boolean 'Returns TRUE of the file exists. Can be used for paths or files.
    FileExists = Not (Dir(Path) = "")
End Function

Public Sub AddReportHeadder() ' Adds the start of the report text to the report Rich-Text control
    FrmResults.btnGenerateReport.Enabled = True

    AddReportText "=============================================================", True
    AddReportText "===================DEEPLOOK PROJECT REPORT==================="
    AddReportText "============================================================="
    AddReportText "==========    Made with DeepLook Version: " & App.Major & "." & App.Minor & "." & App.Revision & "    =========="
    AddReportText "============================================================="
    AddReportText "= You will need the ""Courier New"" font to view this report. ="
    AddReportText "============================================================="
    AddReportText "=         REPORT MADE ON: " & Now & Space(34 - Len(Now)) & "="
    AddReportText "============================================================="
End Sub

Public Sub AddReportFooter() ' Add the end of the Report to the report Rich-Text control
    AddReportText vbNewLine & vbNewLine & "============================================================="
    AddReportText "===============END OF DEEPLOOK PROJECT REPORT================"
    AddReportText "============================================================="
    AddReportText "===========    Made by Dean Camera, 2003-2004    ============"
    AddReportText "============================================================="
    AddReportText "===    Visit the En-Tech website at www.en-tech.i8.com    ==="
    AddReportText "============================================================="
End Sub

Private Sub AddReportText(ByVal AddText As String, Optional NoAddNL As Boolean) ' Adds a new line and the inputted text to the report
    If NoAddNL = False Then ' Add a new line at the start of the string
        FrmReport.rtbReportText.Text = FrmReport.rtbReportText.Text & vbNewLine & AddText
    Else ' Don't add new line
        FrmReport.rtbReportText.Text = FrmReport.rtbReportText.Text & AddText
    End If
End Sub

Public Function IsSysDLL(ByVal FileName As String) As String
    IsSysDLL = "DLL"
    If UCase(FileName) = "KERNEL32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "COMDLG32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "SHELL32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "USER32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "GDI32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "VERSION.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "OLEAUT32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "OLEPRO32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "ADVAPI32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "MSIMG32.DLL" Then
        IsSysDLL = "SysDLL"
    ElseIf UCase(FileName) = "STDOLE2.TLB" Then
        IsSysDLL = "SysDLL"
    End If
End Function

Public Sub KillProgram(ByVal Unloading As Boolean)
    If IsExit = True And Unloading = False Then
        Dim oFrm As Form
        IsExit = False

        For Each oFrm In VB.Forms
            Unload oFrm
        Next

        End
    End If
End Sub

Public Function GetRootDirectory(ByVal FileName As String) As String 'Retrieves only the path (not the filename) from a Path & Filename string.
    Dim SlashPos As Long

    SlashPos = InStrRev(FileName, "\") 'Get position of the last directory slash in the string

    If SlashPos <> 0 Then
        GetRootDirectory = Left$(FileName, SlashPos) 'Trim to only the directory, using the position of the last known directory slash
    End If

    If GetRootDirectory = "" Then Exit Function

    If Right$(GetRootDirectory, 1) <> "\" Then GetRootDirectory = GetRootDirectory & "\" 'If there is no directory slash at the end of the string,
    'add it to make all returned strings uniform.
End Function

Public Function ExtractFileName(ByVal FileNameAndPath As String) As String 'Gets only the filename from a file & path string
    Dim SlashPos As String

    SlashPos = InStrRev(FileNameAndPath, "\") 'Retrieve the position of the last path slash

    ExtractFileName = FileNameAndPath

    If SlashPos = 0 Then Exit Function

    ExtractFileName = Mid$(FileNameAndPath, SlashPos + 1) 'Trim path to only the filename
End Function

Public Function TrimJunk(ByVal Data As String) As String
    Data = Trim$(Data)
    If InStr(1, Data, "(") > 0 Then Data = Mid$(Data, 1, InStr(1, Data, "(") - 1)
    TrimJunk = Data
End Function

' -----------------------------------------------------------------------------------------------
