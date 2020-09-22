Attribute VB_Name = "ModAnalyseDOTNET"
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

'-----------------------------------------------------------------------------------------------
'                                  .NET SCANNING SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   DeepLook is an advanced VB scanner. This module will scan some aspects of a .NET project
'   and display them on a treeview control, but it is still in its infancy and as such can only
'   give basic line stats and references. The .NET scanning engine does offer basic text reports.
'
'   The .NET scanning engine contained in this module is (C) Dean Camera.
'-----------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------
'                         NEW .NET SCANNING FEATURES IN THIS VERSION
'-----------------------------------------------------------------------------------------------
'   (None)
'-----------------------------------------------------------------------------------------------

Option Explicit

' -----------------------------------------------------------------------------------------------
Private vbFile As ClsNETvbFile
Private vbProject As ClsNETproject

Private ScanningVBfile As Boolean
Private CodeDone As Boolean
Private MakeReport As Boolean

Private ControlsInFile As Long
Private WaitForRegion As Boolean

Private FileType As String
Private ShowFNFerrors As Integer
' -----------------------------------------------------------------------------------------------

Private Sub PrepareForScan() ' Clears report, resises controls, etc. ready for scan
    ShowCurrItemPic = GetSetting("DeepLook", "Options", "ShowCurrItemPic", 1) ' /
    ShowFNFerrors = GetSetting("DeepLook", "Options", "ShowFNFErrors", 1)
    
    If ShowCurrItemPic = 1 Then                             '  \
        FrmSelProject.pgbGroupProgress.Width = 4215         '  |
        FrmSelProject.pgbAPB.Width = 4215                   '  |
        FrmSelProject.imgCurrScanObjType.Visible = True     '  |  Rearrange items on the SelProject form
    Else                                                    '  |  depending on the settings retrieved
        FrmSelProject.pgbGroupProgress.Width = 4695         '  |
        FrmSelProject.pgbAPB.Width = 4695                   '  |
        FrmSelProject.imgCurrScanObjType.Visible = False    '  |
    End If                                                  '  /

    FrmReport.rtbReportText.Text = "" ' Clear report
End Sub

Private Function FileExists(ByVal Path As String) As Boolean ' Returns TRUE of the file exists. Can be used for paths or files.
    FileExists = Not (Dir$(Path) = "")
End Function

Private Function ExtractFileName(ByVal FileNameAndPath As String) As String ' Gets only the filename from a file & path string
    Dim SlashPos As String

    SlashPos = InStrRev(FileNameAndPath, "\") ' Retrieve the position of the last path slash

    ExtractFileName = FileNameAndPath

    If SlashPos = 0 Then Exit Function

    ExtractFileName = Mid$(FileNameAndPath, SlashPos + 1) ' Trim path to only the filename
End Function

Private Function ExtractFilePath(ByVal FileNameAndPath As String) As String ' Gets only the path from a file & path string
    Dim SlashPos As String

    SlashPos = InStrRev(FileNameAndPath, "\") ' Retrieve the position of the last path slash

    ExtractFilePath = FileNameAndPath

    If SlashPos = 0 Then Exit Function

    ExtractFilePath = Left$(FileNameAndPath, SlashPos) ' Trim path
End Function

Sub AnalyseDotNetProject(ByVal Path As String) ' Main sub to scan a .NET project
    Dim Linedata As String, ReportPrjData As String, ReportPrjData2 As String
    Dim DonePrj As Long, i As Long

    CurrentScanFile "NETProject" ' Change current scan icon to a VB.NET project

    PrepareForScan ' Resize controls, clear report, etc.
    AddReportHeadder ' Add the headder info to the report

    PROJECTMODE = NET ' Set project mode to .NET

    If FileExists(Path) = False Then ' Make sure the file exists
        MsgBoxEx "Project file not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, , 5
        Exit Sub
    End If

    Open Path For Input As #1

    FrmSelProject.lblScanningName.Caption = ExtractFileName(Path)
    Set vbProject = New ClsNETproject ' Create a new instance to hold project stats

    Line Input #1, Linedata ' Check for valid .NET Project Headder
    If Trim$(Linedata) <> "<VisualStudioProject>" Then GoTo NotVBNetProj
    Line Input #1, Linedata ' Check for valid VB.NET Project Headder
    If Trim$(Linedata) <> "<VisualBasic" Then GoTo NotVBNetProj

    AddReportText vbNewLine & "============================================================="
    AddReportText "================Visual Basic .NET Project File==============="
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(Path)
    AddReportText "            Project Name: " & "?PLACE>ProjectName"
    AddReportText "         Project Version: " & "?PLACE>ProjectVersion"
    AddReportText "============================================================="
    AddReportText "?PLACE>ProjectStats"

    Do While Not EOF(1)
        FrmSelProject.pgbAPB.Value = (95 / LOF(1)) * DonePrj ' Increase progressbar
        Line Input #1, Linedata
        LookAtDOTNETProjLine Linedata ' Scan the current line
        DonePrj = DonePrj + Len(Linedata)
        DoEvents
    Loop

    Close #1

    With FrmResults.TreeView.Nodes ' Add the statistic nodes
        .Item(GetNodeNum("NETLines")).Text = "Lines (Inc. Blanks): " & vbProject.CodeLines
        .Item(GetNodeNum("NETLinesNB")).Text = "Lines (No Blanks): " & vbProject.CodeLinesNB
        .Item(GetNodeNum("NETLinesB")).Text = "Lines (Blanks): " & vbProject.BlankLines
        .Item(GetNodeNum("NETLinesComment")).Text = "Lines (Comment): " & vbProject.CommentLines
    End With

    TotalLines = vbProject.CodeLines + vbProject.CommentLines  ' Set the variable for layer calculation of lines/sec

    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectName", Mid$(FrmResults.TreeView.Nodes(GetNodeNum("NET_ProjectInfo_Name")).Text, 15)) ' Replace temp name in report with project name
    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectVersion", Mid$(FrmResults.TreeView.Nodes(GetNodeNum("NET_AssembleInfo_Version")).Text, 10)) ' Replace temp name in report with project version

    ReportPrjData = vbNewLine & "Lines (Inc. Blanks): " & vbProject.CodeLines & _
        vbNewLine & "Lines (No. Blanks): " & vbProject.CodeLinesNB & _
        vbNewLine & "Lines (Blanks): " & vbProject.BlankLines & _
        vbNewLine & "Lines (Comments): " & vbProject.CommentLines & _
        vbNewLine & vbNewLine & "           ------- References: -------"

    ReportPrjData2 = vbNewLine & "             ------- Imports: -------"

    With FrmResults.TreeView.Nodes
        For i = 1 To .Count
            If Mid$(.Item(i).Key, 1, 10) = "REFERENCE_" Then ReportPrjData = ReportPrjData & vbNewLine & .Item(i).Text
            If Mid$(.Item(i).Key, 1, 7) = "IMPORT_" Then ReportPrjData2 = ReportPrjData2 & vbNewLine & .Item(i).Text
        Next
    End With

    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectStats", ReportPrjData & vbNewLine & ReportPrjData2) ' Replace temp name in report with project stats

    FrmResults.TreeView.Nodes(1).Expanded = True
    Exit Sub

NotVBNetProj:
    MsgBoxEx "Not a VB.NET Project file!", vbCritical, "DeepLook - Scan error", , , , , PicError, , 5
End Sub

Private Sub AddNETprojectToTreeview() ' Adds the .NET project's nodes to the treeview
    With FrmResults.TreeView.Nodes
        .Add 1, tvwChild, "NET_ProjectInfo_FileName", "Project FileName: " & ExtractFileName(FrmSelProject.txtProjectPath.Text), "Info"
        .Add 1, tvwChild, "NET_ProjectInfo_Path", "Project Path: " & ExtractFilePath(FrmSelProject.txtProjectPath.Text), "Info"
        .Add 1, tvwChild, "NET_ProjectInfo_Name", "Project Name: " & .Item(1).Text, "Info"

        .Add 1, tvwChild, "NET_AssembleInfo_Version", "Version: Unknown", "Info"

        .Add 1, tvwChild, "NETLines", "?", "Info"
        .Add 1, tvwChild, "NETLinesNB", "?", "Info"
        .Add 1, tvwChild, "NETLinesB", "?", "Info"
        .Add 1, tvwChild, "NETLinesComment", "?", "Info"

        .Add 1, tvwChild, "NETreferences", "References", "REFCOM"
        .Add 1, tvwChild, "NETimports", "Imports", "SysDLL"

        .Add 1, tvwChild, "NETfiles", "Files", "NETvb"
    End With
End Sub

Private Function GetNodeNum(ByVal NodeKey As String, Optional IndexNum As Long) As Long ' Find the node index of a item on the treeview,
    On Error Resume Next                                                                    ' from the given key.
    GetNodeNum = FrmResults.TreeView.Nodes.Item(NodeKey).Index
End Function

Private Sub LookAtDOTNETProjLine(ByVal Linedata As String) ' Scans each project line
    Static LookingAtSettings As Boolean, SecondaryLineData As String
    Linedata = Trim$(Linedata)

    CurrentScanFile "NETProject" ' Change current scan item icon to a VB.NET file

    If Linedata = "<Settings" Then LookingAtSettings = True: Exit Sub                                      ' \
    If LookingAtSettings = True And Linedata = ">" Then LookingAtSettings = False: Exit Sub                ' | Retrieves settings from the
    If LookingAtSettings = True And Mid$(Linedata, 1, 15) = "AssemblyName = " Then                          ' | .NET project file
        FrmResults.TreeView.Nodes.Add , , "NETproj", Mid$(Linedata, 17, Len(Linedata) - 17), "NETproject"   ' /
        AddNETprojectToTreeview
        Exit Sub
    End If

    If LookingAtSettings = True Then Exit Sub

    If Mid$(Linedata, 1, 20) = "<Import Namespace = " Then AddImport Mid$(Linedata, 22, Len(Linedata) - 25): Exit Sub ' Get Imports from .NET project file
    If Mid$(Linedata, 1, 10) = "<Reference" Then
        Line Input #1, Linedata
        Linedata = Trim$(Linedata)
        AddReference Mid$(Linedata, 9, Len(Linedata) - 9) ' Retrieve references from the .NET project file
        Exit Sub
    End If

    If Linedata = "<File" Then ' Get the filename of each .NET file and scan it
        Line Input #1, Linedata
        Line Input #1, SecondaryLineData

        Do
            If InStr(1, SecondaryLineData, "BuildAction = ") = 0 Then
                Line Input #1, SecondaryLineData
            Else
                Exit Do
            End If
        Loop

        If Trim$(SecondaryLineData) <> "BuildAction = ""Compile""" Then Exit Sub ' Make sure it's not a .TXT, .HTM, etc. file

        Linedata = Trim$(Linedata)
        Linedata = Mid$(Linedata, 12, Len(Linedata) - 12)
        AnalyseVBfile Linedata ' Scan the file
    End If
End Sub

Private Sub AnalyseVBfile(ByVal FileName As String) ' Look at a .NET file
    Dim Linedata As String, i As Long, FileTitle As String

    CurrentScanFile "File" ' Change current scan item icon to a .NET file

    ScanningVBfile = False   '  \
    CodeDone = False         '  | Reset variables
    WaitForRegion = False    '  /

    ControlsInFile = 0

    If FileExists(FileName) = False Then ' Make sure the file exists
        If ShowFNFerrors Then MsgBoxEx "File not found: " & FileName, vbCritical, "DeepLook - Scan Error", , , , , PicError, , 5
        Exit Sub
    End If

    If InStr(1, FileName, "ASSEMBLYINFO.VB", vbTextCompare) <> 0 Then ScanAssemblyFile FileName: Exit Sub ' If the Assembly info file, look at it seperatly

    On Error Resume Next

    With FrmResults.TreeView.Nodes ' Add statistic nodes to treeview
        .Add GetNodeNum("NETfiles"), tvwChild, "TEMPvb", "?", "NETvb"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_Lines", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_LinesNB", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_LinesB", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_LinesComment", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_Controls", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_Imports", "Imports", "SysDLL"
    End With

    Set vbFile = New ClsNETvbFile ' Create new instance

    Open FileName For Input As #2

    Do
        Line Input #2, Linedata
        ScanVBfileLine Linedata ' Look at each line
        DoEvents
    Loop While Not EOF(2)

    Close #2

    vbFile.Controls = ControlsInFile
    FileTitle = FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Text

    With FrmResults.TreeView.Nodes ' Fix the temp names and data
        .Item(GetNodeNum("TEMPvb_Imports")).Key = FileTitle & "_Imports"

        .Item(GetNodeNum("TEMPvb_Lines")).Text = "Lines (Inc. Blanks): " & (vbFile.CodeLines - 1)
        .Item(GetNodeNum("TEMPvb_Lines")).Key = FileTitle & "_Lines"
        .Item(GetNodeNum("TEMPvb_Controls")).Text = "Controls: " & vbFile.Controls
        .Item(GetNodeNum("TEMPvb_Controls")).Key = FileTitle & "_Controls"
        .Item(GetNodeNum("TEMPvb_LinesNB")).Text = "Lines (No Blanks): " & (vbFile.CodeLinesNB - 1)
        .Item(GetNodeNum("TEMPvb_LinesNB")).Key = FileTitle & "_LinesNB"
        .Item(GetNodeNum("TEMPvb_LinesB")).Text = "Lines (Blanks): " & vbFile.BlankLines
        .Item(GetNodeNum("TEMPvb_LinesB")).Key = FileTitle & "_LinesB"
        .Item(GetNodeNum("TEMPvb_LinesComment")).Text = "Lines (Comments): " & vbFile.CommentLines
        .Item(GetNodeNum("TEMPvb_LinesComment")).Key = FileTitle & "_LinesComment"
    End With

    For i = 1 To FrmResults.TreeView.Nodes.Count ' Fix all remaining temp key names
        If InStr(1, FrmResults.TreeView.Nodes(i).Key, "TEMPvb") <> 0 Then
            FrmResults.TreeView.Nodes(i).Key = Replace$(FrmResults.TreeView.Nodes(i).Key, "TEMPvb", FileTitle)
        End If
    Next i

    AddReportText vbNewLine & "-------------------------------------------------------------"
    AddReportText "                   VISUAL BASIC .NET FILE"
    AddReportText "-------------------------------------------------------------"
    AddReportText "               File Name: " & ExtractFileName(FileName)
    AddReportText "               File Type: " & FileType
    AddReportText "                    Name: " & FileTitle
    AddReportText vbNewLine & "     Lines (Inc. Blanks): " & vbFile.CodeLines
    AddReportText "       Lines (No Blanks): " & vbFile.CodeLinesNB
    AddReportText "          Lines (Blanks): " & vbFile.BlankLines
    AddReportText "        Lines (Comments): " & vbFile.CommentLines

    If FileType = "Form" Then AddReportText vbNewLine & "                Controls: " & vbFile.Controls

    vbProject.BlankLines = vbFile.BlankLines     '  \
    vbProject.CodeLines = vbFile.CodeLines       '  | Add to the project's
    vbProject.CodeLinesNB = vbFile.CodeLinesNB   '  | total statistics
    vbProject.CommentLines = vbFile.CommentLines '  /

    FrmResults.TreeView.Nodes(GetNodeNum("NETfiles")).Expanded = True
End Sub

Private Sub AddImport(ByVal Linedata As String) ' Add Imports to the treeview
    Dim NodeNum As Long

    NodeNum = GetNodeNum("NETimports")
    FrmResults.TreeView.Nodes.Add NodeNum, tvwChild, "IMPORT_" & Linedata, Linedata, "SysDLL"
End Sub

Private Sub AddReference(ByVal Linedata As String) ' Add references to the treeview
    Dim NodeNum As Long

    If Linedata = "c" Then Exit Sub

    NodeNum = GetNodeNum("NETreferences")
    FrmResults.TreeView.Nodes.Add NodeNum, tvwChild, "REFERENCE_" & Linedata, Linedata, "DLL"
End Sub


Private Sub ScanVBfileLine(ByVal Linedata As String) ' Scan an individual line of a VB file
    vbProject.TotalLines = 1

RemoveIndent:                                                                                           ' Loops here to remove any
    If Mid$(Linedata, 1, 1) = vbTab Then Linedata = Mid$(Linedata, 5): GoTo RemoveIndent  ' indents (tabs) from the code

    If Mid$(Linedata, 1, 9) = "End Class" Then CodeDone = True: Exit Sub

    If Mid$(Linedata, 1, 48) = "#Region "" Windows Form Designer generated code """ Then WaitForRegion = True
    If Mid$(Linedata, 1, 11) = "#End Region" Then WaitForRegion = False

    If Mid$(Linedata, 1, 8) = "Imports " Then ' Add import data to treeview
        FrmResults.TreeView.Nodes.Add GetNodeNum("TEMPvb_Imports"), tvwChild, "TEMPvb_Import_" & Mid$(Linedata, 9), Mid$(Linedata, 9), "SysDLL"
        Exit Sub
    End If

    If ScanningVBfile = False Then GetNameAndType (Linedata): Exit Sub ' Check what the file type is before scanning

    If CodeDone = True Then Exit Sub
    If WaitForRegion = True Then Exit Sub

    If IsHybridLine(Linedata) = True Then ' Hybrid line (Both code and comment)
        vbFile.CommentLines = 1
        vbFile.CodeLines = 1
        vbFile.CodeLinesNB = 1
        Exit Sub
    End If

    If UCase(Mid$(Linedata, 1, 4)) = "REM " Then vbFile.CommentLines = 1: Exit Sub
    If Mid$(Linedata, 1, 1) = "'" Then vbFile.CommentLines = 1: Exit Sub

    vbFile.CodeLines = 1
    If Linedata = "" Then vbFile.BlankLines = 1: Exit Sub
    vbFile.CodeLinesNB = 1
End Sub

Private Sub GetNameAndType(ByVal Linedata As String) ' Retrieves the type of VB File (Form, etc.) and sets appropriate variables
    If Mid$(Linedata, 1, 7) = "Module " Then ' >>MODULE<<
        FileType = "Module"

        FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Text = Mid$(Linedata, 8)
        FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Image = "Module"
        ScanningVBfile = True
    End If


    If Mid$(Linedata, 1, 13) = "Friend Class " Or Mid$(Linedata, 1, 13) = "Public Class " Then
        FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Text = Mid$(Linedata, 14)
        Line Input #2, Linedata

        If InStr(1, Linedata, "Inherits System.Windows.Forms.Form") <> 0 Then ' >>FORM<<
            FileType = "Form"

            FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Image = "Form"
        Else ' >>CLASS MODULE<<
            FileType = "Class Module"

            FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Image = "Class"
        End If
        ScanningVBfile = True
    End If
End Sub

Private Sub CurrentScanFile(ByVal FileType As String) ' Changes the little icon on the frmSelProject (if turned on) to
    If ShowCurrItemPic = False Then Exit Sub        ' indicate what type of file is being scanned

    With FrmSelProject.imgCurrScanObjType
        Select Case FileType
            Case "NETProject"
                .Picture = FrmResults.ilstImages.ListImages("NETproject").Picture
            Case "File"
                .Picture = FrmResults.ilstImages.ListImages("NETvb").Picture
            Case "Clean"
                .Picture = FrmResults.ilstImages.ListImages(17).Picture
        End Select
    End With
End Sub

Private Sub ScanAssemblyFile(ByVal FileName As String) ' Scans an Assembly.vb File to get project version
    Dim Linedata As String, LinePos As Long

    Open FileName For Input As #2

    Do
        Line Input #2, Linedata
        If InStr(1, Linedata, "AssemblyVersion(") <> 0 Then Exit Do
        DoEvents
    Loop While Not EOF(2)

    If Linedata = "" Then Exit Sub
    LinePos = InStr(1, Linedata, "AssemblyVersion(") ' Find version line
    Linedata = Mid$(Linedata, LinePos + 17, Len(Linedata) - LinePos - 19)
    If Right$(Linedata, 1) = Chr$(34) Then Linedata = Mid$(Linedata, 1, Len(Linedata) - 1)

    FrmResults.TreeView.Nodes.Item(GetNodeNum("NET_AssembleInfo_Version")).Text = "Version: " & Linedata

    Close #2
End Sub

Private Function IsHybridLine(ByVal Linedata As String) As Boolean ' Returns TRUE if the entered string contains both code and comment.
    Dim i As Long, InString As Boolean

    If InStr(1, Linedata, "'") = 0 Then Exit Function ' Comments after code can only use the "'" symbol, not the REM statement. Check to see
    ' if the line contains the "'" character, otherwise skip the sub to save time.
    For i = 1 To Len(Linedata)
        If Mid$(Linedata, i, 1) = """" Then InString = Not InString ' Changes the InString boolean variable as DeepLook finds a " symbol. This prevents
        If Mid$(Linedata, i, 1) = "'" And InString = False Then ' the program mis-interpreting ' symbols in strings as comments.
            IsHybridLine = True
            Exit Function
        End If
    Next
End Function

Private Sub AddReportText(ByVal AddText As String, Optional NoAddNL As Boolean) ' Adds a new line and the inputted text to the report
    If MakeReport = False Then Exit Sub

    If NoAddNL = False Then ' Add a new line at the start of the string
        FrmReport.rtbReportText.Text = FrmReport.rtbReportText.Text & vbNewLine & AddText
    Else ' Don't add new line
        FrmReport.rtbReportText.Text = FrmReport.rtbReportText.Text & AddText
    End If
End Sub
