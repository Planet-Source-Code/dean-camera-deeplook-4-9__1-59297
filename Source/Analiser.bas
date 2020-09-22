Attribute VB_Name = "ModAnalyseVB6"
'  .======================================.
'/         DeepLook Project Scanner       \
'|       By Dean Camera, 2003 - 2005      |
'\   Completely re-written from scratch    /
'  '======================================'
'/  For more FREE software, please visit  \
'|         the En-Tech Website at:        |
'\            www.en-tech.i8.com          /
'  '======================================'
'/ Most of this project is now commented  \
'\           to help developers.          /
'  '======================================'

'-----------------------------------------------------------------------------------------------
'                                    VB6 SCANNER SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   DeepLook is an advanced VB6 scanner. It is capable of showing almost all aspects of a source
'   project. The VB6 scanning engine in this module is (C) Dean Camera, and can show statistics
'   in a treeview and a text report. This module is capable of scanning projects made in all
'   versions of VB (except .NET) but is optimised and written around the VB6 version, and so
'   some errors may occur in projects saved in versions below VB5.
'
'   This VB scanner differs from many other scanners, because it excludes all header data from
'   the code statistics. Every VB file contains many hidden statements for the VB program, such
'   as Control information on forms, and name info. Most other scanners include these lines when
'   they are not part of the actual code.
'
'   The VB scanning engine contained in this module is (C) Dean Camera.
'-----------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------
'                         NEW VB6 SCANNING FEATURES IN THIS VERSION
'-----------------------------------------------------------------------------------------------
'   * Unused variable scanner fixed
'   * A registering BAT file is now also created (alternative to FileRegister.exe)
'   * Speed enhancements by tweaking the code (e.g. frequent use of ElseIf to speed up scanning)
'   * Program now correctly interprets multiple variable declares on a single line
'   * ByVal used on Parameters to increase scan time
'   * Colour coded constants and global/local unused variables
'   * Warnings for old EXE/Project Files
'   * Added speed increase for FixPrjTreeViewItems sub
'   * Added naming conventions on all forms and controls
'
'   TODO: Alter UV scanner to recognise unused variables that use un-unique names
'-----------------------------------------------------------------------------------------------
'                       NEW VB6 SCANNING FEATURES IN PREVIOUS VERSIONS
'-----------------------------------------------------------------------------------------------
'   * Fixed many minor bugs
'   * Added support for designer (.dsr) files
'   * Added support for line numbers
'   * Made heaps of speed tweaks, DeepLook can do more operations in half the time
'   * Added new "statements" section which counts such lines as FOR, IF and DO
'   * Now counts Const, Type and Enum statements
'   * Shifted Potentially Malicious Code section to inside each individual file
'   * Shifted repeated code to one sub for easy management and tweaking
'   * Fixed major Sub/Func/Prop line counting bug
'   * Added a "Project Totals" node for total project lines, etc.
'   * Added Declared DLL File Info
'   * Added support for events
'   * Added total lines in the group node
'   * Empty routines are now coloured red
'   * Now colours different types of SPFs in different colours
'   * Ultra-increased node cleanup routine per project in group projects
'   * Reports are always generated due to a user-request
'   * External subs are now always coloured grey due to new colouring scheme
'   * Dramatically reduced treeview clearing time
'   * Added new treeview right-click menu to highlight special nodes
'   * Added a new "Copy Report"
'   * Fixed group non-relative file not found error
'   * Fixed inability to scan blank files
'   * Removed option to use XP style progressbar
'   * Added a better progress bar, removed AdvancedProgressBar control
'   * Better PMC Scanner Engine
'   * Unused Variable Checker
'   * Mode tab on the results form
'   * Scanning phases shown on the selproject form
'-----------------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------------------------------------------------
Private Project As ClsProjectFile
Private RefComKeyNames As Collection
Private ProjectItem As ClsPrjItem
Private UsesGroup As Boolean
Private ShowSPFParams As Boolean
Private ShowIndividualLines As Boolean
Private CheckForMalicious As Long
Private RelatedDocumentFNames As String
Private TotalSubs As Long, TotalFunctions As Long, TotalProperties As Long, TotalEvents As Long, TotalDecSubs As Long, TotalDecFunctions As Long
Private InSub As Boolean, InFunction As Boolean, InProperty As Boolean
Private CurrSPFLines As Long, CurrSPFLinesNB As Long, CurrSPFName As String, CurrSPFColour As Long
Private FileInfo As clsFileProp
Private TotalRelDocs As Long
Private GroupTotalLines As Long, GroupTotalLinesNB As Long, GroupTotalPrj As Long
Private ThisPrgNodeStart As Long
Private NodeNum As Long

Private Group1Percent As Single, Project1Percent As Single, ShowFNFerrors As Integer
'-----------------------------------------------------------------------------------------------

Public Sub ClearCollections() 'Reset the variables and clear arrays ready for
    Dim i As Long 'another project to be scanned

    ModVariableHandler.ClearGlobals

    FilesRootDirectory = ""
    UsesGroup = False

    If RefComKeyNames Is Nothing Then Set RefComKeyNames = New Collection
    If FileInfo Is Nothing Then Set FileInfo = New clsFileProp

    For i = 0 To RefComKeyNames.Count - 1
        RefComKeyNames.Remove i 'Clears all items in array
    Next

    FrmResults.lstBadCode.Clear 'Potentially Malicious Code is stored in a listbox for sorting,
    'which is prevented from re-drawing via an API call to increase
    'speed. It must be cleared before a new project is scanned.

    ShowSPFParams = GetSetting("DeepLook", "Options", "ShowSPFParams", 1)     '\
    CheckForMalicious = GetSetting("DeepLook", "Options", "PMCCheck", 1)      '| Get program's settings
    ShowCurrItemPic = GetSetting("DeepLook", "Options", "ShowCurrItemPic", 1) '|
    ShowFNFerrors = GetSetting("DeepLook", "Options", "ShowFNFErrors", 1)     '|
    If GetSetting("DeepLook", "Options", "ScanConstAsVar", 1) <> 1 Then       '/
        IncludeConsts = True
    Else
        IncludeConsts = False
    End If

    If GetSetting("DeepLook", "Options", "ShowSPFLines", 0) = 0 Then
        ShowIndividualLines = False
    Else
        ShowIndividualLines = True
    End If

    If ShowCurrItemPic = 1 Then                          '  \
        FrmSelProject.pgbGroupProgress.Width = 4215      '  |
        FrmSelProject.pgbAPB.Width = 4215                '  |
        FrmSelProject.imgCurrScanObjType.Visible = True  '  |  Rearrange items on the SelProject form
    Else                                                 '  |  depending on the settings retrieved
        FrmSelProject.pgbGroupProgress.Width = 4695         '  |
        FrmSelProject.pgbAPB.Width = 4695                '  |
        FrmSelProject.imgCurrScanObjType.Visible = False '  |
    End If                                               '  /

    FrmReport.rtbReportText.Text = "" 'Clear report

    RelatedDocumentFNames = ""
End Sub

'-----------------------------------------------------------------------------------------------

Public Sub AnalyseGroup(GroupFileName As String) 'Master command to analyse a Group (.vbg) file.
    Dim Linedata As String, DoneGrp As Long

    CurrentScanFile "Group" 'Change small picture of current operation to a VB Group symbol

    DoEvents

    FilesRootDirectory = GetRootDirectory(GroupFileName) 'retrieve directory of the group file

    FrmSelProject.lblScanningName.Caption = ExtractFileName(GroupFileName) 'change the label showing the current file name to the group's file name

    GroupFileName = FixRelPath(GroupFileName)

    If FileExists(GroupFileName) = False Then 'Throw a graceful error if the file is not found
        MsgBoxEx "Group file not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, , 5
        Exit Sub
    End If

    FrmResults.TreeView.Nodes.Add , , "GROUP", "Project Group", "Group" 'Add the group to the treeview
    FrmResults.TreeView.Nodes.Add 1, tvwChild, "GROUP_Lines", "?", "Info"
    FrmResults.TreeView.Nodes.Add 1, tvwChild, "GROUP_LinesNB", "?", "Info"
    FrmResults.TreeView.Nodes.Add 1, tvwChild, "GROUP_TotalPrj", "?", "Info"

    GroupTotalPrj = 0
    UsesGroup = True 'Tell the AnalyseProject sub that it should be adding project files to the group node

    GroupTotalLines = 0
    GroupTotalLinesNB = 0

    Open GroupFileName For Input As #1

    AddReportText vbNewLine & "*************************************************************"
    AddReportText "**********************Visual Basic Group*********************"
    AddReportText "*************************************************************"
    AddReportText "                   File Name:" & ExtractFileName(GroupFileName)
    AddReportText "*************************************************************"

    Line Input #1, Linedata      'Get the line data and increment the completed percentage
    DoneGrp = Len(Linedata) + 2

    If InStr(1, Linedata, "VBGROUP", vbBinaryCompare) = 0 Then GoTo InvalidGroupFile 'Group file must contain valid VB6 Group Header

    Group1Percent = (100 / LOF(1))

    Do While Not EOF(1)
        FrmSelProject.pgbGroupProgress.Value = Group1Percent * DoneGrp     'Increment progress bar
        Line Input #1, Linedata 'Retrieve line data
        DoneGrp = DoneGrp + Len(Linedata)
        DoEvents 'Prevent the program from locking up while processing and stopping the progressbar from refreshing

        If Left$(Linedata, 15) = "StartupProject=" Then
            AnalyseVBProject Mid$(Linedata, 16), True
            GroupTotalLines = GroupTotalLines + Project.ProjectLines: GroupTotalLinesNB = GroupTotalLinesNB + Project.ProjectLinesNB  'Pass a TRUE value to the "StartupProjectInGroup" parameter
        End If

        If Left$(Linedata, 8) = "Project=" Then
            AnalyseVBProject Mid$(Linedata, 9)
            GroupTotalLines = GroupTotalLines + Project.ProjectLines
            GroupTotalLinesNB = GroupTotalLinesNB + Project.ProjectLinesNB
        End If
    Loop

    Close #1

    FrmResults.btnFileCopy.Enabled = False

    With FrmResults.TreeView.Nodes
        .Item(2).Text = "Total Lines (Inc. Blanks): " & Format$(GroupTotalLines, "###,###,###")
        .Item(3).Text = "Total Lines (No Blanks): " & Format$(GroupTotalLinesNB, "###,###,###")
        .Item(4).Text = "Total Projects in Group: " & Format$(GroupTotalPrj, "###,###,###")

        .Item(1).Expanded = True 'Expand the group node on the treeview
    End With
    Exit Sub

InvalidGroupFile:
    Close #1
    MsgBoxEx "Invalid Group File - File is not a VB6 Group File!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|", 5
End Sub

Public Sub AnalyseVBProject(ProjectFileName As String, Optional StartupProjectInGroup As Boolean) 'Master command to analyse a VB Project
    Dim Linedata As String, DonePrj As Long

    FrmSelProject.pgbAPB.Value = 0
    FrmSelProject.lblScanPhase.Caption = "Code Scan Phase"
    FrmSelProject.lblScanPhase.ForeColor = &HC000&
    FrmSelProject.pgbAPB.Color = &HC000&

    ModVariableHandler.ClearGlobals
    ModVariableHandler.ClearLocals

    PROJECTMODE = VB6
    GroupTotalPrj = GroupTotalPrj + 1 'Increment total projects scanned (for group scan)

    CurrentScanFile "Project" 'Change small picture of current operation to a VB Project symbol

    FrmSelProject.lblScanningName.Caption = ExtractFileName(ProjectFileName) 'Change the label on the SelProject form to the project's file name

    DoEvents 'Prevent VB from locking up when scanning projects - and so the progressbar can refresh

    ProjectFileName = FixRelPath(ProjectFileName)

    If FileExists(FilesRootDirectory & ProjectFileName) = False And FileExists(ProjectFileName) = False Then
        If UsesGroup = True Then
            If ShowFNFerrors Then MsgBoxEx "Project file """ & ProjectFileName & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|", 5 'Throw an error if the file cannot be found (only shown if show FNF errors setting turned on)
            
            FrmResults.TreeView.Nodes.Add 1, tvwChild, "PROJECT_" & ProjectFileName, ProjectFileName, "Unknown"
        Else
            MsgBoxEx "Project file """ & ProjectFileName & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|", 5 'Throw an error if the file cannot be found (always shown if project not part of a group)
            
            FrmResults.TreeView.Nodes.Add , , "PROJECT_?", "?", "Project" 'File found, so add to treeview
        End If
        
        Exit Sub
    End If

    Set Project = New ClsProjectFile 'Initialise a new instance of ClsProjectFile
    'to store project statistics/data

    ThisPrgNodeStart = FrmResults.TreeView.Nodes.Count 'Recode the node start of the project to increase cleanup time

    Project.ProjectPath = ProjectFileName

    AddProjectToTreeView
    If StartupProjectInGroup = True Then FrmResults.TreeView.Nodes(GetNodeNum("PROJECT_?")).Bold = True 'Make the project bold in the treeview if it is a startup project

    AddReportText vbNewLine & "============================================================="
    AddReportText "===================Visual Basic Project File================="
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(ProjectFileName)
    AddReportText "            Project Name: " & "?PLACE>ProjectName"
    AddReportText "         Project Version: " & "?PLACE>ProjectVersion"
    AddReportText "============================================================="
    AddReportText "?PLACE>ProjectStats"

    FrmResults.btnFileCopy.Enabled = True

    InSub = False
    InFunction = False
    InProperty = False

    Open ProjectFileName For Input As #2

    Project1Percent = (100 / LOF(2))

    Do While Not EOF(2)
        FrmSelProject.pgbAPB.Value = Project1Percent * DonePrj 'Increase progressbar
        Line Input #2, Linedata 'Get the line data from the project file
        AnalyseProjectLine Linedata 'Look at the project line to see what it contains
        DonePrj = DonePrj + Len(Linedata)
        DoEvents 'Prevent VB from locking up when processing - and to let the progressbar refresh
    Loop

    Close #2

    GetProjectEXEStats

    If EXENewOrOld = "1N" Then
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_EXEFileVerDetails", "The Project File is a newer version than your compiled EXE. Please Recompile.", "Warning"
    ElseIf EXENewOrOld = "2N" Then
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_EXEFileVerDetails", "The compiled EXE is a newer version than your Project File.", "Warning"
    End If

    FixPrjTreeViewItems 'Fix up the treeview items, captions and keys. All items use a "?" in the key instead of the project name
    'until the FixPrjTreeViewItems sub changes this to prevent group projects from interfering with each other.

    AddProjectReportText 'Add collected stats about the current project to the report

    ModVariableHandler.AnalyseVBProjectForVars ProjectFileName

    If GetNodeNum("GROUP") = 0 Then FrmResults.TreeView.Nodes(1).Expanded = True 'If not part of a group file, expand the project
End Sub

Public Sub AnalyseSingleVBItem(ByVal FileName As String) 'Master sub to analyse a single file (.fmr, .bas, etc.) rather than a project
    Dim ItemKeyNum As Long, QuestionPos As Long, i As Long

    PROJECTMODE = VB6

    FrmSelProject.pgbAPB.Value = 0
    FrmSelProject.lblScanPhase.Caption = "Code Scan Phase"
    FrmSelProject.lblScanPhase.ForeColor = &HFF00&
    FrmSelProject.pgbAPB.Color = &HFF00&

    DoEvents

    FileName = FixRelPath(FileName)

    If FileExists(FileName) = False Then 'If file doesn't exist, throw an error
        MsgBoxEx "File not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError
        Exit Sub
    End If

    Set Project = New ClsProjectFile 'Create a new project class to stop the analysis subs from giving errors
    'when trying to add to it's values - this is not actually used in the treeview

    Set ProjectItem = New ClsPrjItem
    Project.ProjectPath = ""

    FrmResults.btnFileCopy.Enabled = False

    With FrmResults.TreeView.Nodes

        FrmResults.TreeView.Nodes.Add , , "PROJECT_?", "?", "Project"


        Select Case Right$(UCase(FileName), 4) 'Add the required nodes depending on the file type
            Case ".FRM"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_FORMS", "Forms", "Form"
                .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_VARIABLES", "0"
                GetFormStats FileName
                .Remove GetNodeNum("PROJECT_?_FORMS_LINES")
                .Remove GetNodeNum("PROJECT_?_FORMS_VARIABLES")
            Case ".BAS"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_MODULES", "MODULES", "Module"
                .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_LINES", "0"
                .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_VARIABLES", "0"
                GetModuleStats FileName & ";" & ExtractFileName(FileName)
                .Remove GetNodeNum("PROJECT_?_MODULES_LINES")
                .Remove GetNodeNum("PROJECT_?_MODULES_VARIABLES")
            Case ".CLS"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_CLASSES", "Classes", "Class"
                .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_LINES", "0"
                .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_VARIABLES", "0"
                GetClassStats FileName & ";" & ExtractFileName(FileName)
                .Remove GetNodeNum("PROJECT_?_CLASSES_LINES")
                .Remove GetNodeNum("PROJECT_?_CLASSES_VARIABLES")
            Case ".CTL"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_USERCONTROLS", "User Controls", "UserControl"
                .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_VARIABLES", "0"
                GetUserControlStats FileName
                .Remove GetNodeNum("PROJECT_?_USERCONTROLS_LINES")
                .Remove GetNodeNum("PROJECT_?_USERCONTROLS_VARIABLES")
            Case ".PAG"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_PROPERTYPAGES", "Property Pages", "PropertyPage"
                .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_LINES", "0"
                .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_VARIABLES", "0"
                GetPropertyPageStats FileName
                .Remove GetNodeNum("PROJECT_?_PROPERTYPAGES_LINES")
                .Remove GetNodeNum("PROJECT_?_PROPERTYPAGES_VARIABLES")
            Case ".DOB"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_USERDOCUMENTS", "User Documents", "UserDocument"
                .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_VARIABLES", "0"
                GetUserDocumentStats FileName
                .Remove GetNodeNum("PROJECT_?_USERDOCUMENTS_LINES")
                .Remove GetNodeNum("PROJECT_?_USERDOCUMENTS_VARIABLES")
            Case ".DSR"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_DESIGNERS", "Designers", "Designer"
                .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_VARIABLES", "0"
                GetDesignerStats FileName
                .Remove GetNodeNum("PROJECT_?_DESIGNERS_LINES")
                .Remove GetNodeNum("PROJECT_?_DESIGNERS_VARIABLES")
        End Select

        For i = 1 To .Count 'Expand all nodes
            .Item(i).Expanded = True
        Next

        .Item(GetNodeNum("PROJECT_?")).Text = "(Temp Project)" 'Rename the main node as a "temporary" project

        For ItemKeyNum = 1 To .Count 'Go over every node, checking for Related Documents, and replacing the temp project name "?" with the correct
            Do
                QuestionPos = InStr(1, .Item(ItemKeyNum).Text, "  ") 'Checks for indents, which make line continuations look wrong
                If QuestionPos <> 0 Then '"QuestionPos" variable used here only to save on variable requirements
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, QuestionPos - 1) & Mid$(.Item(ItemKeyNum).Text, QuestionPos + 2)
                Else 'No indents left - exit the loop
                    Exit Do
                End If
            Loop

            If InStr(1, .Item(ItemKeyNum).Text, "[EXT]") <> 0 Then 'Is an external (DLL) call (declared sub or function)
                .Item(ItemKeyNum).Parent.ForeColor = RGB(150, 150, 150) 'Make key text a light-grey colour
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 10) 'Remove the now unnecessary "[EXT]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[1]" Then 'SPF Colour 1
                .Item(ItemKeyNum).Parent.ForeColor = RGB(130, 0, 200) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[2]" Then 'SPF Colour 2
                .Item(ItemKeyNum).Parent.ForeColor = RGB(200, 0, 150) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[3]" Then 'SPF Colour 3
                .Item(ItemKeyNum).Parent.ForeColor = RGB(10, 150, 10) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[6]" Then 'SPF Colour 4
                .Item(ItemKeyNum).Parent.ForeColor = RGB(249, 164, 0) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[7]" Then 'SPF Colour 5
                .Item(ItemKeyNum).Parent.ForeColor = RGB(217, 206, 19) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[8]" Then 'SPF Colour 6
                .Item(ItemKeyNum).Parent.ForeColor = RGB(19, 217, 192) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[4]" Then 'SPF Colour 7
                .Item(ItemKeyNum).Parent.ForeColor = RGB(20, 90, 100) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[5]" Then 'SPF Colour 8
                .Item(ItemKeyNum).Parent.ForeColor = RGB(50, 23, 80) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            End If

            If InStr(1, .Item(ItemKeyNum).Text, "No Blanks):") <> 0 Then
                If Right$(.Item(ItemKeyNum).Text, 2) = "0" Then
                    If InStr(1, .Item(ItemKeyNum).Parent.Key, "_SUB") <> 0 Then
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                    ElseIf InStr(1, .Item(ItemKeyNum).Parent.Key, "_FUNC") <> 0 Then
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                    ElseIf InStr(1, .Item(ItemKeyNum).Parent.Key, "_PROP") <> 0 Then
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                    End If
                End If
            End If
        Next
    End With

    ModVariableHandler.ScanFile FileName
End Sub

Private Sub AnalyseProjectLine(ByVal Linedata As String) 'Sub to look at the contents of the project's line and analyse accordingly
    CurrentScanFile "Project" 'Change the small picture to the project symbol

    'ANALISE, CUT AND PARSE LINE TO APPROPRIATE SUB/FUNCTION
    If Left$(Linedata, 5) = "Name=" Then
        Project.ProjectName = Mid$(Linedata, 7, Len(Linedata) - 7)
    ElseIf Left$(Linedata, 6) = "Title=" Then
        Project.ProjectTitle = Mid$(Linedata, 8, Len(Linedata) - 8)
    ElseIf Left$(Linedata, 5) = "Type=" Then
        Project.ProjectProjectType = Mid$(Linedata, 6)
    ElseIf Left$(Linedata, 8) = "Startup=" Then
        Project.ProjectStartupItem = Mid$(Linedata, 10, Len(Linedata) - 10)

    ElseIf Left$(Linedata, 10) = "ResFile32=" Then
        GetRelatedDocStats Mid$(Linedata, 12, Len(Linedata) - 12)
    ElseIf Left$(Linedata, 11) = "RelatedDoc=" Then
        GetRelatedDocStats Mid$(Linedata, 12)
    ElseIf Left$(Linedata, 7) = "Object=" Then
        GetComponentStats Mid$(Linedata, 8)
    ElseIf Left$(Linedata, 10) = "Reference=" Then
        GetReferenceStats Mid$(Linedata, 10)

    ElseIf Left$(Linedata, 5) = "Form=" Then
        Set ProjectItem = New ClsPrjItem: GetFormStats Mid$(Linedata, 6)
    ElseIf Left$(Linedata, 7) = "Module=" Then
        Set ProjectItem = New ClsPrjItem: GetModuleStats Mid$(Linedata, 8)
    ElseIf Left$(Linedata, 6) = "Class=" Then
        Set ProjectItem = New ClsPrjItem: GetClassStats Mid$(Linedata, 7)
    ElseIf Left$(Linedata, 12) = "UserControl=" Then
        Set ProjectItem = New ClsPrjItem: GetUserControlStats Mid$(Linedata, 13)
    ElseIf Left$(Linedata, 13) = "PropertyPage=" Then
        Set ProjectItem = New ClsPrjItem: GetPropertyPageStats Mid$(Linedata, 14)
    ElseIf Left$(Linedata, 13) = "UserDocument=" Then
        Set ProjectItem = New ClsPrjItem: GetUserDocumentStats Mid$(Linedata, 14)
    ElseIf Left$(Linedata, 9) = "Designer=" Then
        Set ProjectItem = New ClsPrjItem: GetDesignerStats Mid$(Linedata, 10)

    ElseIf Left$(Linedata, 9) = "MajorVer=" Then
        Project.ProjectVersion = Mid$(Linedata, 10) & "."
    ElseIf Left$(Linedata, 9) = "MinorVer=" Then
        Project.ProjectVersion = Mid$(Linedata, 10) & "."
    ElseIf Left$(Linedata, 12) = "RevisionVer=" Then
        Project.ProjectVersion = Mid$(Linedata, 13)

    ElseIf Left$(Linedata, 10) = "ExeName32=" Then
        Project.ProjectEXEFName = Mid$(Linedata, 12, Len(Linedata) - 12)
    ElseIf Left$(Linedata, 7) = "Path32=" Then
        Project.ProjectEXEPath = Mid$(Linedata, 9, Len(Linedata) - 9)
    End If
End Sub

Private Function GetNodeNum(ByVal NodeKey As String) As Long 'Find the node index of a item on the treeview,
    On Error Resume Next                                                                    'from the given key.
    GetNodeNum = FrmResults.TreeView.Nodes.Item(NodeKey).Index
End Function

Private Sub AddProjectToTreeView() 'Sub to add all the standard nodes to the treeview (Forms, Project Info, etc.)
    Dim ProjectRootNode As Long

    If UsesGroup = True Then 'If the project is part of a group, add it to the group node that was added by the Group scanning sub
        FrmResults.TreeView.Nodes.Add 1, tvwChild, "PROJECT_?", "?", "Project"
    Else 'Add it as the first node if not part of a group
        FrmResults.TreeView.Nodes.Add , , "PROJECT_?", "?", "Project"
    End If
    ProjectRootNode = GetNodeNum("PROJECT_?") 'Save time by finding the project's node index only once

    With FrmResults.TreeView.Nodes 'Looks better with a "With" statement
        .Add ProjectRootNode, tvwChild, "PROJECT_?_TITLE", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_VERSION", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_TYPE", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_STARTUPITEM", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_SOURCESAFE", "Source Safe: " & CheckIsSourceSafe(Project.ProjectPath), "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_TOTALS", "Project Totals", "Total"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_LINES", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_LINESNB", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_LINESCOMMENT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_VARIABLES", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_CONSTANTS", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_TYPES", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_ENUMS", "?", "Info"


        .Add ProjectRootNode, tvwChild, "PROJECT_?_PROJINFO", "Project File Information", "LOGFile"

        FileInfo.FindFileInfo Project.ProjectPath, False
        .Add GetNodeNum("PROJECT_?_PROJINFO"), tvwChild, "PROJECT_?_PROJINFO_MODIFIED", "Last Modified: " & FileInfo.LastWriteTime, "Info"
        .Add GetNodeNum("PROJECT_?_PROJINFO"), tvwChild, "PROJECT_?_PROJINFO_ACCESS", "Last Accessed: " & FileInfo.LastAccessTime, "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_FORMS", "Forms", "Form"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_MODULES", "Modules", "Module"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_CLASSES", "Classes", "Class"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_USERCONTROLS", "User Controls", "UserControl"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_USERDOCUMENTS", "User Documents", "UserDocument"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_PROPERTYPAGES", "Property Pages", "PropertyPage"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_DESIGNERS", "Designers", "Designer"

        .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_COUNT", "?", "Info"

        .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_LINES", "0", "Info"

        .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_VARIABLES", "0", "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_RELATEDDOCUMENTS", "Related Documents", "RelatedDocuments"
        .Add GetNodeNum("PROJECT_?_RELATEDDOCUMENTS"), tvwChild, "PROJECT_?_RELATEDDOCUMENTS_COUNT", "Total:", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_REFCOM", "References & Components", "REFCOM"
        .Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COUNT", "Total:", "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_DECDLLS", "Declared DLLs", "DLL"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_SPF", "Subs, Functions & Properties", "SPF"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_SUBS", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_FUNCTIONS", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_PROPERTIES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_EVENTS", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_DECLAREDSUBS", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_DECLAREDFUNCTIONS", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_SUBLINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_FUNCLINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_PROPLINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_AVRSUBLINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_AVRFUNCLINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_SPF"), tvwChild, "PROJECT_?_SPF_AVRPROPLINES", "0", "Info"
    End With
End Sub

Private Sub AddDataToTreeview(ByVal PluralItemName As String, ByVal Linedata As String, ByVal TVPic As String, Optional ResFile As String) 'Adds collected data about the current object to the treeview - easier to manage and smaller
    Dim X As Long, Temp As Long, ParentItemNum As String, LogFileDir As String, i As Long  'than having the same code in all the scan subs

    On Error Resume Next

    With FrmResults.TreeView.Nodes 'Add calculated statistics to treeview
        ParentItemNum = "PROJECT_?_" & PluralItemName & "_" & ProjectItem.PrjItemName

        .Add GetNodeNum("PROJECT_?_" & PluralItemName & ""), tvwChild, ParentItemNum, ProjectItem.PrjItemName, TVPic

        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_CONTROLS", "Controls: " & ProjectItem.PrjItemControls, "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_LINES", "Code Lines (Inc. Blanks): " & ProjectItem.PrjItemCodeLines & " [" & Round((100 / (ProjectItem.PrjItemHybridLines + ProjectItem.PrjItemCommentLines + ProjectItem.PrjItemCodeLines)) * ProjectItem.PrjItemCodeLines, 2) & "%]", "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_LINESNB", "Code Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks & " [" & Round((100 / (ProjectItem.PrjItemHybridLines + ProjectItem.PrjItemCommentLines + ProjectItem.PrjItemCodeLines)) * ProjectItem.PrjItemCodeLinesNoBlanks, 2) & "%]", "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_COMMENTLINES", "Comment Lines: " & ProjectItem.PrjItemCommentLines & " [" & Round((100 / (ProjectItem.PrjItemHybridLines + ProjectItem.PrjItemCommentLines + ProjectItem.PrjItemCodeLines)) * ProjectItem.PrjItemCommentLines, 2) & "%]", "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_HYBRIDLINES", "Hybrid Lines: " & ProjectItem.PrjItemHybridLines & " [" & Round((100 / (ProjectItem.PrjItemHybridLines + ProjectItem.PrjItemCommentLines + ProjectItem.PrjItemCodeLines)) * ProjectItem.PrjItemHybridLines, 2) & "%]", "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_VARIABLES", "Declared Variables: " & ProjectItem.PrjItemVariables, "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_CONSTANTS", "Declared Constants: " & ProjectItem.PrjItemConstants, "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_TYPES", "Declared Types: " & ProjectItem.PrjItemTypes, "Info"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_ENUMS", "Declared Enums: " & ProjectItem.PrjItemEnums, "Info"

        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_SUBS", "Subs", "Method"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_FUNCTIONS", "Functions", "Method"
        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_PROPERTIES", "Properties", "Property"

        If TVPic <> "Module" And TVPic <> "Class" And TVPic <> "Designer" Then .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_EVENTS", "Events", "Event"

        .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_STATMEMENTS", "Statements", "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_FOR", "For/Next: " & ProjectItem.PrjItemStatements(STFOR), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_DO", "Do/Loop: " & ProjectItem.PrjItemStatements(STDO), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_WHILE", "While/Wend: " & ProjectItem.PrjItemStatements(STWHILE), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_IF", "If/End If: " & ProjectItem.PrjItemStatements(STIF), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_SELECT", "Select/End Select: " & ProjectItem.PrjItemStatements(STSELECT), "CodeLoop"

        If CheckForMalicious = 1 Then
            .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_PMC", "Potentially Malicious Code", "BadCode"
            .Add GetNodeNum(ParentItemNum & "_PMC"), tvwChild, ParentItemNum & "_PMC_COUNT", "Total: " & FrmResults.lstBadCode.ListCount, "Info"
            For i = 0 To FrmResults.lstBadCode.ListCount - 1
                .Add GetNodeNum(ParentItemNum & "_PMC"), tvwChild, ParentItemNum & "_PMC_ITEM" & i, FrmResults.lstBadCode.List(i), "BadCode"
            Next
            FrmResults.lstBadCode.Clear
        End If

        .Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_LINES")).Text = Int(.Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_LINES")).Text) + ProjectItem.PrjItemCodeLines

        .Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_VARIABLES")).Text = Int(.Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_VARIABLES")).Text) + ProjectItem.PrjItemVariables

        Project.ProjectLines = ProjectItem.PrjItemCodeLines + ProjectItem.PrjItemHybridLines 'Add to the project total lines (hybrid lines are NOT included as code lines, so must be added separately)
        Project.ProjectLinesNB = ProjectItem.PrjItemCodeLinesNoBlanks + ProjectItem.PrjItemHybridLines 'Ditto

        Temp = GetNodeNum(ParentItemNum & "_SUBS")
        .Add Temp, tvwChild, ParentItemNum & "_SUB_COUNT", "Total: " & ProjectItem.PrjItemItemSubs.Count, "Info"

        For i = 1 To ProjectItem.PrjItemItemSubs.Count
            .Add Temp, tvwChild, ParentItemNum & "_SUB" & i, Left$(ProjectItem.PrjItemItemSubs(i), InStr(1, ProjectItem.PrjItemItemSubs(i), ";") - 1), "Method"
            X = InStr(1, ProjectItem.PrjItemItemSubs(i), ";") + 1
            .Add GetNodeNum(ParentItemNum & "_SUB" & i), tvwChild, ParentItemNum & "_SUB" & i & "_LINES", "Lines (Inc. Blanks): " & Mid$(ProjectItem.PrjItemItemSubs(i), X, InStrRev(ProjectItem.PrjItemItemSubs(i), ":") - X), "Info"
            .Add GetNodeNum(ParentItemNum & "_SUB" & i), tvwChild, ParentItemNum & "_SUB" & i & "_LINESNB", "Lines (No Blanks): " & Mid$(ProjectItem.PrjItemItemSubs(i), InStr(1, ProjectItem.PrjItemItemSubs(i), ":") + 1), "Info"
        Next

        Temp = GetNodeNum(ParentItemNum & "_FUNCTIONS")
        .Add Temp, tvwChild, ParentItemNum & "_FUNCTION_COUNT", "Total: " & ProjectItem.PrjItemItemFunctions.Count, "Info"
        For i = 1 To ProjectItem.PrjItemItemFunctions.Count
            .Add Temp, tvwChild, ParentItemNum & "_FUNCTION" & i, Left$(ProjectItem.PrjItemItemFunctions(i), InStr(1, ProjectItem.PrjItemItemFunctions(i), ";") - 1), "Method"
            X = InStr(1, ProjectItem.PrjItemItemFunctions(i), ";") + 1
            .Add GetNodeNum(ParentItemNum & "_FUNCTION" & i), tvwChild, ParentItemNum & "_FUNCTION" & i & "_LINES", "Lines (Inc. Blanks): " & Mid$(ProjectItem.PrjItemItemFunctions(i), X, InStrRev(ProjectItem.PrjItemItemFunctions(i), ":") - X), "Info"
            .Add GetNodeNum(ParentItemNum & "_FUNCTION" & i), tvwChild, ParentItemNum & "_FUNCTION" & i & "_LINESNB", "Lines (No Blanks): " & Mid$(ProjectItem.PrjItemItemFunctions(i), InStr(1, ProjectItem.PrjItemItemFunctions(i), ":") + 1), "Info"
        Next

        Temp = GetNodeNum(ParentItemNum & "_PROPERTIES")
        .Add Temp, tvwChild, ParentItemNum & "_PROPERTY_COUNT", "Total: " & ProjectItem.PrjItemItemProperties.Count, "Info"
        For i = 1 To ProjectItem.PrjItemItemProperties.Count
            .Add Temp, tvwChild, ParentItemNum & "_PROPERTY" & i, Left$(ProjectItem.PrjItemItemProperties(i), InStr(1, ProjectItem.PrjItemItemProperties(i), ";") - 1), "Property"
            X = InStr(1, ProjectItem.PrjItemItemProperties(i), ";") + 1
            .Add GetNodeNum(ParentItemNum & "_PROPERTY" & i), tvwChild, ParentItemNum & "_PROPERTY" & i & "_LINES", "Lines (Inc. Blanks): " & Mid$(ProjectItem.PrjItemItemProperties(i), X, InStrRev(ProjectItem.PrjItemItemProperties(i), ":") - X), "Info"
            .Add GetNodeNum(ParentItemNum & "_PROPERTY" & i), tvwChild, ParentItemNum & "_PROPERTY" & i & "_LINESNB", "Lines (No Blanks): " & Mid$(ProjectItem.PrjItemItemProperties(i), InStr(1, ProjectItem.PrjItemItemProperties(i), ":") + 1), "Info"
        Next

        If TVPic <> "Module" And TVPic <> "Class" And TVPic <> "Designer" Then
            Temp = GetNodeNum(ParentItemNum & "_EVENTS")
            .Add Temp, tvwChild, ParentItemNum & "_EVENTS_COUNT", "Total: " & ProjectItem.PrjItemItemEvents.Count, "Info"
            For i = 1 To ProjectItem.PrjItemItemEvents.Count
                .Add Temp, tvwChild, ParentItemNum & "_EVENT" & i, ProjectItem.PrjItemItemEvents(i), "Event"
            Next
        End If

        Temp = GetNodeNum(ParentItemNum)
        If ResFile <> "" Then
            .Add Temp, tvwChild, ParentItemNum & "_" & ResFile & "FILE", ResFile & " Resource File", "RelDoc"
            FileInfo.FindFileInfo GetRootDirectory(Project.ProjectPath) & Left$(Linedata, Len(Linedata) - 3) & ResFile, False
            If FileInfo.ByteSize <> "bytes" Then
                .Add GetNodeNum(ParentItemNum & "_" & ResFile & "FILE"), tvwChild, ParentItemNum & "_" & ResFile & "FILE_FILESIZE", "File Size: " & FileInfo.ByteSize, "Info"
            Else
                .Add GetNodeNum(ParentItemNum & "_" & ResFile & "FILE"), tvwChild, ParentItemNum & "_" & ResFile & "FILE_FILESIZE", "File Size: N/A", "Info"
            End If
        End If

        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS", "File Information", "LOGFile"

        Temp = GetNodeNum(ParentItemNum & "_FILESTATS")
        FileInfo.FindFileInfo GetRootDirectory(Project.ProjectPath) & Linedata, False
        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS_FILEMODIFIED", "Last Modified: " & FileInfo.LastWriteTime, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS_FILEOPENED", "Last Accessed: " & FileInfo.LastAccessTime, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS_FILESIZE", "File Size: " & FileInfo.ByteSize, "Info"

        LogFileDir = GetRootDirectory(Project.ProjectPath) & (Left$(Linedata, Len(Linedata) - 4)) & ".log"
        If FileExists(LogFileDir) Then 'Check for a log file
            Temp = GetNodeNum(ParentItemNum)
            .Add Temp, tvwChild, ParentItemNum & "_LOGFILE_" & LogFileDir, "(Double-Click to view Log file)", "LOGFile"
        End If
    End With
End Sub

Private Sub FixPrjTreeViewItems() 'Sub to fix up captions and keys of the treeview's nodes
    Dim ItemKeyNum As Long, QuestionPos As Long, OnePercent As Single, i As Long

    CurrentScanFile "Clean" 'Put the "clean" symbol on the small picture showing the current action

    With FrmSelProject
        .pgbAPB.Value = 0
        .lblScanPhase.Caption = "Treeview Cleanup Phase"
        .lblScanPhase.ForeColor = 6340579
    End With

    DoEvents

    With FrmResults.TreeView.Nodes '"With" statement looks better - cleaner code
        ItemKeyNum = GetNodeNum("PROJECT_?")
        .Item(ItemKeyNum).Text = Project.ProjectName

        ItemKeyNum = GetNodeNum("PROJECT_?_TITLE")
        .Item(ItemKeyNum).Text = "Title: " & Project.ProjectTitle

        ItemKeyNum = GetNodeNum("PROJECT_?_VERSION")
        .Item(ItemKeyNum).Text = "Version: " & Project.ProjectVersion

        ItemKeyNum = GetNodeNum("PROJECT_?_TYPE")
        .Item(ItemKeyNum).Text = "Type: " & Project.ProjectProjectType

        ItemKeyNum = GetNodeNum("PROJECT_?_STARTUPITEM")
        .Item(ItemKeyNum).Text = "Startup Item: " & Project.ProjectStartupItem

        ItemKeyNum = GetNodeNum("PROJECT_?_LINES")
        If Project.ProjectLines = 0 Then
            .Item(ItemKeyNum).Text = "Lines (Inc. Blanks): 0 [0%]"
        Else
            .Item(ItemKeyNum).Text = "Lines (Inc. Blanks): " & Format$(Project.ProjectLines, "###,###,###") & " [" & Round((100 / (Project.ProjectLines + Project.ProjectCommentLines)) * Project.ProjectLines, 2) & "%]"
        End If
        TotalLines = TotalLines + Project.ProjectLines

        ItemKeyNum = GetNodeNum("PROJECT_?_LINESNB")
        If Project.ProjectLinesNB = 0 Then
            .Item(ItemKeyNum).Text = "Lines (No Blanks): 0 [0%]"
        Else
            .Item(ItemKeyNum).Text = "Lines (No Blanks): " & Format$(Project.ProjectLinesNB, "###,###,###") & " [" & Round((100 / (Project.ProjectLines + Project.ProjectCommentLines)) * Project.ProjectLinesNB, 2) & "%]"
        End If

        ItemKeyNum = GetNodeNum("PROJECT_?_LINESCOMMENT")
        If Project.ProjectCommentLines = 0 Then
            .Item(ItemKeyNum).Text = "Lines (Comment): 0 [0%]"
        Else
            .Item(ItemKeyNum).Text = "Lines (Comment): " & Format$(Project.ProjectCommentLines, "###,###,###") & " [" & Round((100 / (Project.ProjectLines + Project.ProjectCommentLines)) * Project.ProjectCommentLines, 2) & "%]"
        End If

        ItemKeyNum = GetNodeNum("PROJECT_?_FORMS_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectForms
        ItemKeyNum = GetNodeNum("PROJECT_?_MODULES_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectModules
        ItemKeyNum = GetNodeNum("PROJECT_?_CLASSES_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectClasses
        ItemKeyNum = GetNodeNum("PROJECT_?_USERCONTROLS_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectUserControls
        ItemKeyNum = GetNodeNum("PROJECT_?_USERDOCUMENTS_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectUserDocuments
        ItemKeyNum = GetNodeNum("PROJECT_?_PROPERTYPAGES_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectPropertyPages
        ItemKeyNum = GetNodeNum("PROJECT_?_DESIGNERS_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectDesigners

        ItemKeyNum = GetNodeNum("PROJECT_?_FORMS_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_MODULES_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_CLASSES_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_USERCONTROLS_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_USERDOCUMENTS_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_PROPERTYPAGES_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_DESIGNERS_LINES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Lines: " & i

        ItemKeyNum = GetNodeNum("PROJECT_?_FORMS_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_MODULES_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_CLASSES_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_USERCONTROLS_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_USERDOCUMENTS_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_PROPERTYPAGES_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i
        ItemKeyNum = GetNodeNum("PROJECT_?_DESIGNERS_VARIABLES")
        i = Int(.Item(ItemKeyNum).Text)
        If i <> 0 Then i = Format$(i, "###,###,###")
        .Item(ItemKeyNum).Text = "Declared Variables: " & i

        ItemKeyNum = GetNodeNum("PROJECT_?_VARIABLES")
        If Project.ProjectVariables = 0 Then
            .Item(ItemKeyNum).Text = "Declared Variables: 0"
        Else
            .Item(ItemKeyNum).Text = "Declared Variables: " & Format$(Project.ProjectVariables, "###,###,###")
        End If
        
        ItemKeyNum = GetNodeNum("PROJECT_?_CONSTANTS")
        If Project.ProjectConstants = 0 Then
            .Item(ItemKeyNum).Text = "Declared Constants: 0"
        Else
            .Item(ItemKeyNum).Text = "Declared Constants: " & Format$(Project.ProjectConstants, "###,###,###")
        End If
        
        ItemKeyNum = GetNodeNum("PROJECT_?_TYPES")
        If Project.ProjectTypes = 0 Then
            .Item(ItemKeyNum).Text = "Declared Types: 0"
        Else
            .Item(ItemKeyNum).Text = "Declared Types: " & Format$(Project.ProjectTypes, "###,###,###")
        End If
        
        ItemKeyNum = GetNodeNum("PROJECT_?_ENUMS")
        If Project.ProjectEnums = 0 Then
            .Item(ItemKeyNum).Text = "Declared Enums: 0"
        Else
            .Item(ItemKeyNum).Text = "Declared Enums: " & Format$(Project.ProjectEnums, "###,###,###")
        End If

        ItemKeyNum = GetNodeNum("PROJECT_?_REFCOM_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & Project.ProjectRefComCount

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_SUBS")
        TotalSubs = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "Total Subs: " & .Item(ItemKeyNum).Text

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_FUNCTIONS")
        TotalFunctions = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "Total Functions: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_PROPERTIES")
        TotalProperties = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "Total Properties: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_EVENTS")
        TotalEvents = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "Total Events: " & ZeroIfNull(Format$(Int(TotalEvents), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_DECLAREDSUBS")
        TotalDecSubs = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "Total Declared Subs: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_DECLAREDFUNCTIONS")
        TotalDecFunctions = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "Total Declared Functions: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_SUBLINES")
        .Item(ItemKeyNum).Text = "Lines in Subs: " & ZeroIfNull(Format$(Project.ProjectSubLines, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_FUNCLINES")
        .Item(ItemKeyNum).Text = "Lines in Functions: " & ZeroIfNull(Format$(Project.ProjectFuncLines, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_PROPLINES")
        .Item(ItemKeyNum).Text = "Lines in Properties: " & ZeroIfNull(Format$(Project.ProjectPropLines, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_RELATEDDOCUMENTS_COUNT")
        .Item(ItemKeyNum).Text = "Total: " & TotalRelDocs

        On Error Resume Next 'Prevent crashes if no SUB/FUNC/PROP lines
        Err.Clear 'If an error occurs it's because there are no S/P/F lines

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_AVRSUBLINES")
        .Item(ItemKeyNum).Text = "Average Lines in Subs: " & Format$(Round(Project.ProjectSubLines / TotalSubs, 2), "###,###.##")
        If Err.Number <> 0 Then .Item(ItemKeyNum).Text = "Average Lines per Sub: 0": Err.Clear

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_AVRFUNCLINES")
        .Item(ItemKeyNum).Text = "Average Lines in Functions: " & Format$(Round(Project.ProjectFuncLines / TotalFunctions, 2), "###,###.##")
        If Err.Number <> 0 Then .Item(ItemKeyNum).Text = "Average Lines per Function: 0": Err.Clear

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_AVRPROPLINES")
        .Item(ItemKeyNum).Text = "Average Lines in Properties: " & Format$(Round(Project.ProjectPropLines / TotalProperties, 2), "###,###.##")
        If Err.Number <> 0 Then .Item(ItemKeyNum).Text = "Average Lines per Property: 0": Err.Clear

        FrmSelProject.pgbAPB.Color = 6340579
        OnePercent = (100 / (.Count - ThisPrgNodeStart))

        If UsesGroup = True Then 'This is made faster by separating the code blocks, rather than evaluation an "If UsesGroup = True" every time
            For ItemKeyNum = ThisPrgNodeStart To .Count 'Go over every node, checking for Related Documents, and replacing the temp project name "?" with the correct
                .Item(ItemKeyNum).Key = Replace$(.Item(ItemKeyNum).Key, "?", Project.ProjectName, 8, 1)

                Do
                    QuestionPos = InStr(1, .Item(ItemKeyNum).Text, "  ") 'Checks for indents, which make line continuations look wrong
                    If QuestionPos <> 0 Then ' "QuestionPos" variable used here only to save on variable requirements
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, QuestionPos - 1) & Mid$(.Item(ItemKeyNum).Text, QuestionPos + 2)
                    Else 'No indents left - exit the loop
                        Exit Do
                    End If
                Loop

                If InStr(1, .Item(ItemKeyNum).Text, "No Blanks):") <> 0 Then
                    If Right$(.Item(ItemKeyNum).Text, 2) = "0" Then
                        If InStr(1, .Item(ItemKeyNum).Parent.Key, "_SUB") <> 0 Then
                            .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                            .Item(ItemKeyNum).Text = "Lines (No Blanks): 0"
                        ElseIf InStr(1, .Item(ItemKeyNum).Parent.Key, "_FUNC") <> 0 Then
                            .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                            .Item(ItemKeyNum).Text = "Lines (No Blanks): 0"
                        ElseIf InStr(1, .Item(ItemKeyNum).Parent.Key, "_PROP") <> 0 Then
                            .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                            .Item(ItemKeyNum).Text = "Lines (No Blanks): 0"
                        End If
                    End If
                End If

                If Int(ItemKeyNum * 0.01) = ItemKeyNum * 0.01 Then FrmSelProject.pgbAPB.Value = OnePercent * (ItemKeyNum - ThisPrgNodeStart)

                If InStr(1, .Item(ItemKeyNum).Text, "[EXT]") <> 0 Then 'Is an external (DLL) call (declared sub or function)
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(150, 150, 150) 'Make key text a light-grey colour
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 10) 'Remove the now unnecessary "[EXT]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[1]" Then 'SPF Colour 1
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(130, 0, 200) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[2]" Then 'SPF Colour 2
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(200, 0, 150) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[3]" Then 'SPF Colour 3
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(10, 150, 10) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[6]" Then 'SPF Colour 4
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(249, 164, 0) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[7]" Then 'SPF Colour 5
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(217, 206, 19) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[8]" Then 'SPF Colour 6
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(19, 217, 192) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[4]" Then 'SPF Colour 7
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(20, 90, 100) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[5]" Then 'SPF Colour 8
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(50, 23, 80) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                End If
            Next
        Else 'Identical to the loop for group projects, but doesn't need to replace the "?" in the node keys
            For ItemKeyNum = 1 To .Count 'Go over every node, checking for Related Documents, and replacing the temp project name "?" with the correct

                Do
                    QuestionPos = InStr(1, .Item(ItemKeyNum).Text, "  ") 'Checks for indents, which make line continuations look wrong
                    If QuestionPos <> 0 Then '"QuestionPos" variable used here only to save on variable requirements
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, QuestionPos - 1) & Mid$(.Item(ItemKeyNum).Text, QuestionPos + 2)
                    Else 'No indents left - exit the loop
                        Exit Do
                    End If
                Loop

                If InStr(1, .Item(ItemKeyNum).Text, "No Blanks):") <> 0 Then
                    If Right$(.Item(ItemKeyNum).Text, 2) = "0" Then
                        If InStr(1, .Item(ItemKeyNum).Parent.Key, "_SUB") <> 0 Then
                            .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                        ElseIf InStr(1, .Item(ItemKeyNum).Parent.Key, "_FUNC") <> 0 Then
                            .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                        ElseIf InStr(1, .Item(ItemKeyNum).Parent.Key, "_PROP") <> 0 Then
                            .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                        End If
                    End If
                End If

                'The colour numbers are out of order because the most frequently used types should be tested before obscure (like static or friend subs) SPF's.
                If Int(ItemKeyNum / 50) = ItemKeyNum / 50 Then FrmSelProject.pgbAPB.Value = OnePercent * (ItemKeyNum - ThisPrgNodeStart)

                If InStr(1, .Item(ItemKeyNum).Text, "[EXT]") <> 0 Then 'Is an external (DLL) call (declared sub or function)
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(150, 150, 150) 'Make key text a light-grey colour
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 10) 'Remove the now unnecessary "[EXT]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[1]" Then 'SPF Colour 1
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(130, 0, 200) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[2]" Then 'SPF Colour 2
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(200, 0, 150) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[3]" Then 'SPF Colour 3
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(10, 150, 10) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[6]" Then 'SPF Colour 4
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(249, 164, 0) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[7]" Then 'SPF Colour 5
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(217, 206, 19) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[8]" Then 'SPF Colour 6
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(19, 217, 192) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[4]" Then 'SPF Colour 7
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(20, 90, 100) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[5]" Then 'SPF Colour 8
                    .Item(ItemKeyNum).Parent.ForeColor = RGB(50, 23, 80) 'Colour Key Text
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                End If
            Next
        End If
    End With
End Sub

Private Sub GetProjectEXEStats()
    Dim EXEFileVer As Variant, PROJFileVer As Variant

    EXENewOrOld = "NF"
    If FileExists(GetRootDirectory(Project.ProjectPath) & Project.ProjectEXEPath & Project.ProjectEXEFName) = False Or Project.ProjectEXEFName = "" Then Exit Sub

    FileInfo.FindFileInfo GetRootDirectory(Project.ProjectPath) & Project.ProjectEXEPath & Project.ProjectEXEFName, False

    With FrmResults.TreeView.Nodes
        .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_EXESTATS", "EXE File", "App"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_NAME", "File Name: " & Project.ProjectEXEFName, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_PATH", "File Path: " & Project.ProjectEXEPath, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_SIZE", "File Size: " & FileInfo.ByteSize, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_CTIME", "Creation Time: " & FileInfo.CreationTime, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_VERSION", "File Version: " & FileInfo.FileVersion, "Info"
    End With

    EXEFileVer = Split(FileInfo.FileVersion, ".")
    PROJFileVer = Split(Project.ProjectVersion, ".")

    EXENewOrOld = CheckVer(PROJFileVer, EXEFileVer)
End Sub

Private Sub GetRelatedDocStats(ByVal FileName As String) 'No stats here, just adding the file to the treeview
    If InStr(1, UCase(FileName), ".RES") <> 0 Then
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_RELATEDDOCUMENTS"), tvwChild, "PROJECT_?_RELATEDDOCUMENT_" & FileName, FileName, "Resource"
    Else
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_RELATEDDOCUMENTS"), tvwChild, "PROJECT_?_RELATEDDOCUMENT_" & FileName, FileName, "RelDoc"
    End If

    TotalRelDocs = TotalRelDocs + 1 'Increase the Related Document count, and add to a string for the report
    RelatedDocumentFNames = RelatedDocumentFNames & FileName & Space$(59 - Len(FileName)) & "|" & vbNewLine

    FileInfo.FindFileInfo FileName, False

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_RELATEDDOCUMENT_" & FileName), tvwChild, "PROJECT_?_RELATEDDOCUMENT_" & FileName & "_SIZE", "File Size: " & FileInfo.ByteSize, "Info"
End Sub

Private Sub GetComponentStats(ByVal FileName As String) 'Get statistics on the component passed
    Dim ClassName As String, OCXFileName As String, ColonPos As Long, HashPos As Long, SecondHashPos As Long

    If InStr(1, FileName, ".vbp", vbTextCompare) Then GoTo IsProjectComponent 'The component is actually a project in a group, skip to the appropriate section

    Project.ProjectRefComCount = 1 'The class actually adds one, rather than setting it as one - saves code

    ColonPos = InStr(1, FileName, ";") 'Finds the colon in the string separating the name and GUID (class name)
    OCXFileName = Mid$(FileName, ColonPos + 1)
    ClassName = Left$(FileName, ColonPos - 1)

    HashPos = InStr(1, ClassName, "#")
    If HashPos = 0 Then GoTo InsertableComponent 'Insertable Components don't show a GUID in the project file
    SecondHashPos = InStrRev(ClassName, "#") - 2

    'Add component and class name to treeview
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName, OCXFileName, "Component"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_CLASSNAME", "Class Name: " & Left$(ClassName, HashPos - 1), "Info"

    ClassName = Left$(ClassName, HashPos - 1) & "\" & Mid$(ClassName, HashPos + 1, Len(ClassName) - SecondHashPos)

    FileInfo.FindFileInfo GetComponentNameFromReg(ClassName), False

    'Add component file info to treeview
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_PATH", "File Path: " & FileInfo.FileName, "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_DESCRIPTION", "Description: " & GetComponentDescFromReg(ClassName), "Info"

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO", "File Information", "LOGFile"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO_SIZE", "File Size: " & FileInfo.ByteSize, "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO_COMPANY", "Company Name: " & FileInfo.CompanyName, "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO_VERSION", "File Version: " & FileInfo.FileVersion, "Info"

    'Add component name and description to an array to make report creation easier
    Project.ProjectRefCom.Add GetComponentNameFromReg(ClassName)
    Project.ProjectRefCom.Add GetComponentDescFromReg(ClassName)
    Exit Sub

IsProjectComponent:                                                                                                                                                                                                                                                                                                                                                                                                                         'Add the project component to the treeview
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & ExtractFileName(FileName), ExtractFileName(FileName), "Project"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & ExtractFileName(FileName)), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & ExtractFileName(FileName) & "_PATH", "Project Path: " & FileName, "Info"

    Project.ProjectRefCom.Add ExtractFileName(FileName) 'Add filename to array
    Project.ProjectRefCom.Add "(Project)" 'Add a unknown description to array
    Exit Sub

InsertableComponent:
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & Left$(FileName, ColonPos - 1), Left$(FileName, ColonPos - 1), "IComponent"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & Left$(FileName, ColonPos - 1)), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & Left$(FileName, ColonPos - 1) & "_PROGRAM", "Parent Program: " & Mid$(FileName, ColonPos + 1), "Info"
End Sub

Private Sub GetReferenceStats(ByVal Linedata As String) 'Get information on a reference (.dll, etc.)
    Dim C As Long, RefName As String, RefDesc As String, FirstHash As Long

    On Error Resume Next

    C = InStrRev(Linedata, "#")
    RefDesc = Mid$(Linedata, C + 1)

    FirstHash = C
    C = InStrRev(Linedata, "#", FirstHash - 1)

    Project.ProjectRefComCount = 1 'This actually adds one - saves code by writing the arithmetic once in the class

    RefName = ExtractFileName(Mid$(Linedata, C + 1, FirstHash - C - 1))

    'Add reference to treeview
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_REFERENCE_" & RefName, RefName, IsSysDLL(RefName), "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_REFERENCE_" & RefName), tvwChild, "PROJECT_?_REFCOM_REFERENCE_" & RefName & "_DESC", "Description: " & RefDesc, "Info"

    'Add reference to array
    Project.ProjectRefCom.Add RefName
    Project.ProjectRefCom.Add RefDesc
End Sub

Private Sub GetDecDllStats(ByVal Linedata As String) 'gets statistics on a declared DLL
    Dim StartQ As Long, EndQ As Long, DLLFName As String

    StartQ = InStr(1, Linedata, """") 'Find start of filename
    EndQ = InStr(StartQ + 1, Linedata, """") 'Find end of filename

    DLLFName = Mid$(Linedata, StartQ + 1, EndQ - StartQ - 1) 'Trim the string to only the filename

    If UCase(Right$(DLLFName, 4)) <> ".DLL" Then DLLFName = DLLFName & ".dll" 'Add a .DLL extension if no extension is present

    If Asc(Right$(DLLFName, 1)) = 34 Then DLLFName = Left$(DLLFName, Len(DLLFName) - 1) 'remove excess characters at end of the string if present

    On Error GoTo AlreadyAdded 'An already added error causes it to skip this section

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS"), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName), DLLFName, IsSysDLL(DLLFName)

    Project.ProjectDecDlls.Add DLLFName

    If FileExists(Environ("windir") & "\System\" & DLLFName) Then
        FileInfo.FindFileInfo Environ("windir") & "\System\" & DLLFName, False
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_PATH", "File Path: " & Environ("windir") & "\System\" & DLLFName, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_SIZE", "File Size: " & FileInfo.ByteSize, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_COMPANYNAME", "Company Name: " & FileInfo.CompanyName, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_DESCRIPTION", "Description: " & FileInfo.FileDescription, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_VERSION", "Version: " & FileInfo.FileVersion, "Info"
    ElseIf FileExists(Environ("windir") & "\System32\" & DLLFName) Then
        FileInfo.FindFileInfo Environ("windir") & "\System32\" & DLLFName, False
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_PATH", "File Path: " & Environ("windir") & "\System32\" & DLLFName, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_SIZE", "File Size: " & FileInfo.ByteSize, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_COMPANYNAME", "Company Name: " & FileInfo.CompanyName, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_DESCRIPTION", "Description: " & FileInfo.FileDescription, "Info"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(DLLFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(DLLFName) & "_VERSION", "Version: " & FileInfo.FileVersion, "Info"
    End If

AlreadyAdded:
End Sub

Private Sub GetFormStats(ByVal Linedata As String) 'Gets the statistics about the form
    Dim FormData As String, JoinLines As Boolean

    Project.ProjectForms = 1 'Add one (adding function stored in class file to save code)
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "Form" 'Change small picture to a Form image

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If

    Do While Not EOF(20)
        Line Input #20, FormData 'Get the data from the file
        FormData = Trim$(FormData)

        'If last line had a line continuation symbol "_", add it to the last line
        If JoinLines = True Then
            FormData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & FormData
            JoinLines = False
        End If

        If Right$(FormData, 1) <> "_" Then 'Is not a line continuation
            LookAtItemLine FormData 'Examine the line
        Else 'Line is continued across multiple lines (continuation)
            JoinLines = True 'Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 'Add to the lines of code statistic
        End If
        ProjectItem.PrjItemPreviousLine = FormData 'Remember the last line, in case of continuation
    Loop

    Close #20

    AddDataToTreeview "FORMS", Linedata, "Form", "FRX"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                      VISUAL BASIC FORM"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls:" & ProjectItem.PrjItemControls
    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
    AddReportText "                  Events:" & ProjectItem.PrjItemItemEvents.Count
End Sub

Private Sub GetModuleStats(ByVal Linedata As String) 'Gets the statistics about the module
    'see "GetFormStatistics" sub for comments

    Dim ModuleData As String, JoinLines As Boolean, Temp As Long

    Temp = InStr(1, Linedata, ";")

    If Temp = 0 Then
        MsgBoxEx "The current project file has corrupt or missing data. Invalid line: """ & Linedata & """." & vbNewLine & "Please open this project with Visual Basic to repair the errors and rescan.", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_" & Linedata, "(Invalid Data) """ & Linedata & """", "Unknown"
        Exit Sub
    End If

    Linedata = GetRootDirectory(Left$(Linedata, Temp - 1)) & Mid$(Linedata, Temp + 2)

    Project.ProjectModules = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "Module"

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If
    Do While Not EOF(20)
        Line Input #20, ModuleData
        ModuleData = Trim$(ModuleData)

        If JoinLines = True Then
            ModuleData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & ModuleData
            JoinLines = False
        End If

        If Right$(ModuleData, 1) <> "_" Then
            LookAtItemLine ModuleData
        Else
            JoinLines = True
            ProjectItem.PrjItemCodeLines = 1
        End If
        ProjectItem.PrjItemPreviousLine = ModuleData
    Loop

    Close #20

    AddDataToTreeview "MODULES", Linedata, "Module"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                     VISUAL BASIC MODULE"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
End Sub

Private Sub GetClassStats(ByVal Linedata As String) 'Gets the statistics about the class module
    'see "GetFormStatistics" sub for comments
    Dim ClassData As String, JoinLines As Boolean, Temp As Long

    Temp = InStr(1, Linedata, ";")

    If Temp = 0 Then
        MsgBoxEx "The current project file has corrupt or missing data. Invalid line: """ & Linedata & """." & vbNewLine & "Please open this project with Visual Basic to repair the errors and rescan.", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_" & Linedata, "(Invalid Data) """ & Linedata & """", "Unknown"
        Exit Sub
    End If

    Linedata = GetRootDirectory(Left$(Linedata, Temp - 1)) & Mid$(Linedata, Temp + 2)

    Project.ProjectClasses = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "Class"

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If

    Do While Not EOF(20)
        Line Input #20, ClassData
        ClassData = Trim$(ClassData)

        If JoinLines = True Then
            ClassData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & ClassData
            JoinLines = False
        End If

        If Right$(ClassData, 1) <> "_" Then
            LookAtItemLine ClassData
        Else
            JoinLines = True
            ProjectItem.PrjItemCodeLines = 1
        End If

        ProjectItem.PrjItemPreviousLine = ClassData
    Loop

    Close #20

    AddDataToTreeview "CLASSES", Linedata, "Class"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC CLASS MODULE"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
End Sub

Private Sub GetUserControlStats(ByVal Linedata As String) 'Gets the statistics about the UserControl
    'see "GetFormStatistics" sub for comments
    Dim UserControlData As String, JoinLines As Boolean

    Project.ProjectUserControls = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "UserControl"

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If

    Do While Not EOF(20)
        Line Input #20, UserControlData
        UserControlData = Trim$(UserControlData)

        If JoinLines = True Then
            UserControlData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & UserControlData
            JoinLines = False
        End If

        If Right$(UserControlData, 1) <> "_" Then
            LookAtItemLine UserControlData
        Else
            JoinLines = True
            ProjectItem.PrjItemCodeLines = 1
        End If
        ProjectItem.PrjItemPreviousLine = UserControlData
    Loop

    Close #20

    AddDataToTreeview "USERCONTROLS", Linedata, "UserControl", "CTX"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC USER CONTROL"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls:" & ProjectItem.PrjItemControls
    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
    AddReportText "                  Events:" & ProjectItem.PrjItemItemEvents.Count
End Sub

Private Sub GetPropertyPageStats(ByVal Linedata As String) 'Gets the statistics about the Property Page
    'see "GetFormStatistics" sub for comments
    Dim PropertyPageData As String, JoinLines As Boolean

    Project.ProjectPropertyPages = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "PropertyPage"

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If

    Do While Not EOF(20)
        Line Input #20, PropertyPageData
        PropertyPageData = Trim$(PropertyPageData)

        If JoinLines = True Then
            PropertyPageData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & PropertyPageData
            JoinLines = False
        End If

        If Right$(PropertyPageData, 1) <> "_" Then
            LookAtItemLine PropertyPageData
        Else
            JoinLines = True
            ProjectItem.PrjItemCodeLines = 1
        End If
        ProjectItem.PrjItemPreviousLine = PropertyPageData
    Loop

    Close #20

    AddDataToTreeview "PROPERTYPAGES", Linedata, "PropertyPage"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC PROPERTY PAGE"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls:" & ProjectItem.PrjItemControls
    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
    AddReportText "                  Events:" & ProjectItem.PrjItemItemEvents.Count
End Sub

Private Sub GetDesignerStats(ByVal Linedata As String) 'Gets the statistics about the Designer
    Dim DesignerData As String, JoinLines As Boolean

    Project.ProjectDesigners = 1 'Add one (adding function stored in class file to save code)
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "Designer" 'Change small picture to a Form image

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If

    Do While Not EOF(20)
        Line Input #20, DesignerData 'Get the data from the file
        DesignerData = Trim$(DesignerData)

        'If last line had a line continuation symbol "_", add it to the last line
        If JoinLines = True Then
            DesignerData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & DesignerData
            JoinLines = False
        End If

        If Right$(DesignerData, 1) <> "_" Then 'Is not a line continuation
            LookAtItemLine DesignerData 'Examine the line
        Else 'Line is continued across multiple lines (continuation)
            JoinLines = True 'Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 'Add to the lines of code statistic
        End If
        ProjectItem.PrjItemPreviousLine = DesignerData 'Remember the last line, in case of continuation
    Loop

    Close #20

    AddDataToTreeview "DESIGNERS", Linedata, "Designer"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                    VISUAL BASIC DESIGNER"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls:" & ProjectItem.PrjItemControls
    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
End Sub

Private Sub GetUserDocumentStats(ByVal Linedata As String) 'Gets the statistics about the User Document
    'see "GetFormStatistics" sub for comments
    Dim UserDocumentData As String, JoinLines As Boolean

    Project.ProjectUserDocuments = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "UserDocument"

    If Mid$(Linedata, 2, 1) = ":" Then
        If FileExists(Linedata) = False Then
            If ShowFNFerrors Then MsgBoxEx "File """ & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_" & Linedata, Linedata, "Unknown"
        Else
            Open Linedata For Input As #20
        End If
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Linedata) = False Then 'File not found, show an error
            If ShowFNFerrors Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & Linedata & """ not found!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_" & Linedata, Linedata, "Unknown"
            Exit Sub
        Else
            Open GetRootDirectory(Project.ProjectPath) & Linedata For Input As #20
        End If
    End If

    Do While Not EOF(20)
        Line Input #20, UserDocumentData
        UserDocumentData = Trim$(UserDocumentData)

        If JoinLines = True Then
            UserDocumentData = Left$(ProjectItem.PrjItemPreviousLine, Len(ProjectItem.PrjItemPreviousLine) - 1) & UserDocumentData
            JoinLines = False
        End If

        If Right$(UserDocumentData, 1) <> "_" Then
            LookAtItemLine UserDocumentData
        Else
            JoinLines = True
            ProjectItem.PrjItemCodeLines = 1
        End If
        ProjectItem.PrjItemPreviousLine = UserDocumentData
    Loop

    Close #20

    AddDataToTreeview "USERDOCUMENTS", Linedata, "UserDocument"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC USER DOCUMENT"
    AddReportText "============================================================="
    AddReportText "               File Name:" & ExtractFileName(GetRootDirectory(Project.ProjectPath) & Linedata)
    AddReportText "                    Name:" & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total):" & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks):" & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment):" & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid):" & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls:" & ProjectItem.PrjItemControls
    AddReportText "               Variables:" & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines:" & ProjectItem.PrjItemItemSubs.Count
    AddReportText "               Functions:" & ProjectItem.PrjItemItemFunctions.Count
    AddReportText "              Properties:" & ProjectItem.PrjItemItemProperties.Count
    AddReportText "                  Events:" & ProjectItem.PrjItemItemEvents.Count
End Sub

Private Function GetComponentDescFromReg(ByVal ClassID As String) As String 'Get component description from the registry by it's classID
    Dim RegFDesc As String

    ClassID = Trim$(ClassID)

    RegFDesc = ModRegistry.QueryValue(&H80000000, "TypeLib\" & ClassID, "") 'Find the component's registry folder

    If RegFDesc = "" Then 'Not found
        GetComponentDescFromReg = "Unknown/Unregistered (ClassID: " & ClassID & ")"
    Else 'Found, return its value
        GetComponentDescFromReg = RegFDesc
    End If
End Function

Private Function GetComponentNameFromReg(ByVal ClassID As String) As String 'Get the component name from the registry
    Dim RegFName As String

    ClassID = Trim$(ClassID)

    ' Component name stored in "HKEY_CLASSES_ROOT\TypeLib\{ClassID}\0\Win32"
    RegFName = ModRegistry.QueryValue(&H80000000, "TypeLib\" & ClassID & "\0\Win32", "")

    If RegFName = "" Then 'Not found
        GetComponentNameFromReg = "Unknown/Unregistered (ClassID: " & ClassID & ")"
    Else 'Found, return its value
        GetComponentNameFromReg = RegFName
    End If
End Function

Private Function CheckIsSourceSafe(ByVal ProjectPath As String) As Boolean
    Dim SCCdata$
    Dim IsRightProject As Boolean

    'Source safe data stored in a file called "MSSCCPRJ.SCC" in the same directory
    'as the project file - but only if SourceSafe is installed. As a result, don't show
    'and error if the file isn't there.

    If FileExists(GetRootDirectory(ProjectPath) & "MSSCCPRJ.SCC") = False Then Exit Function
    Open GetRootDirectory(ProjectPath) & "MSSCCPRJ.SCC" For Input As #10

    Do While Not EOF(10)
        Line Input #10, SCCdata$
        If SCCdata$ = "[" & ExtractFileName(ProjectPath) & "]" Then 'Group files use the same SourceSafe file
            IsRightProject = True                                   'so find the correct project's info
        End If

        If IsRightProject = True And Left$(SCCdata$, 17) = "SCC_Project_Name=" Then 'Data stored in this string
            If Mid$(SCCdata$, 18) <> "this project is not under source code control" Then
                CheckIsSourceSafe = True 'Is source safe
            Else
                CheckIsSourceSafe = False 'Not source safe
            End If
        End If
    Loop

    Close #10
End Function

Private Sub LookAtItemLine(ByVal Linedata As String) 'Examines a line of data from a VB file - this is called by all the
    'GetForm/GetModule/etc. subs to save code and increase speed
    Dim SPFName As String, Vars As Long, X As Long

    On Error Resume Next

    If ProjectItem.PrjItemSeenAttributes = False Then
        'If the file contains control information (Form, User Control, etc.) don't add to statistics until actual code
        If ProjectItem.PrjItemInControls = True Then
            If Left$(Linedata, 6) = "Begin " Then
                ProjectItem.PrjItemControls = 1
                Exit Sub
            End If
        End If

        If Left$(ProjectItem.PrjItemPreviousLine, 10) = "Attribute " Then
            If InStr(1, Linedata, "Attribute", vbBinaryCompare) = 0 Then ProjectItem.PrjItemSeenAttributes = True
        End If

        If InStr(1, Linedata, "VB") <> 0 Then
            If Left$(Linedata, 9) = "Begin VB." Then ProjectItem.PrjItemInControls = True
            If Left$(Linedata, 20) = "Attribute VB_Name = " Then
                ProjectItem.PrjItemName = Mid$(Linedata, 22, Len(Linedata) - 22)
                ProjectItem.PrjItemInControls = False 'This is the last line of header info
            End If
        End If

        Exit Sub 'Don't analyse until all the header info is looked at
    End If

    Linedata = RemLineNumber(Linedata) 'Remove line numbers (if necessary)

    If IsCommentLine(Linedata) Then
        ProjectItem.PrjItemCommentLines = 1 'Add 1 to statistic (adding code is in class to save on space and increase speed)
        Project.ProjectCommentLines = 1 'Add 1 to total statistic
        CurrSPFLines = CurrSPFLines + 1
        Exit Sub
    Else
        If IsHybridLine(Linedata) <> 0 Then 'Checks is hybrid code/comment line
            ProjectItem.PrjItemHybridLines = 1 'Add 1 to statistic (adding code is in class to save on space and increase speed)
            Linedata = Mid$(Linedata, 1, IsHybridLine(Linedata))
        Else
            ProjectItem.PrjItemCodeLines = 1 'Add 1 to statistic (adding code is in class to save on space and increase speed)
        End If

        If Linedata = "" Then
            CurrSPFLines = CurrSPFLines + 1
            Exit Sub      'If it's a blank line, don't scan it - it just wastes time
        End If
    End If

    ProjectItem.PrjItemCodeLinesNoBlanks = 1 'Add 1 to statistic (adding code is in class to save on space and increase speed)
    If InSub = True Then
        Project.ProjectSubLines = 1
        CurrSPFLinesNB = CurrSPFLinesNB + 1
    End If
    If InFunction = True Then
        Project.ProjectFuncLines = 1
        CurrSPFLinesNB = CurrSPFLinesNB + 1
    End If
    If InProperty = True Then
        Project.ProjectPropLines = 1
        CurrSPFLinesNB = CurrSPFLinesNB + 1
    End If

    If CheckIsStatement(Linedata) = True Then GoTo SkipCheck
    If CheckIsConstTypeEnum(Linedata) = True Then GoTo SkipCheck

    SPFName = CheckIsSub(Linedata) 'If the line is a sub, add it to the array for sorting
    If SPFName <> "" Then
        If ShowSPFParams = False Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)
        If InStr(1, SPFName, "Lib") <> 0 Then
            GetDecDllStats SPFName
            InSub = False
            ProjectItem.AddSPF SPFName & ";N/A [EXT]:N/A", SPF_Sub
            CurrSPFLines = 0
        End If
        CurrSPFName = SPFName
        GoTo SkipCheck
    End If

    SPFName = CheckIsFunction(Linedata) 'If the line is a function, add it to the list for sorting
    If SPFName <> "" Then
        If ShowSPFParams = False Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)
        If InStr(1, SPFName, "Lib") <> 0 Then
            GetDecDllStats SPFName
            InFunction = False
            ProjectItem.AddSPF SPFName & ";N/A [EXT]:N/A", SPF_Function
            CurrSPFLines = 0
        End If
        CurrSPFName = SPFName
        GoTo SkipCheck
    End If

    SPFName = CheckIsProperty(Linedata) 'If the line is a property, add it to the array for sorting
    If SPFName <> "" Then
        If ShowSPFParams = False Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)
        CurrSPFName = SPFName
        GoTo SkipCheck
    End If

    If Left$(Linedata, 6) = "Event " Then
        SPFName = Mid$(Linedata, 7)
        If ShowSPFParams = False Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)
        ProjectItem.AddSPF SPFName, SPF_Event
        IncrementTreeViewPrjEvents
        GoTo SkipCheck
    End If

    If CheckForMalicious = 1 Then CheckIsMalicious Linedata 'If the Potentially Malicious Code checking option is enabled, check the line

    Vars = CheckIsVariable(Linedata) 'Check if the line is a variable
    If Vars <> 0 Then
        ProjectItem.PrjItemVariables = Vars
        Project.ProjectVariables = Vars
        GoTo SkipCheck
    End If

    If InStr(1, Linedata, "CreateObject") <> 0 Then
        X = InStr(1, Linedata, "CreateObject") 'CreatObjects are rare, so it's the last SPF thing to check
        If X < 2 Or Mid$(Linedata, X - 1, 1) = " " Then 'Correct start for the CreateObject statement
            If Mid$(Linedata, X + 12, 1) = "(" Then
                SPFName = Mid$(Linedata, X + 13, InStr(X + 13, Linedata, ")") - X - 13)
                Project.ProjectCreateObjects.Add SPFName
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS"), tvwChild, "PROJECT_?_DECDLLS_" & UCase(SPFName), SPFName, "CreateObject"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(SPFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(SPFName) & "_INFO", "VB 'CreateObject' Statement", "Info"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(SPFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(SPFName) & "_INFOLINE", "Code Line: " & Linedata, "Info"
            End If
        End If
    End If

SkipCheck:
End Sub

Private Function IsCommentLine(ByVal Linedata As String) As Boolean 'Returns TRUE if the entered line contains ONLY a comment.
    If Left$(Linedata, 1) = "'" Then
        IsCommentLine = True
        Exit Function
    End If

    If Left$(Linedata, 4) = "Rem" Then IsCommentLine = True
End Function

Private Function IsHybridLine(ByVal Linedata As String) As Long 'Returns TRUE if the entered string contains both code and comment.
    Dim InString As Boolean, i As Long

    If InStr(1, Linedata, "'") = 0 Then Exit Function ' Comments after code can only use the "'" symbol, not the REM statement. Check to see
    'if the line contains the "'" character, otherwise skip the sub to save time.
    If InStr(1, Linedata, """") = 0 Then
        IsHybridLine = InStrRev(Linedata, "'")
        Exit Function 'If there's no quotes, the line is automatically a hybrid
    End If

    For i = 2 To Len(Linedata)
        If Mid$(Linedata, i, 1) = """" Then InString = Not InString ' Changes the InString boolean variable as DeepLook finds a " symbol. This prevents

        If Mid$(Linedata, i, 1) = "'" Then
            If InString = False Then IsHybridLine = i
            Exit Function             ' the program mis-interpreting ' symbols in strings as comments.
        End If
    Next
End Function

Private Function CheckIsSub(ByVal Linedata As String) As String ' Returns TRUE if the line is a Sub
    If InStr(1, Linedata, "Sub", vbBinaryCompare) <> 0 Then
        If Left$(Linedata, 12) = "Private Sub " Then
            CheckIsSub = Mid$(Linedata, 12)
            IncrementTreeViewPrjSubs
            CurrSPFColour = 1
        ElseIf Left$(Linedata, 11) = "Public Sub " Then
            CheckIsSub = Mid$(Linedata, 11)
            IncrementTreeViewPrjSubs
            CurrSPFColour = 2
        ElseIf Left$(Linedata, 4) = "Sub " Then
            CheckIsSub = Mid$(Linedata, 4)
            IncrementTreeViewPrjSubs
            CurrSPFColour = 3
        ElseIf Left$(Linedata, 20) = "Private Declare Sub " Then
            CheckIsSub = Mid$(Linedata, 20)
            IncrementTreeViewPrjSubs True
        ElseIf Left$(Linedata, 19) = "Public Declare Sub " Then
            CheckIsSub = Mid$(Linedata, 19)
            IncrementTreeViewPrjSubs True
        ElseIf Left$(Linedata, 12) = "Declare Sub " Then
            CheckIsSub = Mid$(Linedata, 12)
            IncrementTreeViewPrjSubs True
        ElseIf Left$(Linedata, 11) = "Friend Sub " Then
            CheckIsSub = Mid$(Linedata, 11)
            IncrementTreeViewPrjSubs True
            CurrSPFColour = 4
        ElseIf Left$(Linedata, 11) = "Static Sub " Then
            CheckIsSub = Mid$(Linedata, 11)
            IncrementTreeViewPrjSubs True
            CurrSPFColour = 5
        End If

        If CheckIsSub <> "" Then
            InSub = True
            CurrSPFLines = 0
            CurrSPFLinesNB = 0
            If InStr(1, CheckIsSub, "'") <> 0 Then CheckIsSub = Left$(CheckIsSub, InStr(1, CheckIsSub, "'") - 1)
        Else
            If Left$(Linedata, 7) = "End Sub" Then
                InSub = False
                ProjectItem.AddSPF CurrSPFName & ";" & (CurrSPFLines + CurrSPFLinesNB - 1) & "[" & CurrSPFColour & "]:" & (CurrSPFLinesNB - 1), SPF_Sub
            End If
        End If
    End If
End Function

Private Function CheckIsFunction(ByVal Linedata As String) As String ' Returns TRUE if the line is a Function
    If InStr(1, Linedata, "Function", vbBinaryCompare) Then
        If InStr(1, Linedata, "(") <> 0 Then
            If InStr(1, Linedata, "Declare ", vbBinaryCompare) Then
                If Mid$(Linedata, 1, 25) = "Private Declare Function " Then
                    CheckIsFunction = Mid$(Linedata, 25)
                    IncrementTreeViewPrjFunctions True
                ElseIf Mid$(Linedata, 1, 24) = "Public Declare Function " Then
                    CheckIsFunction = Mid$(Linedata, 24)
                    IncrementTreeViewPrjFunctions True
                ElseIf Mid$(Linedata, 1, 17) = "Declare Function " Then
                    CheckIsFunction = Mid$(Linedata, 17)
                    IncrementTreeViewPrjFunctions True
                End If
            Else
                If Mid$(Linedata, 1, 7) = "Private" Then
                    CheckIsFunction = Mid$(Linedata, 17)
                    IncrementTreeViewPrjFunctions
                    CurrSPFColour = 1
                ElseIf Mid$(Linedata, 1, 6) = "Public" Then
                    CheckIsFunction = Mid$(Linedata, 16)
                    IncrementTreeViewPrjFunctions
                    CurrSPFColour = 2
                ElseIf Mid$(Linedata, 1, 8) = "Function" Then
                    CheckIsFunction = Mid$(Linedata, 9)
                    IncrementTreeViewPrjFunctions
                    CurrSPFColour = 3
                Else

                End If
            End If
        End If

        If CheckIsFunction <> "" Then
            InFunction = True
            If InStr(1, CheckIsFunction, "'") <> 0 Then CheckIsFunction = Left$(CheckIsFunction, InStr(1, CheckIsFunction, "'") - 1)
            CurrSPFName = CheckIsFunction
            CurrSPFLines = 0
            CurrSPFLinesNB = 0
        Else
            If Left$(Linedata, 12) = "End Function" Then
                InFunction = False
                ProjectItem.AddSPF CurrSPFName & ";" & (CurrSPFLines + CurrSPFLinesNB - 1) & "[" & CurrSPFColour & "]:" & (CurrSPFLinesNB - 1), SPF_Function
            End If
        End If
    End If
End Function

Private Function CheckIsProperty(ByVal Linedata As String) As String ' Returns TRUE if the line is a Property
    If InStr(1, Linedata, "Property", vbBinaryCompare) Then
        If Mid$(Linedata, 1, 17) = "Private Property " Then
            If Mid$(Linedata, 18, 3) = "Let" Then
                CurrSPFColour = 6
            ElseIf Mid$(Linedata, 18, 3) = "Get" Then
                CurrSPFColour = 7
            ElseIf Mid$(Linedata, 18, 3) = "Set" Then
                CurrSPFColour = 8
            End If
            CheckIsProperty = Mid$(Linedata, 21)
            IncrementTreeViewPrjProperties
        ElseIf Left$(Linedata, 16) = "Public Property " Then
            If Mid$(Linedata, 17, 3) = "Let" Then
                CurrSPFColour = 6
            ElseIf Mid$(Linedata, 17, 3) = "Get" Then
                CurrSPFColour = 7
            ElseIf Mid$(Linedata, 17, 3) = "Set" Then
                CurrSPFColour = 8
            End If
            CheckIsProperty = Mid$(Linedata, 20)
            IncrementTreeViewPrjProperties
        ElseIf Left$(Linedata, 9) = "Property " Then
            If Mid$(Linedata, 10, 3) = "Let" Then
                CurrSPFColour = 6
            ElseIf Mid$(Linedata, 10, 3) = "Get" Then
                CurrSPFColour = 7
            ElseIf Mid$(Linedata, 10, 3) = "Set" Then
                CurrSPFColour = 8
            End If
            CheckIsProperty = Mid$(Linedata, 13)
            IncrementTreeViewPrjProperties
        End If

        If InStr(1, CheckIsProperty, "(") = 0 Then CheckIsProperty = ""

        If CheckIsProperty <> "" Then
            InProperty = True
            CurrSPFLines = 0
            CurrSPFLinesNB = 0
            If InStr(1, CheckIsProperty, "'") <> 0 Then CheckIsProperty = Left$(CheckIsProperty, InStr(1, CheckIsProperty, "'") - 1)
            CurrSPFName = CheckIsProperty
        Else
            If Left$(Linedata, 12) = "End Property" Then
                InProperty = False
                ProjectItem.AddSPF CurrSPFName & ";" & (CurrSPFLines + CurrSPFLinesNB - 1) & "[" & CurrSPFColour & "]:" & (CurrSPFLinesNB - 1), SPF_Property
            End If
        End If

    End If
End Function

Private Function CheckIsVariable(ByVal Linedata As String) As Long ' Returns TRUE if the line is a Variable
    Dim i As Long

    If Left$(Linedata, 4) = "Dim " Then
        CheckIsVariable = 1
    ElseIf Left$(Linedata, 7) = "Static " Then
        CheckIsVariable = 1
    ElseIf Left$(Linedata, 7) = "Global " Then
        CheckIsVariable = 1
    ElseIf Left$(Linedata, 8) = "Private " Then
        CheckIsVariable = 1
    ElseIf Left$(Linedata, 7) = "Public " Then
        CheckIsVariable = 1
    Else
        Exit Function
    End If

    If InStr(1, Linedata, " Sub ") <> 0 Then
        Exit Function
    ElseIf InStr(1, Linedata, " Function ") <> 0 Then
        Exit Function
    ElseIf InStr(1, Linedata, " Property ") <> 0 Then
        Exit Function
    ElseIf InStr(1, Linedata, " Type ") <> 0 Then
        Exit Function
    ElseIf InStr(1, Linedata, " Enum ") <> 0 Then
        Exit Function
    ElseIf InStr(1, Linedata, " Event ") <> 0 Then
        Exit Function
    ElseIf InStr(1, Linedata, " Const ") <> 0 And IncludeConsts = True Then
        Exit Function
    ElseIf InStr(1, Linedata, " WithEvents ") <> 0 Then
        Exit Function
    End If

    If Left$(Linedata, 7) = "Global " Or Left$(Linedata, 7) = "Public " Then ' Global variables are added to a collection
        If InStr(1, Linedata, "'") Then Linedata = Mid$(Linedata, 1, InStr(1, Linedata, "'") - 1)

        If InStr(1, Linedata, ",") = 0 Then
            If InStr(1, Linedata, " As ") <> 0 Then
                GlobalVars.Add TrimJunk(Mid$(Linedata, 8, InStr(1, Linedata, " As ") - 8))
                GlobalVarsLoc.Add ProjectItem.PrjItemName
            Else
                GlobalVars.Add TrimJunk(Mid$(Linedata, 8))
                GlobalVarsLoc.Add ProjectItem.PrjItemName
            End If
        Else
            Dim VarNames As Variant, Var As Variant
            
            Linedata = Mid$(Linedata, 8)
            If InStr(1, Linedata, ":") Then Linedata = Mid$(Linedata, 1, InStr(1, Linedata, ":"))
            
            VarNames = Split(Linedata, ",")
            For Var = LBound(VarNames) To UBound(VarNames)
                If InStr(1, VarNames(Var), " As ") <> 0 Then
                    GlobalVars.Add TrimJunk(Mid$(VarNames(Var), 1, InStr(1, VarNames(Var), " As ")))
                    GlobalVarsLoc.Add ProjectItem.PrjItemName
                Else
                    GlobalVars.Add TrimJunk(Mid$(VarNames(Var), 1))
                    GlobalVarsLoc.Add ProjectItem.PrjItemName
                End If
            Next
        End If
    Else ' Other varibles need less processing
        If InStr(1, Linedata, ",") <> 0 Then
            For i = 1 To Len(Linedata)
                If Mid$(Linedata, i, 1) = "'" Then Exit For
                If Mid$(Linedata, i, 1) = "," Then CheckIsVariable = CheckIsVariable + 1
            Next
        End If
    End If
End Function

Private Function CheckIsConstTypeEnum(ByVal Linedata As String) As Boolean
    Dim MidLen As Long

    If Left$(Linedata, 7) = "Global " Then
        MidLen = 8
    ElseIf Left$(Linedata, 8) = "Private " Then
        MidLen = 9
    ElseIf Left$(Linedata, 7) = "Public " Then
        MidLen = 8
    ElseIf Left$(Linedata, 7) = "Const " Then
        MidLen = 1
    ElseIf Left$(Linedata, 7) = "Type " Then
        MidLen = 1
    ElseIf Left$(Linedata, 8) = "Enum " Then
        MidLen = 1
    Else
        Exit Function
    End If

    If Mid$(Linedata, MidLen, 5) = "Type " Then
        ProjectItem.PrjItemTypes = 1
        Project.ProjectTypes = 1
        CheckIsConstTypeEnum = True
        Exit Function
    ElseIf Mid$(Linedata, MidLen, 5) = "Enum " Then
        ProjectItem.PrjItemEnums = 1
        Project.ProjectEnums = 1
        CheckIsConstTypeEnum = True
        Exit Function
    ElseIf Mid$(Linedata, MidLen, 6) = "Const " Then
        ProjectItem.PrjItemConstants = 1
        Project.ProjectConstants = 1
        CheckIsConstTypeEnum = True
    End If
End Function

Private Function CheckIsStatement(ByVal Linedata As String) As Boolean
    CheckIsStatement = True
    If Left$(Linedata, 3) = "If " Then
        ProjectItem.AddToStatement STIF
    ElseIf Left$(Linedata, 4) = "For " Then
        ProjectItem.AddToStatement STFOR
    ElseIf Left$(Linedata, 12) = "Select Case " Then
        ProjectItem.AddToStatement STSELECT
    ElseIf Left$(Linedata, 3) = "Do " Then
        ProjectItem.AddToStatement STDO
    ElseIf Left$(Linedata, 6) = "While " Then
        ProjectItem.AddToStatement STWHILE
    Else
        CheckIsStatement = False
    End If
End Function

Private Sub CheckIsMalicious(ByVal Linedata As String) ' Checks if a line contains code that could cause trojan or virus like actions
    Dim BLPos As Long, NoAdd As Boolean, i As Long
    Dim CorrectStart As Boolean, CorrectEnd As Boolean, BCLegnth As Long

    ' Since this is not a dedicated PMC scanner, I havn't spent much time making
    ' this - it errs on the side of safety by giving more false positives - I could
    ' extend it but it would be _very_ slow.

    '            SCAN FOR KEYWORDS
    BLPos = InStr(1, Linedata, "Kill")
    If BLPos <> 0 Then BCLegnth = 4: GoTo Finish
    BLPos = InStr(1, Linedata, "Delete")
    If BLPos <> 0 Then BCLegnth = 6: GoTo Finish
    BLPos = InStr(1, Linedata, "FileCopy")
    If BLPos <> 0 Then BCLegnth = 8: GoTo Finish
    BLPos = InStr(1, Linedata, "LoadResData")
    If BLPos <> 0 Then BCLegnth = 11: GoTo Finish
    BLPos = InStr(1, Linedata, "RegOpenKeyEx")
    If BLPos <> 0 Then BCLegnth = 12: GoTo Finish
    BLPos = InStr(1, Linedata, "ShellExecute")
    If BLPos <> 0 Then BCLegnth = 12: GoTo Finish
    BLPos = InStr(1, Linedata, "WNETENUMCACHEDPASSWORDS", vbTextCompare)
    If BLPos <> 0 Then BCLegnth = 23: GoTo Finish
    BLPos = InStr(1, Linedata, "Move")
    If BLPos <> 0 Then BCLegnth = 4: GoTo Finish
    BLPos = InStr(1, Linedata, "WNETADDCONNECTION", vbTextCompare)
    If BLPos <> 0 Then BCLegnth = 17: GoTo Finish
    BLPos = InStr(1, Linedata, "Append")
    If BLPos <> 0 Then BCLegnth = 6: GoTo Finish
    BLPos = InStr(1, Linedata, "Output")
    If BLPos <> 0 Then BCLegnth = 6: GoTo Finish
    BLPos = InStr(1, Linedata, "Binary")
    If BLPos <> 0 Then BCLegnth = 6: GoTo Finish
    BLPos = InStr(1, Linedata, "Connect")
    If BLPos <> 0 Then BCLegnth = 7: GoTo Finish
    BLPos = InStr(1, Linedata, "Listen")
    If BLPos <> 0 Then BCLegnth = 6: GoTo Finish
    BLPos = InStr(1, Linedata, "Winsock")
    If BLPos <> 0 Then BCLegnth = 7: GoTo Finish

    Exit Sub ' No keywords found

Finish:                                                                                                                                                                                                                                                                                                                                                                                 ' Keyword found

    If BLPos = 1 Then BLPos = 2
    If Mid$(Linedata, BLPos - 1, 1) = " " Or Mid$(Linedata, BLPos - 1, 1) = "." Or Mid$(Linedata, BLPos - 1, 1) = "(" Or BLPos = 2 Then CorrectStart = True ' Keyword must have a bracket, space, period or be the start of a line

    If BCLegnth = Len(Linedata) Then
        CorrectEnd = True
    Else
        If InStr(1, " ().", Mid$(Linedata, BLPos + BCLegnth, 1)) <> 0 Then CorrectEnd = True  ' Must have a bracket, space or nothing after it
    End If

    If CorrectStart = True Then
        If CorrectEnd = True Then ' Both requirements satisfied
            NoAdd = False

            For i = 0 To FrmResults.lstBadCode.ListCount - 1
                If FrmResults.lstBadCode.List(i) = Linedata Then
                    NoAdd = True
                    Exit For ' Already added, set the variable to TRUE
                End If
            Next

            If NoAdd = False Then ' Variable FALSE (not added), so add it to the list for sorting
                FrmResults.lstBadCode.AddItem Linedata
            End If
        End If
    End If
End Sub

Private Sub IncrementTreeViewPrjSubs(Optional Declared As Boolean)
    If Declared = True Then
        NodeNum = GetNodeNum("PROJECT_?_SPF_DECLAREDSUBS") ' Get the node index of the "Total Declared Subs"
    Else
        NodeNum = GetNodeNum("PROJECT_?_SPF_SUBS") ' Get the node index of the "Total Subs"
    End If

    If NodeNum = 0 Then Exit Sub

    FrmResults.TreeView.Nodes(NodeNum).Text = Int(FrmResults.TreeView.Nodes(NodeNum).Text) + 1
End Sub

Private Sub IncrementTreeViewPrjFunctions(Optional Declared As Boolean)
    If Declared = True Then
        NodeNum = GetNodeNum("PROJECT_?_SPF_DECLAREDFUNCTIONS") ' Get the node index of the "Total Declared Functions"
    Else
        NodeNum = GetNodeNum("PROJECT_?_SPF_FUNCTIONS") ' Get the node index of the "Total Functions"
    End If

    If NodeNum = 0 Then Exit Sub

    FrmResults.TreeView.Nodes(NodeNum).Text = Int(FrmResults.TreeView.Nodes(NodeNum).Text) + 1
End Sub

Private Sub IncrementTreeViewPrjProperties()
    NodeNum = GetNodeNum("PROJECT_?_SPF_PROPERTIES") ' Get the node index of the "Total Properties"

    If NodeNum = 0 Then Exit Sub

    FrmResults.TreeView.Nodes(NodeNum).Text = Int(FrmResults.TreeView.Nodes(NodeNum).Text) + 1
End Sub

Private Sub IncrementTreeViewPrjEvents()
    NodeNum = GetNodeNum("PROJECT_?_SPF_EVENTS") ' Get the node index of the "Total Events"

    If NodeNum = 0 Then Exit Sub

    FrmResults.TreeView.Nodes(NodeNum).Text = Int(FrmResults.TreeView.Nodes(NodeNum).Text) + 1
End Sub

Private Sub CurrentScanFile(ByVal FileType As String) ' Changes the little icon on the frmSelProject (if turned on) to
    If ShowCurrItemPic = False Then Exit Sub    ' indicate what type of file is being scanned

    With FrmSelProject.imgCurrScanObjType
        Select Case FileType
            Case "Project"
                .Picture = FrmResults.ilstImages.ListImages(1).Picture
            Case "Group"
                .Picture = FrmResults.ilstImages.ListImages(2).Picture
            Case "Form"
                .Picture = FrmResults.ilstImages.ListImages(8).Picture
            Case "Module"
                .Picture = FrmResults.ilstImages.ListImages(9).Picture
            Case "Class"
                .Picture = FrmResults.ilstImages.ListImages(10).Picture
            Case "UserControl"
                .Picture = FrmResults.ilstImages.ListImages(11).Picture
            Case "UserDocument"
                .Picture = FrmResults.ilstImages.ListImages(12).Picture
            Case "PropetyPage"
                .Picture = FrmResults.ilstImages.ListImages(13).Picture
            Case "Designer"
                .Picture = FrmResults.ilstImages.ListImages(28).Picture
            Case "Clean"
                .Picture = FrmResults.ilstImages.ListImages(17).Picture
        End Select
    End With

    DoEvents ' Make sure the picture refreshes
End Sub

Private Sub AddProjectReportText() ' Add the stats collected about the current project to the report
    Dim StatHeadder As String, Temp As String, i As Long

    StatHeadder = "            Startup Item: " & Project.ProjectStartupItem
    StatHeadder = StatHeadder & vbNewLine & "             Source Safe: " & CheckIsSourceSafe(Project.ProjectPath) & vbNewLine

    StatHeadder = StatHeadder & vbNewLine & "            Project Type: " & Project.ProjectProjectType
    StatHeadder = StatHeadder & vbNewLine & "                   Lines: " & Project.ProjectLines
    StatHeadder = StatHeadder & vbNewLine & "       Lines (No Blanks): " & Project.ProjectLinesNB
    StatHeadder = StatHeadder & vbNewLine & "         Lines (Comment): " & Project.ProjectCommentLines
    StatHeadder = StatHeadder & vbNewLine & "      Declared Variables: " & Project.ProjectVariables
    StatHeadder = StatHeadder & vbNewLine & vbNewLine & "                   Forms: " & Project.ProjectForms

    StatHeadder = StatHeadder & vbNewLine & "                 Modules: " & Project.ProjectModules
    StatHeadder = StatHeadder & vbNewLine & "           Class Modules: " & Project.ProjectClasses
    StatHeadder = StatHeadder & vbNewLine & "           User Controls: " & Project.ProjectUserControls
    StatHeadder = StatHeadder & vbNewLine & "          User Documents: " & Project.ProjectUserDocuments
    StatHeadder = StatHeadder & vbNewLine & "          Property Pages: " & Project.ProjectPropertyPages
    StatHeadder = StatHeadder & vbNewLine & "                Designer: " & Project.ProjectDesigners

    StatHeadder = StatHeadder & vbNewLine & vbNewLine & "              --- Subs/Functions/Properties ---" & vbNewLine
    StatHeadder = StatHeadder & vbNewLine & "                    Subs: " & TotalSubs
    StatHeadder = StatHeadder & vbNewLine & "               Functions: " & TotalFunctions
    StatHeadder = StatHeadder & vbNewLine & "              Properties: " & TotalProperties
    StatHeadder = StatHeadder & vbNewLine & "                  Events: " & TotalEvents
    StatHeadder = StatHeadder & vbNewLine & "      Declared Ext. Subs: " & TotalDecSubs
    StatHeadder = StatHeadder & vbNewLine & " Declared Ext. Functions: " & TotalDecFunctions

    StatHeadder = StatHeadder & vbNewLine & vbNewLine & "                --- Components/References ---" & vbNewLine

    StatHeadder = StatHeadder & vbNewLine & "-----------------+------------------------------------------+"
    StatHeadder = StatHeadder & vbNewLine & " FileName:       | Name:                                    |"
    StatHeadder = StatHeadder & vbNewLine & "-----------------+------------------------------------------+"

    For i = 1 To Project.ProjectRefCom.Count Step 2 ' Step 2 to skip SPF description
        Temp = Left$(ExtractFileName(Project.ProjectRefCom.Item(i)), 16)   ' Gets the filename of the Component/Reference
        If Asc(Right$(Temp, 1)) = 0 Then Temp = Mid(Temp, 1, Len(Temp) - 1) ' Gets rid of unprintable characters that are sometimes at the end of the strings
        Temp = Temp & Space$(16 - Len(Temp)) ' Adds the spaces if nessesary to preserve the text table formatting
        StatHeadder = StatHeadder & vbCrLf & Temp & " | " ' Add the text table seperator
        Temp = Left$(Project.ProjectRefCom.Item(i + 1), 40) ' Get the Component/Reference name
        If Asc(Right$(Temp, 1)) = 0 Then Temp = Mid(Temp, 1, Len(Temp) - 1) ' Gets rid of unprintable characters that are sometimes at the end of the strings
        StatHeadder = StatHeadder & Temp & Space$(40 - Len(Temp)) & " |" ' Add the final text table seperator
    Next

    StatHeadder = StatHeadder & vbNewLine & "-----------------+------------------------------------------+"

    StatHeadder = StatHeadder & vbNewLine & vbNewLine & "                    --- Declared DLLs ---" & vbNewLine

    StatHeadder = StatHeadder & vbNewLine & "------------------------------------------------------------+"
    StatHeadder = StatHeadder & vbNewLine & " FileName:                                                  |"
    StatHeadder = StatHeadder & vbNewLine & "------------------------------------------------------------+"

    For i = 1 To Project.ProjectDecDlls.Count
        StatHeadder = StatHeadder & vbNewLine & Project.ProjectDecDlls.Item(i) & Space$(59 - Len(Project.ProjectDecDlls.Item(i))) & " |"
    Next

    If Project.ProjectDecDlls.Count = 0 Then StatHeadder = StatHeadder & vbNewLine & "(None)" & Space$(54) & "|"

    StatHeadder = StatHeadder & vbNewLine & "------------------------------------------------------------+"

    StatHeadder = StatHeadder & vbNewLine & vbNewLine & "                  --- Related Documents ---" & vbNewLine
    StatHeadder = StatHeadder & vbNewLine & "------------------------------------------------------------+"
    StatHeadder = StatHeadder & vbNewLine & " FileName:                                                  |"
    StatHeadder = StatHeadder & vbNewLine & "------------------------------------------------------------+"

    If RelatedDocumentFNames <> "" Then
        StatHeadder = StatHeadder & vbNewLine & RelatedDocumentFNames
        StatHeadder = StatHeadder & "------------------------------------------------------------+"
    Else
        StatHeadder = StatHeadder & vbNewLine & "(None)" & Space$(54) & "|"
        StatHeadder = StatHeadder & vbNewLine & "------------------------------------------------------------+"
    End If

    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectName", Project.ProjectName)
    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectVersion", Project.ProjectVersion)
    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectStats", StatHeadder)
End Sub

Private Sub AddReportText(ByVal AddText As String, Optional NoAddNL As Boolean) ' Adds a new line and the inputted text to the report
    If NoAddNL = False Then ' Add a new line at the start of the string
        FrmReport.rtbReportText.Text = FrmReport.rtbReportText.Text & vbNewLine & AddText
    Else ' Don't add new line
        FrmReport.rtbReportText.Text = FrmReport.rtbReportText.Text & AddText
    End If
End Sub

Private Function RemLineNumber(ByVal Linedata As String) As String ' Remove line numbers from code lines
    RemLineNumber = Linedata

    If Linedata = vbNullString Then Exit Function
    If Int(Left$(Linedata, 1)) = Left$(Linedata, 1) Then GoTo RemLineNum
    Exit Function

RemLineNum:                                                                                                                                                                                                                         ' Line number found
    If InStr(1, Linedata, " ") = 0 Then
        RemLineNumber = vbNullString
        Exit Function ' Blank line
    Else
        RemLineNumber = Mid$(Linedata, InStr(1, Linedata, " ") + 1) ' Get line data
    End If
End Function

Private Function CheckVer(File1Ver As Variant, File2Ver As Variant) As String ' Returns which file is newer from two version numbers
    Dim F1P1 As Integer, F1P2 As Integer, F1P3 As Integer
    Dim F2P1 As Integer, F2P2 As Integer, F2P3 As Integer

    On Error Resume Next

    F1P1 = Int(File1Ver(0)) ' Input is an array,
    F1P2 = Int(File1Ver(1)) ' so split the data
    F1P3 = Int(File1Ver(2)) ' into variables

    F2P1 = Int(File2Ver(0)) ' Input is an array,
    F2P2 = Int(File2Ver(1)) ' so split the data
    F2P3 = Int(File2Ver(2)) ' into variables

    If F1P1 < F2P1 Then CheckVer = "2N": Exit Function
    If F1P1 > F2P1 Then CheckVer = "1N": Exit Function

    If F1P2 < F2P2 Then CheckVer = "2N": Exit Function
    If F1P2 > F2P2 Then CheckVer = "1N": Exit Function

    If F1P3 < F2P3 Then CheckVer = "2N": Exit Function
    If F1P3 > F2P3 Then CheckVer = "1N": Exit Function

    CheckVer = "EQ"
    Exit Function
End Function

Private Function ZeroIfNull(Data As String) As String
    ZeroIfNull = Data
    If Data = vbNullString Then ZeroIfNull = "0"
End Function
