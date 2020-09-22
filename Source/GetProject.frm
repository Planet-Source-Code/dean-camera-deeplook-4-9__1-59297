VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSelProject 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DeepLook Project Scanner"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "GetProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GetProject.frx":23D2
   ScaleHeight     =   2130
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProjectPath 
      DragIcon        =   "GetProject.frx":2714
      Height          =   285
      Left            =   840
      OLEDropMode     =   2  'Automatic
      TabIndex        =   1
      Top             =   1110
      Width           =   4215
   End
   Begin DeepLook.ucProgressBar pgbAPB 
      Height          =   135
      Left            =   840
      TabIndex        =   8
      Top             =   1150
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   238
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   49152
      Scrolling       =   9
   End
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6210
      _ExtentX        =   11060
      _ExtentY        =   661
   End
   Begin DeepLook.ucchameleonButton btnBrowseButton 
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Browse for a VB project"
      Top             =   1125
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   4
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "GetProject.frx":2DFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DeepLook.ucchameleonButton btnGoAnalise 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Analyse the selected VB project"
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Analyse"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "GetProject.frx":2E1A
      PICN            =   "GetProject.frx":2E36
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdgCommonDialog 
      Left            =   5280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DeepLook.ucchameleonButton btnExitDeepLook 
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      ToolTipText     =   "Exit DeepLook"
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "GetProject.frx":2F96
      PICN            =   "GetProject.frx":2FB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DeepLook.ucchameleonButton btnOptions 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Set Options"
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Options"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "GetProject.frx":3344
      PICN            =   "GetProject.frx":3360
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DeepLook.ucProgressBar pgbGroupProgress 
      Height          =   135
      Left            =   840
      TabIndex        =   9
      Top             =   990
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   238
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16711680
      Scrolling       =   9
   End
   Begin VB.Image imgScanFileIcon 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "GetProject.frx":36E9
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblScanPhase 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1500
      Width           =   4215
   End
   Begin VB.Image imgCurrScanObjType 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5160
      Picture         =   "GetProject.frx":3FB3
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblScanningName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblLocateProjText 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the filename of the Visual Basic file you wish to be scanned."
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
End
Attribute VB_Name = "FrmSelProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim ProjectExt As String
Dim Unloading As Boolean

' -----------------------------------------------------------------------------------------------

Private Sub btnBrowseButton_Click()
    cdgCommonDialog.Filter = "VB6 and .NET Projects (*.vbp, *.vbg, *.vbproj)|*.vbp;*.vbg;*.vbproj|Visual Basic 6 Project Files (*.vbp *.vbg)|*.vbp;*.vbg|Single Visual Basic 6 File|*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob;*.dsr|.NET Project (*.vbproj)|*.vbproj"
    cdgCommonDialog.ShowOpen
    txtProjectPath.Text = cdgCommonDialog.FileName
End Sub

Private Sub btnExitDeepLook_Click()
    Me.Visible = False
    IsExit = True
    Unload Me
End Sub

Private Sub Form_Load()
    ModNoClose.DisableCloseButton Me, True

    Load FrmOptions

    Unloading = False
    txtProjectPath.Visible = True
    pgbAPB.Visible = False
    lblScanningName.Visible = False
    btnGoAnalise.Enabled = True
    btnBrowseButton.Visible = True
    pgbGroupProgress.Visible = False
    imgCurrScanObjType.Visible = False
    Me.Caption = "DeepLook Project Scanner"
    lblLocateProjText.Caption = "Please enter the filename of the Visual Basic file you wish to be scanned."
    lblScanPhase.Caption = ""
    btnOptions.Enabled = True

    Screen.MousePointer = 0

    btnGoAnalise.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    KillProgram Unloading
End Sub

Private Sub btnGoAnalise_Click()
    Dim StartTime As Long, TimeTemp As String, AddS As String, LinesPerSec As Long, LPSData As String, i As Long

    TotalLines = 0

    ProjectPath = txtProjectPath.Text

    Me.Caption = "DeepLook Project Scanner - Scanning..."

    btnOptions.Enabled = False
    btnGoAnalise.Enabled = False
    btnBrowseButton.Visible = False
    txtProjectPath.Visible = False
    pgbGroupProgress.Value = 0
    lblScanningName.Visible = True
    imgCurrScanObjType.Visible = True
    pgbAPB.Visible = True
    imgCurrScanObjType.Top = 1080

    Screen.MousePointer = 11

    FrmReport.LockRefresh True
    ModAnalyseVB6.ClearCollections
    ProjectExt = Right$(UCase(txtProjectPath.Text), 4)

    StartTime = Timer

    AddReportHeadder

    SendMessage FrmResults.TreeView.hwnd, WM_SETREDRAW, 0, 0

    FrmResults.TreeView.Visible = False
    FrmResults.tabModeSelect.Visible = True
    FrmResults.tabModeSelect.Enabled = True

    Select Case ProjectExt
        Case ".VBG"
            lblLocateProjText.Caption = "Scanning Group:"
            pgbGroupProgress.Visible = True
            imgCurrScanObjType.Top = 1000
            ModAnalyseVB6.AnalyseGroup txtProjectPath.Text
        Case ".VBP"
            lblLocateProjText.Caption = "Scanning Project:"
            ModAnalyseVB6.AnalyseVBProject txtProjectPath.Text
        Case ".FRM", ".BAS", ".CLS", ".CTL", ".PAG", ".DOB", ".DSR"
            lblLocateProjText.Caption = "Scanning File:"
            ModAnalyseVB6.AnalyseSingleVBItem txtProjectPath.Text
        Case "PROJ"
            lblLocateProjText.Caption = "Scanning Project:"
            ModAnalyseDOTNET.AnalyseDotNetProject txtProjectPath.Text
        Case Else
            MsgBoxEx "Not a Visual Basic file!", vbCritical, "DeepLook - Scan Error", , , , , PicError, "Oops!|"
            Form_Load
            Exit Sub
    End Select

    SendMessage FrmResults.TreeView.hwnd, WM_SETREDRAW, 1, 0

    AddReportFooter

    FrmReport.LockRefresh False
    Screen.MousePointer = 0

    pgbGroupProgress.Value = 100

    TimeTemp = Round(Timer - StartTime, 2)
    If TimeTemp = "1" Then AddS = "" Else AddS = "s"

    If Int(TimeTemp) < 1 Then LinesPerSec = 1 Else LinesPerSec = Int(TimeTemp)

    LinesPerSec = Round(TotalLines / LinesPerSec, 2)
    If LinesPerSec < 1 Then LPSData = "" Else LPSData = " (" & LinesPerSec & " lines/sec)"

    DefaultStatText = "Scan took " & TimeTemp & " second" & AddS & LPSData & ". This is DeepLook version " & App.Major & "." & App.Minor & "." & App.Revision & "."
    FrmResults.sbrStatus.Text = DefaultStatText

    If PROJECTMODE = VB6 Then
        FrmResults.btnFileCopy.Enabled = True
        FrmResults.btnGenerateReport.Enabled = True
    Else
        FrmResults.btnFileCopy.Enabled = False
        FrmResults.btnGenerateReport.Enabled = False
    End If

    If FrmResults.TreeView.Nodes.Count = 1 Then  ' If only one node in results treeview (only project or group file not found) then don't show results
        Unload Me   ' Unload current select project window
        Load Me     ' Load a new select project window
        Me.Show     ' Show the new select project window
        Exit Sub    ' Don't continue processing the scanned data
    End If

    FrmResults.TreeView.Visible = True
    FrmResults.Show
    Unloading = True
    Unload Me

    With FrmResults.lstVarList
        i = .ListItems.Count

        .ListItems.Add , , ""
        .ListItems.Add , , "Total: " & i
        .ListItems(.ListItems.Count).Bold = True

        .ListItems.Add , , ""
        .ListItems.Add , , ""
        .ListItems.Add , , "(Variable Scanner Note)"
        .ListItems(.ListItems.Count).ForeColor = RGB(0, 0, 150)

        If i = 0 Then
            FrmResults.TreeView.Nodes.Add 1, tvwChild, "UnusedVarDetails", "No unused variables were found in the scanned project(s).", "NoUnusedVar"
        Else
            FrmResults.TreeView.Nodes.Add 1, tvwChild, "UnusedVarDetails", "Unused variables were found in the scanned project(s).", "UnusedVar"
        End If
    End With

    If FrmResults.TreeView.Nodes.Count > 16000 Then
        FrmResults.TreeView.Nodes.Add 1, tvwChild, "NodeWarning", "Warning: There are a large amount (" & Format$(FrmResults.TreeView.Nodes.Count, "###,###,###") & ") of nodes in the treeview.", "Warning"
        FrmResults.TreeView.Nodes.Add 1, tvwChild, "NodeWarningP2", "This large number may cause DeepLook or your computer to run slowly or stop responding."
    End If
End Sub

Private Sub btnOptions_Click()
    FrmOptions.Show 1
End Sub

Private Sub txtProjectPath_Change()
    btnGoAnalise.Enabled = Len(txtProjectPath.Text)
End Sub
