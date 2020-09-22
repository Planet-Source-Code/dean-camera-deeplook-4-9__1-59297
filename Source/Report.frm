VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmReport 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DeepLook Project Report"
   ClientHeight    =   7260
   ClientLeft      =   4200
   ClientTop       =   2550
   ClientWidth     =   7080
   Icon            =   "Report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin DeepLook.ucDeepLookHeader ucDeepLookHeadder1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog cdgCommonDialgue 
      Left            =   2880
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbReportText 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11033
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Report.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DeepLook.ucchameleonButton btnCloseButton 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      ToolTipText     =   "Close this window"
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
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
      MICON           =   "Report.frx":0652
      PICN            =   "Report.frx":066E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DeepLook.ucchameleonButton btnSaveReport 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Close this window"
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save as Text File"
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
      MICON           =   "Report.frx":0A00
      PICN            =   "Report.frx":0A1C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmReport"
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

Sub LockRefresh(YesNo As Boolean)
    If YesNo = True Then
        SendMessage rtbReportText.hwnd, WM_SETREDRAW, 0, 0
    Else
        SendMessage rtbReportText.hwnd, WM_SETREDRAW, 1, 0
    End If
End Sub

Private Sub btnCloseButton_Click()
    Me.Visible = False
End Sub

Private Sub btnSaveReport_Click()
    cdgCommonDialgue.Filter = "DeepLook Report File (*.txt)|*.txt"
    cdgCommonDialgue.ShowSave
    rtbReportText.SaveFile cdgCommonDialgue.FileName, rtfText
End Sub

