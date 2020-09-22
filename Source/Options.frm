VERSION 5.00
Begin VB.Form FrmOptions 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   3330
   StartUpPosition =   2  'CenterScreen
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   661
   End
   Begin DeepLook.ucchameleonButton btnSaveSettings 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Options.frx":058A
      PICN            =   "Options.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fmeOptions 
      BackColor       =   &H00D5E6EA&
      Caption         =   "DeepLook Options"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      Begin VB.CheckBox chkFNFE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D5E6EA&
         Caption         =   "Show ""File Not Found"" Errors"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CheckBox chkSCAV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D5E6EA&
         Caption         =   "Scan Constants as Variables"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox chkSSPFLines 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D5E6EA&
         Caption         =   "Show individual Sub/Function/Property Lines"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   2895
      End
      Begin DeepLook.ucThreeDLineLR linSep3 
         Height          =   90
         Left            =   120
         TabIndex        =   13
         Top             =   4080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   159
      End
      Begin DeepLook.ucchameleonButton btnFast 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BTYPE           =   1
         TX              =   "Fast"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Options.frx":0974
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkSGROCSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D5E6EA&
         Caption         =   "Show Graphical Representation of the type of Object being scanned"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkSSFPPARA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D5E6EA&
         Caption         =   "Show Subs/Functions/Properties Parameters"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkCheckMaliciousCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D5E6EA&
         Caption         =   "Check for Potentially Malicious Code"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin DeepLook.ucchameleonButton btnThorough 
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   3360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   1
         TX              =   "Thorough"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Options.frx":0990
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucchameleonButton btnBestLooking 
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   1
         TX              =   "Best Looking"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Options.frx":09AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucchameleonButton btnThoroughBestLooking 
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   3720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BTYPE           =   1
         TX              =   "Thorough && Best Looking"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Options.frx":09C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucThreeDLineLR linSep1 
         Height          =   90
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep2 
         Height          =   90
         Left            =   2280
         TabIndex        =   12
         Top             =   3240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   159
      End
      Begin VB.Label lblSettingsTip 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tip: Some extra settings can be accessed by right-clicking the treeview after a project has been fully scanned by DeepLook."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   4125
         Width           =   2895
      End
      Begin VB.Label lblQuickSettings 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Settings:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   2895
      End
   End
   Begin DeepLook.ucchameleonButton btnCancel 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Options.frx":09E4
      PICN            =   "Options.frx":0A00
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
Attribute VB_Name = "FrmOptions"
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

Private Sub btnSaveSettings_Click()
    SaveSetting "DeepLook", "Options", "PMCCheck", chkCheckMaliciousCode.Value
    SaveSetting "DeepLook", "Options", "ShowSPFParams", chkSSFPPARA.Value
    SaveSetting "DeepLook", "Options", "ShowCurrItemPic", chkSGROCSO.Value
    SaveSetting "DeepLook", "Options", "ShowSPFLines", chkSSPFLines.Value
    SaveSetting "DeepLook", "Options", "ScanConstAsVar", chkSCAV.Value
    SaveSetting "DeepLook", "Options", "ShowFNFErrors", chkFNFE.Value

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnBestLooking_Click()
    chkCheckMaliciousCode.Value = 0
    chkSSFPPARA.Value = 0
    chkSGROCSO.Value = 1
    chkSSPFLines.Value = 0
    chkSCAV.Value = 1
End Sub

Private Sub btnFast_Click()
    chkCheckMaliciousCode.Value = 0
    chkSSFPPARA.Value = 1
    chkSGROCSO.Value = 0
    chkSSPFLines.Value = 0
    chkSCAV.Value = 0
End Sub

Private Sub Form_Load()
    chkCheckMaliciousCode.Value = GetSetting("DeepLook", "Options", "PMCCheck", 1)
    chkSSFPPARA.Value = GetSetting("DeepLook", "Options", "ShowSPFParams", 1)
    chkSGROCSO.Value = GetSetting("DeepLook", "Options", "ShowCurrItemPic", 1)
    chkSSPFLines.Value = GetSetting("DeepLook", "Options", "ShowSPFLines", 0)
    chkSCAV.Value = GetSetting("DeepLook", "Options", "ScanConstAsVar", 0)
    chkFNFE.Value = GetSetting("DeepLook", "Options", "ShowFNFErrors", 1)
End Sub

Private Sub btnThorough_Click()
    chkCheckMaliciousCode.Value = 1
    chkSSFPPARA.Value = 1
    chkSGROCSO.Value = 0
    chkSSPFLines.Value = 1
    chkSCAV.Value = 1
End Sub

Private Sub btnThoroughBestLooking_Click()
    chkCheckMaliciousCode.Value = 1
    chkSSFPPARA.Value = 0
    chkSGROCSO.Value = 1
    chkSSPFLines.Value = 1
    chkSCAV.Value = 1
End Sub
