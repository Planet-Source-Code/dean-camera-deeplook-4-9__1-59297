VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCopyReport 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DeepLook Copy Required Files Report - Copying..."
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvwNonCopyItemsTV 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
      _Version        =   393217
      Style           =   7
      ImageList       =   "ilstCopyImages"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilstCopyImages 
      Left            =   2640
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0000
            Key             =   "Component"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0554
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0AA8
            Key             =   "SysDLL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0FFC
            Key             =   "CreateObject"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":1558
            Key             =   "Done"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":1AAC
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":2000
            Key             =   "CurrentCopy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwItemsTV 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393217
      Style           =   7
      ImageList       =   "ilstCopyImages"
      Appearance      =   1
   End
   Begin DeepLook.ucchameleonButton btnCloseButton 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Close this window"
      Top             =   6240
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
      MICON           =   "FrmCopyReport.frx":2354
      PICN            =   "FrmCopyReport.frx":2370
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TreeView tvwManualCopyTV 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
      _Version        =   393217
      Style           =   7
      ImageList       =   "ilstCopyImages"
      Appearance      =   1
   End
   Begin DeepLook.ucProgressBar pgbPercentBar 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   65280
      Scrolling       =   9
      ShowText        =   -1  'True
   End
   Begin VB.Label lblManualCopy 
      BackStyle       =   0  'Transparent
      Caption         =   "The following files may need to be copied manually:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label lblUnessesaryCopyFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "The following files do not need to be copied:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lblCopyDir 
      BackStyle       =   0  'Transparent
      Caption         =   "(Dir)"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Label lblCopyDirLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Directory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label lblFilesBeingCopied 
      BackStyle       =   0  'Transparent
      Caption         =   "The following files are being copied:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "FrmCopyReport"
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

Private Sub btnCloseButton_Click()
Unload Me
End Sub

