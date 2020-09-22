VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmResults 
   BackColor       =   &H00D5E6EA&
   Caption         =   "DeepLook Scan Results"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12240
   Icon            =   "Results.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TreeView 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9763
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ilstImages"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstUnusedVarImages 
      Left            =   2760
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":058A
            Key             =   "Name"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":069E
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":080A
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":0C16
            Key             =   "GlobalLocal"
         EndProperty
      EndProperty
   End
   Begin DeepLook.ucStatusBar sbrStatus 
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   6840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   582
   End
   Begin DeepLook.ucchameleonButton btnAbout 
      Height          =   375
      Left            =   11280
      TabIndex        =   13
      ToolTipText     =   "Show the About Dialog"
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "About"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "Results.frx":0D2A
      PICN            =   "Results.frx":0D46
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstBadCode 
      Height          =   450
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox pbxControlPanel 
      BackColor       =   &H00D5E6EA&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   8415
      TabIndex        =   2
      Top             =   6360
      Width           =   8415
      Begin DeepLook.ucchameleonButton btnFileCopy 
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         ToolTipText     =   "Copy all the Project's DLL and OCX Files to the \DLLOCX\ Directory."
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Copy Dependancies"
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
         MICON           =   "Results.frx":1134
         PICN            =   "Results.frx":1150
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucThreeDLineUD linSep2 
         Height          =   375
         Left            =   5640
         TabIndex        =   3
         Top             =   0
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   661
      End
      Begin DeepLook.ucThreeDLineUD linSep1 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   0
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   661
      End
      Begin DeepLook.ucchameleonButton btnExitDeepLook 
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         ToolTipText     =   "Exit DeepLook"
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "Results.frx":1535
         PICN            =   "Results.frx":1551
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucchameleonButton btnScanAnother 
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         ToolTipText     =   "Close this window and scan another VB Project"
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Scan Another"
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Results.frx":18E3
         PICN            =   "Results.frx":18FF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucchameleonButton btnGenerateReport 
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         ToolTipText     =   "Show a text report based on the project scan"
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Show Report"
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
         MICON           =   "Results.frx":1A5F
         PICN            =   "Results.frx":1A7B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucchameleonButton btnCollapseAll 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         ToolTipText     =   "Collapse all items"
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "- All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14018282
         BCOLO           =   14018282
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Results.frx":1E29
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin DeepLook.ucchameleonButton btnExpand 
         Height          =   375
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Expand Items"
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "+ Item(s)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14018282
         BCOLO           =   14018282
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Results.frx":1E45
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
   Begin MSComctlLib.ImageList ilstImages 
      Left            =   1440
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12583104
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":1E61
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":226D
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":2635
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":2B89
            Key             =   "SysDLL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":30DD
            Key             =   "Component"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":3631
            Key             =   "REFCOM"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":3B85
            Key             =   "RelatedDocuments"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":40D9
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":446D
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":4881
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":4C95
            Key             =   "UserControl"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":50E5
            Key             =   "UserDocument"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":5479
            Key             =   "PropertyPage"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":588D
            Key             =   "Method"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":5CA5
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":60B9
            Key             =   "BadCode"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":61CD
            Key             =   "Clean"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":6561
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":68B5
            Key             =   "SPF"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":6C51
            Key             =   "Resource"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":71F7
            Key             =   "RelDoc"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":7361
            Key             =   "NETproject"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":76B5
            Key             =   "NETvb"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":7A09
            Key             =   "Info"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":8A5D
            Key             =   "LOGFile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":8E31
            Key             =   "IComponent"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":9385
            Key             =   "App"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":98DD
            Key             =   "Designer"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":9CF5
            Key             =   "CodeLoop"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":9E09
            Key             =   "Total"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":A15D
            Key             =   "Event"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":A2B1
            Key             =   "CreateObject"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":A80D
            Key             =   "UnusedVar"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":AD61
            Key             =   "NoUnusedVar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Results.frx":B2B5
            Key             =   "Warning"
         EndProperty
      EndProperty
   End
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12870
      _ExtentX        =   21802
      _ExtentY        =   661
   End
   Begin MSComctlLib.ListView lstVarList 
      Height          =   5535
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ilstUnusedVarImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project File"
         Object.Width           =   4939
         ImageKey        =   "Project"
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   4939
         ImageKey        =   "File"
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Variable Name"
         Object.Width           =   4939
         ImageKey        =   "Name"
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   1852
         ImageKey        =   "GlobalLocal"
      EndProperty
   End
   Begin TabDlg.SSTab tabModeSelect 
      Height          =   420
      Left            =   1440
      TabIndex        =   15
      Top             =   420
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   741
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   2469
      BackColor       =   14018282
      TabCaption(0)   =   "Statistics"
      TabPicture(0)   =   "Results.frx":B609
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Unused Variables"
      TabPicture(1)   =   "Results.frx":B625
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.Label lblExternal 
      BackStyle       =   0  'Transparent
      Caption         =   "External"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4200
      TabIndex        =   26
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   3960
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblPrivate 
      BackStyle       =   0  'Transparent
      Caption         =   "Private"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   5040
      TabIndex        =   25
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   4800
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblPublic 
      BackStyle       =   0  'Transparent
      Caption         =   "Public"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   5880
      TabIndex        =   24
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   5640
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   6480
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   7320
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   8160
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   11520
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   9840
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   10680
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "Empty"
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
      Left            =   9240
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   9000
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblScanResults 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Results:"
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
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblNormal 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
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
      Left            =   6720
      TabIndex        =   23
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblFriend 
      BackStyle       =   0  'Transparent
      Caption         =   "Friend"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   7560
      TabIndex        =   22
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblStatic 
      BackStyle       =   0  'Transparent
      Caption         =   "Static"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   8400
      TabIndex        =   21
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblLet 
      BackStyle       =   0  'Transparent
      Caption         =   "Let"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   11760
      TabIndex        =   20
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   10920
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblGet 
      BackStyle       =   0  'Transparent
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   10080
      TabIndex        =   19
      Top             =   480
      Width           =   615
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show"
      Visible         =   0   'False
      Begin VB.Menu mnuShowSUBS 
         Caption         =   "Subs"
      End
      Begin VB.Menu mnuShowFUNCTIONS 
         Caption         =   "Functions"
      End
      Begin VB.Menu mnuShowPROPERTIES 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuShowEVENTS 
         Caption         =   "Events"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowFORMS 
         Caption         =   "Forms"
      End
      Begin VB.Menu mnuShowMODULES 
         Caption         =   "Modules"
      End
      Begin VB.Menu mnuShowCLASSES 
         Caption         =   "Classes"
      End
      Begin VB.Menu mnuShowUC 
         Caption         =   "User Controls"
      End
      Begin VB.Menu mnuShowUD 
         Caption         =   "User Documents"
      End
      Begin VB.Menu mnuShowPP 
         Caption         =   "Property Pages"
      End
      Begin VB.Menu mnuShowAllVB 
         Caption         =   "All VB Items"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowREFCOM 
         Caption         =   "References, Components && Declared DLLs"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowall 
         Caption         =   "All"
      End
   End
   Begin VB.Menu mnuNetshow 
      Caption         =   "NETshow"
      Visible         =   0   'False
      Begin VB.Menu mnuShowIMPORTS 
         Caption         =   "Imports"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowall2 
         Caption         =   "All"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ExOptns"
      Visible         =   0   'False
      Begin VB.Menu mnuHLGHTpmc 
         Caption         =   "Highlight Potentially Malicious Code"
      End
      Begin VB.Menu mnuHLGHTespfs 
         Caption         =   "Highlight Empty SPF's"
      End
      Begin VB.Menu mnuHLGHTexsf 
         Caption         =   "Highlight External SF's"
      End
   End
End
Attribute VB_Name = "FrmResults"
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

Private Sub btnAbout_Click()
    FrmAbout.Show
End Sub

' -----------------------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------------------

Private Sub btnCollapseAll_Click()
    Dim i As Long
    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    For i = 1 To TreeView.Nodes.Count
        TreeView.Nodes(i).Expanded = False
    Next

    TreeView.Nodes(1).Expanded = True

    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub btnExitDeepLook_Click()
    Dim X As Long

    On Error Resume Next
    For X = 1 To Me.Controls.Count
        Me.Controls(X).Enabled = False
    Next

    Me.sbrStatus.Text = "Clearing variable lists, please wait..."
    DoEvents

    IsExit = True
    TreeView.Enabled = False
    DoEvents
    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0
    btnCollapseAll_Click
    FrmResults.lstVarList.ListItems.Clear
    DoEvents

    Me.sbrStatus.Text = "Clearing treeview nodes, please wait..."
    DoEvents

    On Error Resume Next
    For X = 1 To TreeView.Nodes.Count
        TreeView.Nodes.Remove X ' Interestingly, this is MUCH faster than the .Nodes.Clear command
    Next

    Unload Me
End Sub

Private Sub btnExpand_Click()
    If PROJECTMODE = VB6 Then PopupMenu mnuShow
    If PROJECTMODE = NET Then PopupMenu mnuNetshow
End Sub

Private Sub btnFileCopy_Click()
    Me.MousePointer = 13
    btnFileCopy.Enabled = False
    ModFileSearchHandler.CopyDLLOCX
    btnFileCopy.Enabled = True
    Me.MousePointer = 1
End Sub

Private Sub btnGenerateReport_Click()
    FrmReport.Show
    FrmReport.Visible = True
End Sub

Private Sub btnScanAnother_Click()
    Dim X As Long
    X = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    Me.sbrStatus.Text = "Clearing treeview nodes, please wait..."

    DoEvents

    On Error Resume Next
    FrmSelProject.Show
    For X = 1 To Forms.Count
        If Forms(X).Name <> "FrmSelProject" Then Unload Forms(X) ' Kill everything except this form, so that only the SelProject form is shown once this closes
    Next

    FrmResults.lstVarList.ListItems.Clear

    Unload Me

    X = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub Form_Load()
    ModNoClose.DisableCloseButton Me, True

    IsExit = False
    SendMessage lstBadCode.hwnd, WM_SETREDRAW, 0, 0
    TreeView.HotTracking = True

    shpKeyColour(0).FillColor = RGB(150, 150, 150)
    shpKeyColour(1).FillColor = RGB(130, 0, 200)
    shpKeyColour(2).FillColor = RGB(200, 0, 150)
    shpKeyColour(3).FillColor = RGB(10, 150, 10)
    shpKeyColour(4).FillColor = RGB(20, 90, 100)
    shpKeyColour(5).FillColor = RGB(50, 23, 80)
    shpKeyColour(6).FillColor = RGB(255, 0, 0)
    shpKeyColour(7).FillColor = RGB(249, 164, 0)
    shpKeyColour(8).FillColor = RGB(217, 206, 19)
    shpKeyColour(9).FillColor = RGB(19, 217, 192)

    Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsExit = True Then KillProgram False
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    sbrStatus.Width = Me.Width - 125
    sbrStatus.Top = Me.ScaleHeight - 350

    pbxControlPanel.Top = Me.Height - pbxControlPanel.Height - 850
    pbxControlPanel.Left = (Me.Width / 2) - (pbxControlPanel.Width / 2)

    TreeView.Width = Me.Width - 350
    lstVarList.Width = TreeView.Width

    TreeView.Height = pbxControlPanel.Top - 850
    lstVarList.Height = TreeView.Height

    If Me.Height < 4000 Then Me.Height = 4000
    If Me.Width < 12360 Then Me.Width = 12360

    btnAbout.Left = Me.Width - btnAbout.Width - 130
    hedDeepLookHeader.ResizeMe
End Sub

Private Sub lstVarList_DblClick()
    If lstVarList.SelectedItem.Index = lstVarList.ListItems.Count Then
        MsgBoxEx "The DeepLook 4 Unused Variable Scanning engine does not currently" & vbNewLine & _
               "detect unused variables whose names are not unique in the same file." & vbNewLine & _
               vbNewLine & _
               "This means that an unused variable ""Bananna"" will only be detected if there" & vbNewLine & _
               "are no other variables named ""Bannana"" that are used in the same VB file.", vbInformation, "DeepLook Unused Variable Scanning Engine Note"
    End If
End Sub

Private Sub MNUHLGHTespfs_Click()
    Dim i As Long

    mnuHLGHTespfs.Checked = Not mnuHLGHTespfs.Checked

    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0

    If mnuHLGHTespfs.Checked Then
        For i = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(i).ForeColor = RGB(255, 0, 0) Then
                TreeView.Nodes(i).ForeColor = RGB(255, 255, 255)
                TreeView.Nodes(i).BackColor = RGB(255, 0, 0)
            End If
        Next
    Else
        For i = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(i).BackColor = RGB(255, 0, 0) Then
                TreeView.Nodes(i).BackColor = RGB(255, 255, 255)
                TreeView.Nodes(i).ForeColor = RGB(255, 0, 0)
            End If
        Next
    End If

    SendMessage TreeView.hwnd, WM_SETREDRAW, 1, 0
End Sub

Private Sub MNUHLGHTexsf_Click()
    Dim i As Long

    mnuHLGHTexsf.Checked = Not mnuHLGHTexsf.Checked

    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0

    If mnuHLGHTexsf.Checked Then
        For i = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(i).ForeColor = RGB(150, 150, 150) Then
                TreeView.Nodes(i).ForeColor = RGB(255, 255, 255)
                TreeView.Nodes(i).BackColor = RGB(200, 200, 200)
            End If
        Next
    Else
        For i = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(i).BackColor = RGB(150, 150, 150) Then
                TreeView.Nodes(i).BackColor = RGB(255, 255, 255)
                TreeView.Nodes(i).ForeColor = RGB(150, 150, 150)
            End If
        Next
    End If

    SendMessage TreeView.hwnd, WM_SETREDRAW, 1, 0
End Sub

Private Sub MNUHLGHTpmc_Click()
    Dim i As Long

    mnuHLGHTpmc.Checked = Not mnuHLGHTpmc.Checked

    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0

    If mnuHLGHTpmc.Checked Then
        For i = 1 To TreeView.Nodes.Count
            If InStrB(1, TreeView.Nodes(i).Key, "_PMC_") <> 0 Then
                If Right$(TreeView.Nodes(i).Key, 5) <> "COUNT" Then
                    TreeView.Nodes(i).ForeColor = RGB(255, 255, 255)
                    TreeView.Nodes(i).BackColor = RGB(0, 180, 0)
                End If
            End If
        Next
    Else
        For i = 1 To TreeView.Nodes.Count
            If InStrB(1, TreeView.Nodes(i).Key, "_PMC_") <> 0 Then
                If Right$(TreeView.Nodes(i).Key, 5) <> "COUNT" Then
                    TreeView.Nodes(i).BackColor = RGB(255, 255, 255)
                    TreeView.Nodes(i).ForeColor = RGB(0, 0, 0)
                End If
            End If
        Next
    End If

    SendMessage TreeView.hwnd, WM_SETREDRAW, 1, 0
End Sub

Private Sub MNUshowAll_Click()
    Dim i As Long
    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    For i = 1 To TreeView.Nodes.Count
        TreeView.Nodes(i).Expanded = True
    Next

    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub MNUshowall2_Click()
    Dim i As Long

    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    For i = 1 To TreeView.Nodes.Count
        TreeView.Nodes.Item(i).Expanded = True
    Next i

    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub MNUshowAllVB_Click()
    ExpandKeyByPic "Form"
    ExpandKeyByPic "Class"
    ExpandKeyByPic "Module"
    ExpandKeyByPic "PropertyPage"
    ExpandKeyByPic "UserControl"
    ExpandKeyByPic "UserDocument"
End Sub

Private Sub MNUshowCLASSES_Click()
    ExpandKeyByPic "Class"
End Sub

Private Sub MNUshowEVENTS_Click()
    ExpandKey "_EVENTS"
End Sub

Private Sub MNUshowFORMS_Click()
    ExpandKeyByPic "Form"
End Sub

Private Sub MNUshowFUNCTIONS_Click()
    ExpandKey "_FUNCTIONS"
End Sub

Sub ExpandKey(Suffix As String)
    On Error Resume Next
    Dim i As Long, X As Long, z As Long
    X = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    For i = 1 To TreeView.Nodes.Count
        If Right$(TreeView.Nodes(i).Key, Len(Suffix)) = Suffix Then
            TreeView.Nodes(i).Expanded = True
            TreeView.Nodes(i).Parent.Expanded = True
            z = TreeView.Nodes(i).Parent.Index
            TreeView.Nodes(z).Parent.Expanded = True
            TreeView.Nodes(TreeView.Nodes(z).Parent.Index).Parent.Expanded = True
        End If

        If TreeView.Nodes(i).Key = "GROUP" Then TreeView.Nodes(i).Expanded = True
    Next

    X = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Sub ExpandKeyByPic(ImageKey As String, Optional DoubleParent As Boolean)
    Dim i As Long, z As Long
    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    On Error Resume Next

    For i = 1 To TreeView.Nodes.Count
        If TreeView.Nodes(i).Image = ImageKey Then
            TreeView.Nodes(i).Expanded = True
            TreeView.Nodes(i).Parent.Expanded = True

            If DoubleParent = True Then
                z = TreeView.Nodes(i).Parent.Index
                TreeView.Nodes(z).Parent.Expanded = True
            End If
        End If

        If TreeView.Nodes(i).Key = "GROUP" Then TreeView.Nodes(i).Expanded = True
    Next

    i = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub MNUshowIMPORTS_Click()
    ExpandKeyByPic "SysDLL"
    ExpandKey "NETfiles"
End Sub

Private Sub MNUshowMODULES_Click()
    ExpandKeyByPic "Module"
End Sub

Private Sub MNUshowPP_Click()
    ExpandKeyByPic "PropertyPage"
End Sub

Private Sub MNUshowPROPERTIES_Click()
    ExpandKey "_PROPERTIES"
End Sub

Private Sub MNUshowREFCOM_Click()
    ExpandKeyByPic "DLL", True
    ExpandKeyByPic "Component", True
End Sub

Private Sub MNUshowSUBS_Click()
    ExpandKey "_SUBS"
End Sub

Private Sub MNUshowUC_Click()
    ExpandKeyByPic "UserControl"
End Sub

Private Sub MNUshowUD_Click()
    ExpandKeyByPic "UserDocument"
End Sub

Private Sub tabModeSelect_Click(PreviousTab As Integer)
    If tabModeSelect.Tab = 0 Then
        TreeView.Visible = True
        lstVarList.Visible = False
    Else
        TreeView.Visible = False
        lstVarList.Visible = True
    End If
End Sub

Private Sub TreeView_Click()
    Dim CurrKey As String

    CurrKey = Replace$(Mid$(TreeView.SelectedItem.FullPath, 1, InStrRev(TreeView.SelectedItem.FullPath, "\")), "\ ", "\")
    CurrKey = Replace$(CurrKey, "&", "&&", 1)

    If CurrKey = "" Then
        sbrStatus.Text = DefaultStatText
    Else
        If ShowItemKey = True Then
            sbrStatus.Text = TreeView.SelectedItem.Key & " (" & TreeView.SelectedItem.Index & ") Node Total: " & TreeView.Nodes.Count
        Else
            sbrStatus.Text = "Current Node Location: " & CurrKey
        End If
    End If
End Sub

Private Sub TreeView_DblClick()
    With TreeView
        If InStr(1, .SelectedItem.Text, "(Double-Click", vbTextCompare) <> 0 Then
            Shell "Notepad.exe " & Mid$(.SelectedItem.Key, InStrRev(.SelectedItem.Key, "_") + 1), vbNormalFocus
        End If
    End With
End Sub

Private Sub TreeView_KeyPress(KeyAscii As Integer)
    TreeView_Click
End Sub

Private Sub TreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 1 Then PopupMenu mnuPopup
End Sub

Private Sub TreeView_NodeCheck(ByVal Node As MSComctlLib.Node)
    TreeView_Click
End Sub

Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    TreeView_Click
End Sub

Private Sub TreeView_Validate(Cancel As Boolean)
    TreeView_Click
End Sub
