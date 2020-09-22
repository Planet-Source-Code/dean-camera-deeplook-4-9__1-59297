VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About DeepLook"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5220
      _ExtentX        =   9049
      _ExtentY        =   661
   End
   Begin VB.Frame fmeAbout 
      BackColor       =   &H00D5E6EA&
      Caption         =   "About DeepLook"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      Begin DeepLook.ucThreeDLineLR linSep6 
         Height          =   90
         Left            =   120
         TabIndex        =   19
         Top             =   4680
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep5 
         Height          =   90
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep2 
         Height          =   90
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep4 
         Height          =   90
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep1 
         Height          =   90
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep3 
         Height          =   90
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep8 
         Height          =   90
         Left            =   120
         TabIndex        =   21
         Top             =   5760
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin DeepLook.ucThreeDLineLR linSep7 
         Height          =   90
         Left            =   120
         TabIndex        =   24
         Top             =   5400
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   159
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Charmelion Button Control by Gonchuki used under special licence."
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
         Left            =   120
         TabIndex        =   23
         Top             =   5520
         Width           =   4575
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Email me at dean_camera@hotmail.com."
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
         Left            =   120
         TabIndex        =   22
         Top             =   5880
         Width           =   4575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":06EA
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
         TabIndex        =   20
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label lblAuthorOf 
         BackStyle       =   0  'Transparent
         Caption         =   "Author Of:"
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
         Left            =   3360
         TabIndex        =   17
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblSpecialThanks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":07BC
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4275
         Width           =   4575
      End
      Begin VB.Image imgAuthor 
         Height          =   1140
         Left            =   120
         Picture         =   "About.frx":0844
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeepLook Project Scanner"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   11
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00D5E6EA&
         Caption         =   "By Dean Camera, 2003-2005"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "www.en-tech.i8.com."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2205
         TabIndex        =   8
         Top             =   2475
         Width           =   2175
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is DeepLook version #.#.#."
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
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:  Nationality:   Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D5E6EA&
         BackStyle       =   0  'Transparent
         Caption         =   "Dean Camera   Australian      16"
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Please feel free to visit the En-Tech Website for        FREE software at"
         Height          =   615
         Left            =   600
         TabIndex        =   9
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Label ControlCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":3699
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   15
         Top             =   2865
         Width           =   4455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeepLook, etURP, ComTalk"
         Height          =   495
         Left            =   3120
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
   End
   Begin DeepLook.ucchameleonButton CloseButton 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      ToolTipText     =   "Close this window"
      Top             =   6720
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
      MICON           =   "About.frx":37D2
      PICN            =   "About.frx":37EE
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
Attribute VB_Name = "FrmAbout"
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

Private Sub CloseButton_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "This is DeepLook version " & App.Major & "." & App.Minor & "." & App.Revision & "."
End Sub

