VERSION 5.00
Begin VB.UserControl ucStatusBar 
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   ControlContainer=   -1  'True
   ScaleHeight     =   765
   ScaleWidth      =   2250
   ToolboxBitmap   =   "UCStatBarCtrl.ctx":0000
   Begin VB.Label TBarText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StatText"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   75
      Width           =   600
   End
   Begin VB.Image TBarRight 
      Height          =   330
      Left            =   2160
      Picture         =   "UCStatBarCtrl.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   75
   End
   Begin VB.Image TBarLeft 
      Height          =   330
      Left            =   0
      Picture         =   "UCStatBarCtrl.ctx":0534
      Stretch         =   -1  'True
      Top             =   0
      Width           =   75
   End
   Begin VB.Image TBarPicMiddle 
      Height          =   330
      Left            =   0
      Picture         =   "UCStatBarCtrl.ctx":0756
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ucStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  .======================================.
' /         DeepLook Project Scanner       \
' |           By Dean Camera, 2004         |
' \   Completly re-written from scratch    /
'  '======================================'
' /  For more FREE software, please visit  \
' |         the En-Tech Website at:        |
' \            www.en-tech.i8.com          /
'  '======================================'
' / Most of this project is now commented  \
' \           to help developers.          /
'  '======================================'

' Made by Dean Camera - Simple graphical statusbar

Private Sub UserControl_Resize()
TBarPicMiddle.Width = UserControl.Width
TBarText.Left = 40
TBarRight.Left = UserControl.Width - TBarRight.Width
UserControl.Height = TBarPicMiddle.Height
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Text = TBarText.Caption
End Property

Public Property Let Text(ByVal New_Text As String)
    TBarText.Caption() = New_Text
    PropertyChanged "Text"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    TBarText.Caption = PropBag.ReadProperty("Text", "TBLText")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", TBarText.Caption, "TBLText")
End Sub

