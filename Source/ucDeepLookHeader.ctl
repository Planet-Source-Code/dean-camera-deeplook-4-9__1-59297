VERSION 5.00
Begin VB.UserControl ucDeepLookHeader 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   ScaleHeight     =   810
   ScaleWidth      =   5355
   ToolboxBitmap   =   "ucDeepLookHeader.ctx":0000
   Begin VB.PictureBox DLLogo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   370
      Left            =   -1
      Picture         =   "ucDeepLookHeader.ctx":0312
      ScaleHeight     =   375
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "ucDeepLookHeader"
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

' Used to minimize memory requirements for DeepLook by only storing one logo

Private Sub UserControl_Resize()
On Error Resume Next

    DLLogo.Left = 0
    DLLogo.Top = 0
    DLLogo.Width = UserControl.Width

    UserControl.Height = DLLogo.Height
    UserControl.Width = UserControl.Parent.Width
End Sub

Sub ResizeMe()
    UserControl_Resize
End Sub
