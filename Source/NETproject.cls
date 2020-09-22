VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNETproject"
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

Option Explicit

' -----------------------------------------------------------------------------------------------

Dim Code_LinesNB As Long
Dim Code_Lines As Long
Dim Blank_Lines As Long
Dim Comment_Lines As Long
Dim Total_Lines As Long

Property Get TotalLines() As Long
    TotalLines = Total_Lines
End Property

Property Let TotalLines(AddAmount As Long)
    Total_Lines = Total_Lines + AddAmount
End Property

Property Get CodeLinesNB() As Long
    CodeLinesNB = Code_LinesNB
End Property

Property Let CodeLinesNB(AddAmount As Long)
    Code_LinesNB = Code_LinesNB + AddAmount
End Property

Property Get CodeLines() As Long
    CodeLines = Code_Lines
End Property

Property Let CodeLines(AddAmount As Long)
    Code_Lines = Code_Lines + AddAmount
End Property

Property Get BlankLines() As Long
    BlankLines = Blank_Lines
End Property

Property Let BlankLines(AddAmount As Long)
    Blank_Lines = Blank_Lines + AddAmount
End Property

Property Get CommentLines() As Long
    CommentLines = Comment_Lines
End Property

Property Let CommentLines(AddAmount As Long)
    Comment_Lines = Comment_Lines + AddAmount
End Property

