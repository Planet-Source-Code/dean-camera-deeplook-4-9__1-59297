VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2029
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"File"
Attribute VB_Ext_KEY = "Member0" ,"File"
'               THIS CLASS WAS NOT WRITTEN BY EN-TECH.
'             I CANNOT OFFER ANY SUPPORT FOR THIS CLASS.

Option Explicit

Private mCol As Collection

Public Function Add(ByVal FileNameFull As String, Optional sKey As String) As ClsFile
Attribute Add.VB_HelpID = 2030
    'create a new object
    Dim objNewMember As ClsFile
    Set objNewMember = New ClsFile

    'set the properties passed into the method
    objNewMember.FileNameFull = FileNameFull
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(ByVal vntIndexKey As Variant) As ClsFile
Attribute Item.VB_HelpID = 2031
Attribute Item.VB_UserMemId = 0
On Error Resume Next
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
Attribute Count.VB_HelpID = 2032
    Count = mCol.Count
End Property

Public Sub Remove(ByVal vntIndexKey As Variant)
Attribute Remove.VB_HelpID = 2033
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_HelpID = 2034
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

