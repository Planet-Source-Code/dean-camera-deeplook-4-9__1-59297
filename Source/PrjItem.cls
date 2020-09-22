VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrjItem"
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

Private ItemName As String
Private ItemControls As Long

Private ItemInControls As Boolean
Private ItemSeenAttributes As Boolean
Private ItemLastLine As String

Private ItemCodeLines As Long
Private ItemCodeLinesNoBlanks As Long
Private ItemCommentLines As Long
Private ItemHybridLines As Long

Private ItemSubs As Collection
Private ItemFunctions As Collection
Private ItemProperties As Collection
Private ItemEvents As Collection

Private ItemVariables As Long
Private ItemConstants As Long
Private ItemTypes As Long
Private ItemEnums As Long

Private ItemStatementDO As Long
Private ItemStatementWHILE As Long
Private ItemStatementIF As Long
Private ItemStatementFOR As Long
Private ItemStatementSELECT As Long

Public Enum StatementType
    STDO = 1
    STWHILE = 2
    STIF = 3
    STFOR = 4
    STSELECT = 5
End Enum

Public Enum SPFTypes
    SPF_Sub = 0
    SPF_Function = 1
    SPF_Property = 2
    SPF_Event = 3
End Enum

'------------------------------------------------------------------------------

Property Get PrjItemName() As String
    PrjItemName = ItemName
End Property

Property Let PrjItemName(NewName As String)
    ItemName = NewName
End Property

'------------------------------------------------------------------------------

Property Get PrjItemControls() As Long
    PrjItemControls = ItemControls
End Property

Property Let PrjItemControls(AddAmount As Long)
    ItemControls = ItemControls + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get PrjItemSeenAttributes() As Boolean
    PrjItemSeenAttributes = ItemSeenAttributes
End Property

Property Let PrjItemSeenAttributes(YesNo As Boolean)
    ItemSeenAttributes = YesNo
End Property

'------------------------------------------------------------------------------

Property Get PrjItemPreviousLine() As String
    PrjItemPreviousLine = ItemLastLine
End Property

Property Let PrjItemPreviousLine(LastLine As String)
    ItemLastLine = LastLine
End Property

'------------------------------------------------------------------------------

Property Get PrjItemInControls() As Boolean
    PrjItemInControls = ItemInControls
End Property

Property Let PrjItemInControls(YesNo As Boolean)
    ItemInControls = YesNo
End Property

'------------------------------------------------------------------------------

Property Get PrjItemCodeLines() As Long
    PrjItemCodeLines = ItemCodeLines
End Property

Property Let PrjItemCodeLines(AddAmount As Long)
    ItemCodeLines = ItemCodeLines + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get PrjItemCodeLinesNoBlanks() As Long
    PrjItemCodeLinesNoBlanks = ItemCodeLinesNoBlanks
End Property

Property Let PrjItemCodeLinesNoBlanks(AddAmount As Long)
    ItemCodeLinesNoBlanks = ItemCodeLinesNoBlanks + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get PrjItemCommentLines() As Long
    PrjItemCommentLines = ItemCommentLines
End Property

Property Let PrjItemCommentLines(AddAmount As Long)
    ItemCommentLines = ItemCommentLines + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get PrjItemHybridLines() As Long
    PrjItemHybridLines = ItemHybridLines
End Property

Property Let PrjItemHybridLines(AddAmount As Long)
    ItemHybridLines = ItemHybridLines + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get PrjItemItemSubs() As Collection
    Set PrjItemItemSubs = ItemSubs
End Property

Property Get PrjItemItemFunctions() As Collection
    Set PrjItemItemFunctions = ItemFunctions
End Property

Property Get PrjItemItemProperties() As Collection
    Set PrjItemItemProperties = ItemProperties
End Property

Property Get PrjItemItemEvents() As Collection
    Set PrjItemItemEvents = ItemEvents
End Property

Sub AddSPF(AddName As String, SPFType As SPFTypes)
    If SPFType = SPF_Sub Then
        ItemSubs.Add AddName
    ElseIf SPFType = SPF_Function Then
        ItemFunctions.Add AddName
    ElseIf SPFType = SPF_Property Then
        ItemProperties.Add AddName
    ElseIf SPFType = SPF_Event Then
        ItemEvents.Add AddName
    End If
End Sub

'------------------------------------------------------------------------------

Property Get PrjItemVariables() As Long
    PrjItemVariables = ItemVariables
End Property

Property Let PrjItemVariables(AddAmount As Long)
    ItemVariables = ItemVariables + AddAmount
End Property

Sub AddToStatement(STTYPE As StatementType)
    If STTYPE = STDO Then
        ItemStatementDO = ItemStatementDO + 1
    ElseIf STTYPE = STFOR Then
        ItemStatementFOR = ItemStatementFOR + 1
    ElseIf STTYPE = STIF Then
        ItemStatementIF = ItemStatementIF + 1
    ElseIf STTYPE = STWHILE Then
        ItemStatementWHILE = ItemStatementWHILE + 1
    Else
        ItemStatementSELECT = ItemStatementSELECT + 1
    End If
End Sub

Property Get PrjItemStatements(STTYPE As StatementType)
    If STTYPE = STDO Then
        PrjItemStatements = ItemStatementDO
    ElseIf STTYPE = STFOR Then
        PrjItemStatements = ItemStatementFOR
    ElseIf STTYPE = STIF Then
        PrjItemStatements = ItemStatementIF
    ElseIf STTYPE = STWHILE Then
        PrjItemStatements = ItemStatementWHILE
    Else
        PrjItemStatements = ItemStatementSELECT
    End If
End Property

Property Get PrjItemConstants() As Long
    PrjItemConstants = ItemConstants
End Property

Property Let PrjItemConstants(AddAmount As Long)
    ItemConstants = PrjItemConstants + AddAmount
End Property

Property Get PrjItemTypes() As Long
    PrjItemTypes = ItemTypes
End Property

Property Let PrjItemTypes(AddAmount As Long)
    ItemTypes = PrjItemTypes + AddAmount
End Property

Property Get PrjItemEnums() As Long
    PrjItemEnums = ItemEnums
End Property

Property Let PrjItemEnums(AddAmount As Long)
    ItemEnums = PrjItemEnums + AddAmount
End Property

Private Sub Class_Initialize()
    Set ItemSubs = New Collection
    Set ItemFunctions = New Collection
    Set ItemProperties = New Collection
    Set ItemEvents = New Collection
End Sub
