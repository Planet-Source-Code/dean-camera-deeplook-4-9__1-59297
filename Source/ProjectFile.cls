VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProjectFile"
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

Private PrjName As String
Private PrjTitle As String
Private PrjType As String
Private PrjStartUpItem As String
Private PrjPath As String
Private PrjRefComCount As Long

Private PrjClass As Long
Private PrjModule As Long
Private PrjForm As Long
Private PrjUserControl As Long
Private PrjUserDocument As Long
Private PrjPropertyPage As Long
Private PrjDesigner As Long

Private PrjTotalLines As Long
Private PrjTotalLinesNB As Long
Private PrjTotalCommentLines As Long

Private PrjVersion As String

Private PrjVariables As Long
Private PrjConstants As Long
Private PrjTypes As Long
Private PrjEnums As Long

Private PrjRefCom As Collection
Private PrjDecDll As Collection
Private PrjCreateObj As Collection

Private PrjSubLines As Long
Private PrjFuncLines As Long
Private PrjPropLines As Long

Private PrjEXEPath As String
Private PrjEXEFName As String

'------------------------------------------------------------------------------

Private Function FixProjType(PrjType As String)
    FixProjType = UCase(PrjType)

    Select Case FixProjType
        Case "EXE"
            FixProjType = "Standard EXE"
        Case "OLEEXE"
            FixProjType = "ActiveX EXE"
        Case "OLEDLL"
            FixProjType = "ActiveX DLL"
        Case Else
            FixProjType = "ActiveX Control"
    End Select
End Function

'------------------------------------------------------------------------------

Property Get ProjectName() As String
    ProjectName = PrjName
End Property

Property Let ProjectName(NewName As String)
    PrjName = NewName
End Property

'------------------------------------------------------------------------------

Property Get ProjectRefComCount() As Long
    ProjectRefComCount = PrjRefComCount
End Property

Property Let ProjectRefComCount(AddAmount As Long)
    PrjRefComCount = PrjRefComCount + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get ProjectTitle() As String
    ProjectTitle = PrjTitle
End Property

Property Let ProjectTitle(NewTitle As String)
    PrjTitle = NewTitle
End Property

'------------------------------------------------------------------------------

Property Get ProjectProjectType() As String
    ProjectProjectType = FixProjType(PrjType)
End Property

Property Let ProjectProjectType(NewType As String)
    PrjType = NewType
End Property

'------------------------------------------------------------------------------

Property Get ProjectStartupItem() As String
    ProjectStartupItem = PrjStartUpItem
End Property

Property Let ProjectStartupItem(NewStartupItem As String)
    PrjStartUpItem = NewStartupItem
End Property

'------------------------------------------------------------------------------

Property Get ProjectPath() As String
    ProjectPath = PrjPath
End Property

Property Let ProjectPath(NewPath As String)
    PrjPath = NewPath
End Property

'------------------------------------------------------------------------------

Property Get ProjectSubLines() As Long
    ProjectSubLines = PrjSubLines
End Property

Property Let ProjectSubLines(AddAmount As Long)
    PrjSubLines = PrjSubLines + AddAmount
End Property

Property Get ProjectFuncLines() As Long
    ProjectFuncLines = PrjFuncLines
End Property

Property Let ProjectFuncLines(AddAmount As Long)
    PrjFuncLines = PrjFuncLines + AddAmount
End Property

Property Get ProjectPropLines() As Long
    ProjectPropLines = PrjPropLines
End Property

Property Let ProjectPropLines(AddAmount As Long)
    PrjPropLines = PrjPropLines + AddAmount
End Property

'------------------------------------------------------------------------------

Property Get ProjectForms() As Long
    ProjectForms = PrjForm
End Property

Property Let ProjectForms(AddAmount As Long)
    PrjForm = PrjForm + AddAmount
End Property

Property Get ProjectModules() As Long
    ProjectModules = PrjModule
End Property

Property Let ProjectModules(AddAmount As Long)
    PrjModule = PrjModule + AddAmount
End Property

Property Get ProjectClasses() As Long
    ProjectClasses = PrjClass
End Property

Property Let ProjectClasses(AddAmount As Long)
    PrjClass = PrjClass + AddAmount
End Property

Property Get ProjectUserControls() As Long
    ProjectUserControls = PrjUserControl
End Property

Property Let ProjectUserControls(AddAmount As Long)
    PrjUserControl = PrjUserControl + AddAmount
End Property

Property Get ProjectUserDocuments() As Long
    ProjectUserDocuments = PrjUserDocument
End Property

Property Let ProjectUserDocuments(AddAmount As Long)
    PrjUserDocument = PrjUserDocument + AddAmount
End Property

Property Get ProjectPropertyPages() As Long
    ProjectPropertyPages = PrjPropertyPage
End Property

Property Let ProjectPropertyPages(AddAmount As Long)
    PrjPropertyPage = PrjPropertyPage + AddAmount
End Property

Property Get ProjectDesigners() As Long
    ProjectDesigners = PrjDesigner
End Property

Property Let ProjectDesigners(AddAmount As Long)
    PrjDesigner = PrjDesigner + AddAmount
End Property

Property Get ProjectLines() As Long
    ProjectLines = PrjTotalLines
End Property

Property Let ProjectLines(AddAmount As Long)
    PrjTotalLines = PrjTotalLines + AddAmount
End Property

Property Get ProjectCommentLines() As Long
    ProjectCommentLines = PrjTotalCommentLines
End Property

Property Let ProjectCommentLines(AddAmount As Long)
    PrjTotalCommentLines = PrjTotalCommentLines + AddAmount
End Property

Property Get ProjectLinesNB() As Long
    ProjectLinesNB = PrjTotalLinesNB
End Property

Property Let ProjectLinesNB(AddAmount As Long)
    PrjTotalLinesNB = PrjTotalLinesNB + AddAmount
End Property

Property Get ProjectVariables() As Long
    ProjectVariables = PrjVariables
End Property

Property Let ProjectVariables(AddAmount As Long)
    PrjVariables = PrjVariables + AddAmount
End Property

Property Get ProjectConstants() As Long
    ProjectConstants = PrjConstants
End Property

Property Let ProjectConstants(AddAmount As Long)
    PrjConstants = PrjConstants + AddAmount
End Property

Property Get ProjectTypes() As Long
    ProjectTypes = PrjTypes
End Property

Property Let ProjectTypes(AddAmount As Long)
    PrjTypes = PrjTypes + AddAmount
End Property

Property Get ProjectEnums() As Long
    ProjectEnums = PrjEnums
End Property

Property Let ProjectEnums(AddAmount As Long)
    PrjEnums = PrjEnums + AddAmount
End Property

Property Get ProjectVersion() As String
    ProjectVersion = PrjVersion
End Property

Property Let ProjectVersion(AddString As String)
    PrjVersion = PrjVersion & AddString
End Property

Property Get ProjectRefCom() As Collection
    Set ProjectRefCom = PrjRefCom
End Property

Property Let ProjectRefCom(AddData)
    ProjectRefCom.Add AddData
End Property

Property Get ProjectDecDlls() As Collection
    Set ProjectDecDlls = PrjDecDll
End Property

Property Let ProjectDecDlls(AddData)
    PrjDecDll.Add AddData
End Property

Property Let ProjectCreateObjects(AddData)
    PrjCreateObj.Add AddData
End Property

Property Get ProjectCreateObjects() As Collection
    Set ProjectCreateObjects = PrjCreateObj
End Property

Property Get ProjectEXEPath() As String
    ProjectEXEPath = PrjEXEPath
End Property

Property Let ProjectEXEPath(Path As String)
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    PrjEXEPath = Path
End Property

Property Get ProjectEXEFName() As String
    ProjectEXEFName = PrjEXEFName
End Property

Property Let ProjectEXEFName(FName As String)
    PrjEXEFName = FName
End Property

Private Sub Class_Initialize()
    Set PrjRefCom = New Collection
    Set PrjDecDll = New Collection
    Set PrjCreateObj = New Collection
End Sub
