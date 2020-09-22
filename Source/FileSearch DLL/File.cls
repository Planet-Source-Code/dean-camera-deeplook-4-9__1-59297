VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2035
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'               THIS CLASS WAS NOT WRITTEN BY EN-TECH.
'             I CANNOT OFFER ANY SUPPORT FOR THIS CLASS.

Option Explicit

Private m_FileName As String
Private m_FileSize As Double
Private m_Created As Date
Private m_Modified As Date
Private m_Accessed As Date
Private m_FilePath As String
Private m_FileAttributes As Long
Private m_FileNameFull As String

Public Property Get FileAttributes() As Long
Attribute FileAttributes.VB_Description = "Specifies the file attributes of the file found.  "
Attribute FileAttributes.VB_HelpID = 2037
    '##BLOCK_DESCRIPTION Specifies the file attributes of the file found.
    FileAttributes = m_FileAttributes
End Property

Public Property Let FileAttributes(ByVal Value As Long)
    m_FileAttributes = Value
End Property

Public Property Get FilePath() As String
Attribute FilePath.VB_HelpID = 2038
    FilePath = m_FilePath
End Property
Public Property Let FilePath(ByVal Value As String)
    m_FilePath = Value
End Property

Public Property Get Accessed() As Date
Attribute Accessed.VB_Description = "Specifies the time that the file was last accessed.  "
Attribute Accessed.VB_HelpID = 2039
    '##BLOCK_DESCRIPTION Specifies the time that the file was last accessed.
    Accessed = m_Accessed
End Property
Public Property Let Accessed(ByVal Value As Date)
    m_Accessed = Value
End Property

Public Property Get Modified() As Date
Attribute Modified.VB_Description = "Specifies the time that the file was last written to.  "
Attribute Modified.VB_HelpID = 2040
    '##BLOCK_DESCRIPTION Specifies the time that the file was last written to.
    Modified = m_Modified
End Property
Public Property Let Modified(ByVal Value As Date)
    m_Modified = Value
End Property

Public Property Get Created() As Date
Attribute Created.VB_Description = "Specifies the time the file was created.  "
Attribute Created.VB_HelpID = 2041
    '##BLOCK_DESCRIPTION Specifies the time the file was created.
    Created = m_Created
End Property
Public Property Let Created(ByVal Value As Date)
    m_Created = Value
End Property

Public Property Get FileSize() As Double
Attribute FileSize.VB_Description = "Specifies the file size, in bytes.  "
Attribute FileSize.VB_HelpID = 2042
    '##BLOCK_DESCRIPTION Specifies the file size, in bytes.
    FileSize = m_FileSize
End Property
Public Property Let FileSize(ByVal Value As Double)
    m_FileSize = Value
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "A string that is the name of the file.  "
Attribute FileName.VB_HelpID = 2043
    '##BLOCK_DESCRIPTION A string that is the name of the file.
    FileName = m_FileName
End Property
Public Property Let FileName(ByVal Value As String)
    m_FileName = Value
End Property

Public Property Let FileNameFull(ByVal Value As String)
    m_FileNameFull = Value
End Property
Public Property Get FileNameFull() As String
    FileNameFull = m_FileNameFull
End Property
