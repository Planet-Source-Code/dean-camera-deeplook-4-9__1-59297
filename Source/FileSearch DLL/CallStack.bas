Attribute VB_Name = "ModCallStack"
' THIS MODULE WAS NOT WRITTEN BY EN-TECH. I CANNOT OFFER ANY SUPPORT FOR THIS MODULE.

' CallStack Module, by Paul Welter

Option Explicit

Private m_Stack() As typeStack

Private Type typeStack
    tModuleName As String
    tMethodName As String
    tParams As String
    tGlobalParams As String
End Type

Public pubLastDLLErrorNum As Long
Public pubLastDLLError As String

Public Sub RuntimeError( _
       ByVal ModuleName As String, _
       ByVal MethodName As String, _
       ByVal objErr As ErrObject, _
       Optional ByVal ErrLine As Long, _
       Optional ByVal Description As String)
    
    Dim strMsgText As String '  text of error report
    
    pubLastDLLErrorNum = objErr.Number
    pubLastDLLError = objErr.Description
        
    Debug.Print "ERROR: " & pubLastDLLErrorNum & " - " & pubLastDLLError
        
    On Error Resume Next
    
    strMsgText = vbCrLf & vbCrLf & Now & "  RunTime Error in " & App.EXEName & vbCrLf & vbCrLf
    strMsgText = strMsgText & "RunTime Error: " & CStr(pubLastDLLErrorNum) & " - " & _
       pubLastDLLError & vbCrLf
    strMsgText = strMsgText & vbTab & "Module: " & ModuleName & _
       vbCrLf & vbTab & "Method: " & MethodName & _
       vbCrLf & vbTab & "Line Number: " & CStr(ErrLine) & _
       vbCrLf & vbTab & "Description: :" & Description & _
       vbCrLf
       
    strMsgText = strMsgText & vbCrLf & "Call Stack:" & vbCrLf & StackRead
    
    'writing to event log
    Dim oEventLog As ClsEventLog
    Set oEventLog = New ClsEventLog
    oEventLog.WriteEvent strMsgText, App.EXEName, 5000, evnERROR
    Set oEventLog = Nothing
    
End Sub

Public Sub StackAdd( _
       ByVal ModuleName As String, _
       ByVal MethodName As String, _
       Optional Params As String = "", _
       Optional GlobalParams As String = "" _
       )
Attribute StackAdd.VB_Description = "StackAdd Adds a procedure call to the call stack"
    ' ******************************************************************************
    ' Routine:           StackAdd
    ' Created by:        Paul Welter
    ' Date-Time:         8/28/00 2:46:42 PM
    '
    'Document!VB Tags
    ' ##BD StackAdd Adds a procedure call to the call stack
    ' ##PD ModuleName The name of the class or module
    ' ##PD MethodName The name of the method that is currently being executing
    ' ##PD Params List of input parameter for the current method
    ' ##PD GlobalParams List of global variables and values
    ' ******************************************************************************
    On Error Resume Next
    
    Call StackInt 'setting up stack
    
    Dim x As Integer
        
    x = UBound(m_Stack) + 1
    ReDim Preserve m_Stack(x) As typeStack
    
    With m_Stack(x)
        .tModuleName = ModuleName
        .tMethodName = MethodName
        .tParams = Params
        .tGlobalParams = GlobalParams
    End With
       
End Sub

Public Sub StackRemove()
Attribute StackRemove.VB_Description = "StackRemove Removes the last call from the stack"
    ' ******************************************************************************
    ' Routine:           StackRemove
    ' Created by:        Paul Welter
    ' Date-Time:         8/28/002:47:51 PM
    '
    'Document!VB Tags
    ' ##BD StackRemove Removes the last call from the stack
    ' ******************************************************************************
    On Error Resume Next
    
    Call StackInt 'setting up stack
    
    Dim x As Integer
    x = UBound(m_Stack) - 1
    ReDim Preserve m_Stack(x) As typeStack
    
End Sub

Public Function StackRead() As String
Attribute StackRead.VB_Description = "StackRead Formats the CallStack to a string message for error logging, call this when an error occurs."
    ' ******************************************************************************
    ' Routine:           StackRead
    ' Created by:        Paul Welter
    ' Date-Time:         8/28/002:47:57 PM
    '
    'Document!VB Tags
    ' ##BD StackRead Formats the CallStack to a string message for error logging, _
      call this when an error occurs.
    ' ##RD Returns a formated string message
    ' ******************************************************************************
    On Error Resume Next
    Dim strTemp As String
    Dim x As Integer
    
    Call StackInt 'setting up stack
    
    strTemp = ""
    If UBound(m_Stack) > 0 Then
        For x = UBound(m_Stack) To 1 Step -1
            strTemp = strTemp & _
               vbTab & "Module:   " & m_Stack(x).tModuleName & vbCrLf & _
               vbTab & "Method:   " & m_Stack(x).tMethodName & "(" & m_Stack(x).tParams & ")" & vbCrLf & _
               vbTab & "Global:   " & m_Stack(x).tGlobalParams & vbCrLf & vbCrLf
        Next
    End If
    StackRead = strTemp
    
End Function

Private Function StackInt()

    On Error Resume Next
    
    Dim x As Integer
    
    x = UBound(m_Stack)
    If Err.Number <> 0 Then
        ReDim Preserve m_Stack(0) As typeStack
    End If
    
End Function
