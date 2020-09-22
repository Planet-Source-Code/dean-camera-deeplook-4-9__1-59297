Attribute VB_Name = "ModVariableHandler"
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

'-----------------------------------------------------------------------------------------------
'                                 VB6 UNUSED VARIABLE SCANNER SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   This module contains a simple are reasonably fast unused variable scanner. It is designed to
'   be as acurate as possible without sacrificing speed.
'
'   The VB unused variable scanning engine contained in this module is (C) Dean Camera.
'-----------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------------------------------------------------
Private CurrFilename As String
Private CurrProjName As String
Private LocalVars As Collection
Private LocalVarsLoc As Collection
Private FilesRoot As String
Private Project1Percent As Single

Private Lop As Integer
Private VarNames As Variant, Var As Variant
'-----------------------------------------------------------------------------------------------

Public Sub ClearGlobals()
    Set GlobalVars = Nothing
    Set GlobalVars = New Collection

    Set GlobalVarsLoc = Nothing
    Set GlobalVarsLoc = New Collection
End Sub

Public Sub ClearLocals()
    Set LocalVars = Nothing
    Set LocalVars = New Collection

    Set LocalVarsLoc = Nothing
    Set LocalVarsLoc = New Collection
End Sub

Public Sub AnalyseVBProjectForVars(ByVal ProjectFileName As String)
    Dim Linedata As String, i As Long, DonePrj As Long

    On Error Resume Next

    ProjectFileName = FixRelPath(ProjectFileName)

    If FileExists(FilesRootDirectory & ProjectFileName) = False And FileExists(ProjectFileName) = False Then
        ' No warning nessesary as it would have been picked up by the normal scanner
        Exit Sub
    End If

    CurrProjName = ExtractFileName(ProjectFileName)

    FilesRoot = Mid$(ProjectFileName, 1, InStrRev(ProjectFileName, "\"))

    Open ProjectFileName For Input As #2

    With FrmSelProject
        .pgbAPB.Value = 0
        .pgbAPB.Color = 8421631
        .lblScanPhase.Caption = "Unused Variable Scan Phase"
        .lblScanPhase.ForeColor = 8421631
    End With

    FrmSelProject.imgCurrScanObjType.Picture = FrmResults.ilstImages.ListImages(29).Picture

    DoEvents

    Project1Percent = (100 / LOF(2))

    Do While Not EOF(2)
        Line Input #2, Linedata ' Get the line data from the project file
        LookAtPRJLine Linedata
        FrmSelProject.pgbAPB.Value = Project1Percent * DonePrj ' Increase progressbar
        DonePrj = DonePrj + Len(Linedata)

        DoEvents ' Prevent VB from locking up when processing
    Loop

    Close #2

    With FrmResults.lstVarList
        For i = 1 To GlobalVars.Count
            .ListItems.Add , , CurrProjName
            .ListItems(.ListItems.Count).ListSubItems.Add , , GlobalVarsLoc(i)
            .ListItems(.ListItems.Count).ListSubItems.Add , , GlobalVars(i)
            .ListItems(.ListItems.Count).ListSubItems.Add , , "Global"
            If Left(.ListItems(.ListItems.Count).ListSubItems(2).Text, 5) = "Const" Then .ListItems(.ListItems.Count).ListSubItems(2).ForeColor = RGB(100, 100, 100)
            .ListItems(.ListItems.Count).ListSubItems(3).ForeColor = RGB(140, 0, 140)
        Next
    End With
End Sub

Public Sub ScanFile(ByVal Linedata As String) ' No statistics are being generated, so the headder dosn't need to be ignored
    Dim i As Long

    If InStr(1, Linedata, ";") > 0 Then Linedata = Mid$(Linedata, InStr(1, Linedata, ";") + 1)

    On Error Resume Next
    Linedata = Trim$(Linedata)

    If FileExists(FilesRoot & Linedata) = False Then
        If FileExists(Linedata) = False Then
            Exit Sub
        Else
            Open Linedata For Input As #3
        End If
    Else
        Open FilesRoot & Linedata For Input As #3
    End If

    Do While Not EOF(3)
        Line Input #3, Linedata
        If Left$(Linedata, 20) = "Attribute VB_Name = " Then CurrFilename = Mid$(Linedata, 21)

        Linedata = Trim$(Linedata)

        CheckIfVarIsUsed Linedata
        CheckIfIsDeclare Linedata
    Loop

    Close #3

    For i = 1 To LocalVars.Count
        If Trim$(LocalVarsLoc(i)) = "" Or Trim$(LocalVars(i)) = "" Then
            ' Don't use a <> operator, only works like this for some reason
        Else
            With FrmResults.lstVarList
                .ListItems.Add , , CurrProjName
                .ListItems(.ListItems.Count).ListSubItems.Add , , LocalVarsLoc(i)
                .ListItems(.ListItems.Count).ListSubItems.Add , , LocalVars(i)
                .ListItems(.ListItems.Count).ListSubItems.Add , , "Local"
                If Left(.ListItems(.ListItems.Count).ListSubItems(2).Text, 5) = "Const" Then .ListItems(.ListItems.Count).ListSubItems(2).ForeColor = RGB(0, 150, 0)
                .ListItems(.ListItems.Count).ListSubItems(3).ForeColor = vbBlue
            End With
        End If
    Next
End Sub

Private Sub LookAtPRJLine(ByVal Linedata As String)
    If Left$(Linedata, 5) = "Form=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 6)
    ElseIf Left$(Linedata, 7) = "Module=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 8)
    ElseIf Left$(Linedata, 6) = "Class=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 7)
    ElseIf Left$(Linedata, 12) = "UserControl=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 13)
    ElseIf Left$(Linedata, 13) = "PropertyPage=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 14)
    ElseIf Left$(Linedata, 13) = "UserDocument=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 14)
    ElseIf Left$(Linedata, 9) = "Designer=" Then
        ClearLocals
        ScanFile Mid$(Linedata, 10)
    End If
End Sub

Private Sub CheckVar(ByVal VType As Integer, ByVal i As Long, ByVal Linedata As String)
    Dim SearchFor As String, TempByte As String * 1, TempInt As Long

    Select Case VType
        Case 1
            SearchFor = GlobalVars(i)
        Case 2
            SearchFor = LocalVars(i)
    End Select
    
    TempInt = InStr(1, Linedata, SearchFor)
    If TempInt > 1 Then
        TempByte = Mid$(Linedata, TempInt - 1, 1)
        If InStr(1, " (.", TempByte) > 0 Then GoTo CheckEnd
    Else
        GoTo CheckEnd
    End If

    Exit Sub
CheckEnd:

    If Len(Linedata) <> (TempInt + Len(SearchFor)) Then
        TempByte = Mid$(Linedata, TempInt + Len(SearchFor), 1)
        If InStr(1, " )(.,", TempByte) > 0 Then GoTo IncUsed
    Else
        GoTo IncUsed
    End If

    Exit Sub
IncUsed:

    Select Case VType
        Case 1
            GlobalVars.Remove i
            GlobalVarsLoc.Remove i
        Case 2
            LocalVars.Remove i
            LocalVarsLoc.Remove i
    End Select
End Sub

Private Sub CheckIfVarIsUsed(ByVal Linedata As String)
    Dim i As Long

    If InStr(1, Linedata, "'") Then Linedata = Mid$(Linedata, 1, InStr(1, Linedata, "'"))

    On Error GoTo 0

    For i = 1 To GlobalVars.Count
        If InStr(1, Linedata, GlobalVars(i)) Then CheckVar 1, i, Linedata
    Next
    
    For i = 1 To LocalVars.Count
        If InStr(1, Linedata, LocalVars(i)) Then CheckVar 2, i, Linedata
    Next
End Sub

Private Sub CheckIfIsDeclare(ByVal Linedata As String)
    If Left$(Linedata, 4) = "Dim " Then
        Lop = 4
    ElseIf Left$(Linedata, 7) = "Static " Then
        Lop = 7
    ElseIf Left$(Linedata, 8) = "Private " Then
        Lop = 8
    ElseIf Left$(Linedata, 8) = "Public " Then
        Lop = 7
    Else
        Exit Sub
    End If

    If InStr(1, Linedata, " Sub ") Then
        Exit Sub
    ElseIf InStr(1, Linedata, " Function ") Then
        Exit Sub
    ElseIf InStr(1, Linedata, " Property ") Then
        Exit Sub
    ElseIf InStr(1, Linedata, " Type ") Then
        Exit Sub
    ElseIf InStr(1, Linedata, " Enum ") Then
        Exit Sub
    ElseIf InStr(1, Linedata, " Event ") Then
        Exit Sub
    ElseIf InStr(1, Linedata, " Const ") And IncludeConsts = True Then
        Exit Sub
    ElseIf Mid(Linedata, 1, 4) = "Sub " Then
        Exit Sub
    End If

    If InStr(1, Linedata, "'") Then Linedata = Mid$(Linedata, 1, InStr(1, Linedata, "'") - 1)

    If InStr(1, Linedata, ",") = 0 Then
        If InStr(1, Linedata, " As ") <> 0 Then
            LocalVarsLoc.Add FixQuotes(CurrFilename)
            LocalVars.Add TrimJunk(Mid$(Linedata, Lop, InStr(1, Linedata, " As ") - Lop))
        Else
            LocalVarsLoc.Add FixQuotes(CurrFilename)
            LocalVars.Add TrimJunk(Mid$(Linedata, Lop))
        End If
    Else
        Linedata = Mid$(Linedata, Lop)
        If InStr(1, Linedata, ":") Then Linedata = Mid(Linedata, 1, InStr(1, Linedata, ":"))

        VarNames = Split(Linedata, ",")
        For Var = LBound(VarNames) To UBound(VarNames)
            If InStr(1, VarNames(Var), " As ") <> 0 Then
                LocalVarsLoc.Add FixQuotes(CurrFilename)
                LocalVars.Add TrimJunk(Mid$(VarNames(Var), 1, InStr(1, VarNames(Var), " As ")))
            Else
                LocalVarsLoc.Add FixQuotes(CurrFilename)
                LocalVars.Add TrimJunk(Mid$(VarNames(Var), 1))
            End If
        Next
    End If
End Sub

Private Function FixQuotes(ByVal Data As String) As String
    If Left$(Data, 1) = """" Then
        Data = Mid$(Data, 2)
        If Right$(Data, 1) = """" Then Data = Mid$(Data, 1, Len(Data) - 1)
    End If

    FixQuotes = Data
End Function
