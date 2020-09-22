Attribute VB_Name = "ModFileSearchHandler"
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

Private BatFiles As Collection

Sub CopyDLLOCX()
    Dim i As Long, z As Long, NodeNum As Long, AllowAdd As Boolean

    FrmResults.sbrStatus.Text = "Commencing File Copy..."

    Set BatFiles = New Collection

    On Error Resume Next

    MkDir GetRootDirectory(ProjectPath) & "Res\"

    With FrmCopyReport
        .lblCopyDir.Caption = GetRootDirectory(ProjectPath) & "Res\"
        .tvwItemsTV.Nodes.Clear
        .tvwNonCopyItemsTV.Nodes.Clear
        .Caption = "DeepLook Copy Required Files Report - Copying..."
        .btnCloseButton.Enabled = False
        .pgbPercentBar.Value = 0
        DoEvents        ' \
        .Show           ' |
        DoEvents        ' | Ensures the copy report form is visible before copying
        .Visible = True ' |
        DoEvents        ' /
    End With


    With FrmResults.TreeView.Nodes
        For i = 1 To .Count

            ' Adds the Declared DLLs to the Copy Report
            AllowAdd = True
            NodeNum = InStr(1, .Item(i).Key, "DECDLLS_") ' Is a DecDLL
            If NodeNum <> 0 Then
                If .Item(i).Image = "DLL" Then ' Prevent System (non-copy) DLLs from being showed
                    If InStrRev(.Item(i).Key, "_") = NodeNum + 7 Then ' Only allow the root item node (no DLL information nodes)
                        For z = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                            If UCase(FrmCopyReport.tvwItemsTV.Nodes(z).Text) = UCase(Mid$(.Item(i).Key, NodeNum + 8)) Then AllowAdd = False
                        Next

                        If AllowAdd = True Then
                            FrmCopyReport.tvwItemsTV.Nodes.Add , , .Item(i).Key, Mid$(.Item(i).Key, NodeNum + 8), "DLL"
                            BatFiles.Add Mid$(.Item(i).Key, NodeNum + 8)
                        End If
                    End If
                Else
                    If InStrRev(.Item(i).Key, "_") = NodeNum + 7 Then ' Only allow the root item node (no DLL information nodes)
                        For z = 1 To FrmCopyReport.tvwNonCopyItemsTV.Nodes.Count
                            If UCase(FrmCopyReport.tvwNonCopyItemsTV.Nodes(z).Text) = UCase(Mid$(.Item(i).Key, NodeNum + 8)) Then AllowAdd = False
                        Next

                        If AllowAdd = True Then
                            FrmCopyReport.tvwNonCopyItemsTV.Nodes.Add , , .Item(i).Key, Mid$(.Item(i).Key, NodeNum + 8), "SysDLL"
                        End If
                    End If
                End If
            End If

            ' Adds the Reference DLLs to the Copy Report
            NodeNum = InStr(1, .Item(i).Key, "REFCOM_REFERENCE_") ' Is a DLL
            If NodeNum <> 0 Then
                If .Item(i).Image = "DLL" Then ' Prevent System (non-copy) DLLs from being showed
                    If InStrRev(.Item(i).Key, "_") = NodeNum + 16 Then
                        For z = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                            If UCase(FrmCopyReport.tvwItemsTV.Nodes(z).Text) = UCase(Mid$(.Item(i).Key, NodeNum + 17)) Then AllowAdd = False
                        Next

                        If AllowAdd = True Then
                            FrmCopyReport.tvwItemsTV.Nodes.Add , , .Item(i).Key, Mid$(.Item(i).Key, NodeNum + 17), "DLL"
                            BatFiles.Add Mid$(.Item(i).Key, NodeNum + 17)
                        End If
                    End If
                Else
                    For z = 1 To FrmCopyReport.tvwNonCopyItemsTV.Nodes.Count
                        If UCase(FrmCopyReport.tvwNonCopyItemsTV.Nodes(z).Text) = UCase(Mid$(.Item(i).Key, NodeNum + 17)) Then AllowAdd = False
                    Next

                    If AllowAdd = True And InStrRev(.Item(i).Key, "_") = NodeNum + 16 Then
                        FrmCopyReport.tvwNonCopyItemsTV.Nodes.Add , , .Item(i).Key, Mid$(.Item(i).Key, NodeNum + 17), "SysDLL"
                    End If
                End If
            End If

            ' Adds Components to the Copy Report
            NodeNum = InStr(1, .Item(i).Key, "REFCOM_COMPONENT_") ' Is a Component
            If NodeNum <> 0 Then
                If InStrRev(.Item(i).Key, "_") = NodeNum + 16 Then
                    For z = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                        If UCase(FrmCopyReport.tvwItemsTV.Nodes(z).Text) = UCase(Mid$(.Item(i).Key, NodeNum + 18)) Then AllowAdd = False
                    Next

                    If AllowAdd = True Then
                        FrmCopyReport.tvwItemsTV.Nodes.Add , , .Item(i).Key, Mid$(.Item(i).Key, NodeNum + 18), "Component"
                        BatFiles.Add Mid$(.Item(i).Key, NodeNum + 18)
                    End If
                End If
            End If

            ' Add CreateObject statements to Copy Report
            If .Item(i).Image = "CreateObject" Then
                FrmCopyReport.tvwManualCopyTV.Nodes.Add , , .Item(i).Key, .Item(i).Text & " (CreateObject Statement)", "CreateObject"
            End If

            DoEvents
        Next
    End With

    For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
        z = (100 / FrmCopyReport.tvwItemsTV.Nodes.Count) * i
        Findfile Mid$(FrmCopyReport.tvwItemsTV.Nodes.Item(i).Key, NodeNum + 8)
        FrmCopyReport.pgbPercentBar.Value = z
        DoEvents
    Next

    If FileExists(App.Path & "\FileRegister.exe") = False Then
        ModFileRegisterBatCreator.CreateHeadder GetRootDirectory(ProjectPath) & "Res\FileRegister.bat"

        For i = 1 To BatFiles.Count
            ModFileRegisterBatCreator.AddRegAndCopyFile BatFiles(i), i, BatFiles.Count
        Next

        ModFileRegisterBatCreator.AddFooter GetRootDirectory(ProjectPath) & "Res\FileRegister.bat"
    Else
        FileCopy App.Path & "\FileRegister.exe", GetRootDirectory(ProjectPath) & "Res\FileRegister.exe"
    End If

    FrmCopyReport.Caption = "DeepLook Copy Required Files Report - Done."
    FrmCopyReport.btnCloseButton.Enabled = True
    FrmResults.sbrStatus.Text = "File Copy Complete. Files copied to " & GetRootDirectory(ProjectPath) & "Res\" & "."
End Sub

Private Function GetRootDirectory(FileName As String) As String
    Dim SlashPos As Long

    SlashPos = InStrRev(FileName, "\")

    If SlashPos <> 0 Then
        GetRootDirectory = Mid$(FileName, 1, SlashPos)
    End If

    If GetRootDirectory = "" Then Exit Function

    If Right$(GetRootDirectory, 1) <> "\" Then GetRootDirectory = GetRootDirectory & "\"
End Function

Private Sub Findfile(FileName As String)
    Dim FileSearch As ClsSearch, i As Long

    Set FileSearch = New ClsSearch

    If InStr(1, FileName, ".") = 0 Then Exit Sub

    FileName = Trim$(Mid$(FileName, InStrRev(FileName, "_") + 1))

    DoEvents

    FrmResults.sbrStatus.Text = "Commencing File Copy: " & FileName

    For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
        If FrmCopyReport.tvwItemsTV.Nodes(i).Text = FileName Then
            FrmCopyReport.tvwItemsTV.Nodes(i).Image = "CurrentCopy"
            Exit For
        End If
    Next

    DoEvents

    With FileSearch
        .SearchFiles Environ("windir"), FileName, True

        If .Files.Count = 0 Then AltFindFile FileName: Exit Sub

        On Error Resume Next
        Err.Clear
        FileCopy .Files.Item(1).FileNameFull, GetRootDirectory(ProjectPath) & "Res\" & .Files.Item(1).FileName
        On Error GoTo 0

        For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
            If FrmCopyReport.tvwItemsTV.Nodes(i).Text = FileName Then
                If Err.Number = 0 Then
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Done"
                Else
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Error"
                End If
                Exit For
            End If
        Next
    End With
End Sub

Private Sub AltFindFile(FileName As String)
    Dim FileSearch As ClsSearch, Task As Long, i As Long

    Set FileSearch = New ClsSearch

    With FileSearch
        .SearchFiles GetRootDirectory(ProjectPath), FileName, True

        If .Files.Count = 0 Then
            Task = MsgBoxEx("Cannot find file """ & FileName & """ for copying. Would you like search drive C for it?", vbYesNo, "File Copy Error", , , , , PicError)
            DoEvents

            If Task = vbYes Then
                FrmResults.sbrStatus.Text = "Commencing File Copy: " & FileName & " [Searching C:\]"
                SearchCForFile FileName
            Else
                For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                    If FrmCopyReport.tvwItemsTV.Nodes(i).Text = FileName Then
                        FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Error"
                        Exit For
                    End If
                Next
                Exit Sub
            End If
        End If

        On Error Resume Next
        Err.Clear
        FileCopy .Files.Item(1).FileNameFull, GetRootDirectory(ProjectPath) & "Res\" & .Files.Item(1).FileName
        On Error GoTo 0

        For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
            If FrmCopyReport.tvwItemsTV.Nodes(i).Text = FileName Then
                If Err.Number = 0 Then
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Done"
                Else
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Error"
                End If
                Exit For
            End If
        Next
    End With
End Sub

Private Sub SearchCForFile(FileName As String)
    Dim FileSearch As ClsSearch, i As Long

    Set FileSearch = New ClsSearch

    With FileSearch
        .SearchFiles "C:\", FileName, True

        If .Files.Count = 0 Then
            MsgBoxEx "Cannot find file """ & FileName & """ for copying. Please manually copy this file.", vbOKOnly, "File Copy Error", , , , , PicError
            For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                If FrmCopyReport.tvwItemsTV.Nodes(i).Text = FileName Then
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Error"
                    Exit For
                End If
            Next

            Exit Sub
        End If

        On Error Resume Next
        Err.Clear
        FileCopy .Files.Item(1).FileNameFull, GetRootDirectory(ProjectPath) & "Res\" & .Files.Item(1).FileName
        On Error GoTo 0

        For i = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
            If FrmCopyReport.tvwItemsTV.Nodes(i).Text = FileName Then
                If Err.Number = 0 Then
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Done"
                Else
                    FrmCopyReport.tvwItemsTV.Nodes(i).Image = "Error"
                End If
                Exit For
            End If
        Next
    End With
End Sub
