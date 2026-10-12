Attribute VB_Name = "modAddinsPublish"
Option Explicit

Private Const ADDINS_MANIFEST_FILE As String = "addins-manifest.json"

Private mSharePointRootOverride As String
Private mLastAddinsPublishReport As String

Public Function CaptureVerificationRoot() As String
    CaptureVerificationRoot = ResolveAddinsRootPath()
End Function

Public Function VerifyAddinsPublished(Optional ByVal capturedRoot As String = "", Optional ByVal useCapturedRoot As Boolean = False) As Boolean
    Dim addinsRoot As String
    Dim addinNames As Variant
    Dim i As Long
    Dim targetPath As String
    Dim detail As String
    Dim hadFailure As Boolean

    On Error GoTo FailVerify

    mLastAddinsPublishReport = vbNullString
    If useCapturedRoot Then
        addinsRoot = capturedRoot
    Else
        addinsRoot = ResolveAddinsRootPath()
    End If
    If addinsRoot = "" Then
        detail = "PathSharePointRoot is not configured."
        LogDiagnosticEvent "ADDINS-PUBLISH", "Verify skipped|" & detail
        mLastAddinsPublishReport = detail
        Exit Function
    End If

    addinNames = GetRequiredAddinNames()
    For i = LBound(addinNames) To UBound(addinNames)
        targetPath = addinsRoot & CStr(addinNames(i))
        If Not FileExistsAddins(targetPath) Then
            detail = "Missing add-in: " & targetPath
            LogDiagnosticEvent "ADDINS-PUBLISH", detail
            hadFailure = True
        ElseIf SafeFileLenAddins(targetPath) <= 0 Then
            detail = "Zero-byte add-in: " & targetPath
            LogDiagnosticEvent "ADDINS-PUBLISH", detail
            hadFailure = True
        End If
    Next i

    VerifyAddinsPublished = Not hadFailure
    If VerifyAddinsPublished Then
        mLastAddinsPublishReport = "OK"
    Else
        mLastAddinsPublishReport = "Add-ins verification failed."
    End If
    Exit Function

FailVerify:
    mLastAddinsPublishReport = "VerifyAddinsPublished failed: " & Err.Description
    LogDiagnosticEvent "ADDINS-PUBLISH", mLastAddinsPublishReport
End Function

Public Function PublishAddins(ByVal sourceDir As String) As Boolean
    Dim addinsRoot As String
    Dim manifestPath As String
    Dim manifestBackupPath As String
    Dim stageRoot As String
    Dim backupRoot As String
    Dim addinNames As Variant
    Dim sourceRoot As String
    Dim sourcePath As String
    Dim stagePath As String
    Dim targetPath As String
    Dim backupPath As String
    Dim fileStatuses() As String
    Dim changed() As Boolean
    Dim hadTarget() As Boolean
    Dim i As Long
    Dim report As String
    Dim copyStatus As String
    Dim manifestText As String
    Dim changedCount As Long
    Dim currentStep As String

    On Error GoTo FailPublish

    mLastAddinsPublishReport = vbNullString
    currentStep = "resolve-root"
    addinsRoot = ResolveAddinsRootPath()
    If addinsRoot = "" Then
        report = "PathSharePointRoot is not configured."
        GoTo FailSoft
    End If

    currentStep = "normalize-source"
    sourceRoot = NormalizeFolderPathAddins(sourceDir)
    If sourceRoot = "" Then
        report = "Source add-ins folder is required."
        GoTo FailSoft
    End If
    If Not FolderExistsAddins(sourceRoot) Then
        report = "Source add-ins folder not found: " & sourceRoot
        GoTo FailSoft
    End If

    currentStep = "ensure-addins-root"
    EnsureFolderRecursiveAddins addinsRoot
    currentStep = "build-stage-root"
    stageRoot = BuildWorkingFolderAddins("stage")
    currentStep = "build-backup-root"
    backupRoot = BuildWorkingFolderAddins("backup")
    currentStep = "ensure-stage-root"
    EnsureFolderRecursiveAddins stageRoot
    currentStep = "ensure-backup-root"
    EnsureFolderRecursiveAddins backupRoot

    addinNames = GetRequiredAddinNames()
    ReDim fileStatuses(LBound(addinNames) To UBound(addinNames))
    ReDim changed(LBound(addinNames) To UBound(addinNames))
    ReDim hadTarget(LBound(addinNames) To UBound(addinNames))

    For i = LBound(addinNames) To UBound(addinNames)
        sourcePath = sourceRoot & CStr(addinNames(i))
        targetPath = addinsRoot & CStr(addinNames(i))
        stagePath = stageRoot & CStr(addinNames(i))

        currentStep = "validate-source:" & CStr(addinNames(i))
        If Not FileExistsAddins(sourcePath) Then
            report = "Source add-in missing: " & sourcePath
            LogDiagnosticEvent "ADDINS-PUBLISH", report
            GoTo FailSoft
        End If
        If SafeFileLenAddins(sourcePath) <= 0 Then
            report = "Source add-in is zero-byte: " & sourcePath
            LogDiagnosticEvent "ADDINS-PUBLISH", report
            GoTo FailSoft
        End If

        If FileExistsAddins(targetPath) And SafeFileLenAddins(targetPath) = SafeFileLenAddins(sourcePath) Then
            fileStatuses(i) = CStr(addinNames(i)) & "=SKIPPED"
        Else
            currentStep = "stage-copy:" & CStr(addinNames(i))
            CopyFileVerifiedAddins sourcePath, stagePath
            If SafeFileLenAddins(stagePath) <> SafeFileLenAddins(sourcePath) Then
                report = "Stage size mismatch for " & CStr(addinNames(i))
                LogDiagnosticEvent "ADDINS-PUBLISH", report
                GoTo FailSoft
            End If
            changed(i) = True
            changedCount = changedCount + 1
        End If
    Next i

    For i = LBound(addinNames) To UBound(addinNames)
        If changed(i) Then
            targetPath = addinsRoot & CStr(addinNames(i))
            backupPath = backupRoot & CStr(addinNames(i))
            stagePath = stageRoot & CStr(addinNames(i))

            currentStep = "detect-existing:" & CStr(addinNames(i))
            hadTarget(i) = FileExistsAddins(targetPath)
            currentStep = "backup-current:" & CStr(addinNames(i))
            If hadTarget(i) Then CopyFileVerifiedAddins targetPath, backupPath

            currentStep = "publish-target:" & CStr(addinNames(i))
            If Not modWarehouseSync.PublishFileToTargetPath(stagePath, targetPath, copyStatus) Then
                report = "Publish failed for " & CStr(addinNames(i)) & ": " & copyStatus
                LogDiagnosticEvent "ADDINS-PUBLISH", report
                GoTo RollbackChanges
            End If
            If SafeFileLenAddins(targetPath) <> SafeFileLenAddins(stagePath) Then
                report = "Published size mismatch for " & CStr(addinNames(i))
                LogDiagnosticEvent "ADDINS-PUBLISH", report
                GoTo RollbackChanges
            End If
            fileStatuses(i) = CStr(addinNames(i)) & "=" & copyStatus
        End If
    Next i

    currentStep = "backup-manifest"
    manifestPath = addinsRoot & ADDINS_MANIFEST_FILE
    manifestBackupPath = backupRoot & ADDINS_MANIFEST_FILE
    If FileExistsAddins(manifestPath) Then CopyFileVerifiedAddins manifestPath, manifestBackupPath

    currentStep = "build-manifest"
    manifestText = GetAddinsManifest()
    If manifestText = "" Then
        report = "GetAddinsManifest returned empty text."
        LogDiagnosticEvent "ADDINS-PUBLISH", report
        GoTo RollbackChanges
    End If
    currentStep = "write-manifest"
    If Not WriteTextFileAddins(manifestPath, manifestText) Then
        report = "Failed to write add-ins manifest: " & manifestPath
        LogDiagnosticEvent "ADDINS-PUBLISH", report
        GoTo RollbackChanges
    End If
    If SafeFileLenAddins(manifestPath) <= 0 Then
        report = "Manifest is zero-byte: " & manifestPath
        LogDiagnosticEvent "ADDINS-PUBLISH", report
        GoTo RollbackChanges
    End If

    currentStep = "verify-published"
    If Not VerifyAddinsPublished() Then
        report = "VerifyAddinsPublished failed after publish."
        GoTo RollbackChanges
    End If

    PublishAddins = True
    report = "OK|Changed=" & CStr(changedCount) & "|" & JoinStringArrayAddins(fileStatuses, "|")
    LogDiagnosticEvent "ADDINS-PUBLISH", "Publish succeeded|" & report
    GoTo CleanExit

RollbackChanges:
    RollBackPublishedAddins addinsRoot, backupRoot, changed, hadTarget, addinNames
    If FileExistsAddins(manifestBackupPath) Then
        CopyFileVerifiedAddins manifestBackupPath, manifestPath
    Else
        DeleteFileIfPresentAddins manifestPath
    End If
    PublishAddins = False
    If Len(report) = 0 Then report = "PublishAddins failed during commit."
    GoTo CleanExit

FailSoft:
    PublishAddins = False
    If Len(report) = 0 Then report = "PublishAddins failed."
    LogDiagnosticEvent "ADDINS-PUBLISH", report
    GoTo CleanExit

FailPublish:
    report = "PublishAddins failed at " & currentStep & ": " & Err.Description
    PublishAddins = False
    LogDiagnosticEvent "ADDINS-PUBLISH", report

CleanExit:
    mLastAddinsPublishReport = report
    DeleteFolderRecursiveAddins stageRoot
    DeleteFolderRecursiveAddins backupRoot
End Function

Public Function GetAddinsManifest() As String
    Dim addinsRoot As String
    Dim addinNames As Variant
    Dim i As Long
    Dim targetPath As String
    Dim lines() As String
    Dim idx As Long

    addinsRoot = ResolveAddinsRootPath()
    If addinsRoot = "" Then Exit Function

    addinNames = GetRequiredAddinNames()
    ReDim lines(0 To (UBound(addinNames) - LBound(addinNames) + 1) + 4)

    idx = 0
    lines(idx) = "{": idx = idx + 1
    lines(idx) = "  ""published_utc"": """ & EscapeJsonAddins(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")) & """,": idx = idx + 1
    lines(idx) = "  ""files"": [": idx = idx + 1
    For i = LBound(addinNames) To UBound(addinNames)
        targetPath = addinsRoot & CStr(addinNames(i))
        lines(idx) = "    { ""name"": """ & EscapeJsonAddins(CStr(addinNames(i))) & """, ""size_bytes"": " & CStr(SafeFileLenAddins(targetPath)) & " }" & IIf(i < UBound(addinNames), ",", "")
        idx = idx + 1
    Next i
    lines(idx) = "  ]": idx = idx + 1
    lines(idx) = "}": idx = idx + 1

    ReDim Preserve lines(0 To idx - 1)
    GetAddinsManifest = Join(lines, vbCrLf)
End Function

Public Function GetLastAddinsPublishReport() As String
    GetLastAddinsPublishReport = mLastAddinsPublishReport
End Function

Public Sub SetAddinsPublishSharePointRootOverride(ByVal rootPath As String)
    mSharePointRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearAddinsPublishSharePointRootOverride()
    mSharePointRootOverride = vbNullString
End Sub

Private Function GetRequiredAddinNames() As Variant
    GetRequiredAddinNames = Array( _
        "invSys.Core.xlam", _
        "invSys.Inventory.Domain.xlam", _
        "invSys.Designs.Domain.xlam", _
        "invSys.Operations.xlam", _
        "invSys.Admin.xlam")
End Function

Private Function ResolveAddinsRootPath() As String
    Dim sharePointRoot As String

    sharePointRoot = NormalizeFolderPathAddins(mSharePointRootOverride)
    If sharePointRoot = "" Then sharePointRoot = NormalizeFolderPathAddins(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot = "" Then Exit Function

    ResolveAddinsRootPath = sharePointRoot & "Addins\"
End Function

Private Function NormalizeFolderPathAddins(ByVal folderPath As String) As String
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathAddins = folderPath
End Function

Private Function FileExistsAddins(ByVal filePath As String) As Boolean
    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function
    FileExistsAddins = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Function FolderExistsAddins(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" And Len(folderPath) > 3 Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    FolderExistsAddins = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function SafeFileLenAddins(ByVal filePath As String) As Long
    On Error Resume Next
    SafeFileLenAddins = FileLen(filePath)
    On Error GoTo 0
End Function

Private Sub EnsureFolderRecursiveAddins(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    Dim slashPos As Long

    folderPath = NormalizeFolderPathAddins(folderPath)
    If folderPath = "" Then Exit Sub
    folderPath = Left$(folderPath, Len(folderPath) - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then GoTo CleanExit

    slashPos = InStrRev(folderPath, "\")
    If slashPos > 3 Then
        parentPath = Left$(folderPath, slashPos - 1)
        If Not fso.FolderExists(parentPath) Then EnsureFolderRecursiveAddins parentPath
    End If

    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

CleanExit:
    Set fso = Nothing
End Sub

Private Sub DeleteFolderRecursiveAddins(ByVal folderPath As String)
    Dim itemName As String
    Dim childPath As String

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Right$(folderPath, 1) = "\" And Len(folderPath) > 3 Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    itemName = Dir$(folderPath & "\*", vbNormal Or vbHidden Or vbSystem Or vbDirectory)
    Do While itemName <> ""
        If itemName <> "." And itemName <> ".." Then
            childPath = folderPath & "\" & itemName
            If (GetAttr(childPath) And vbDirectory) = vbDirectory Then
                DeleteFolderRecursiveAddins childPath
            Else
                On Error Resume Next
                Kill childPath
                On Error GoTo 0
            End If
        End If
        itemName = Dir$
    Loop

    On Error Resume Next
    RmDir folderPath
    On Error GoTo 0
End Sub

Private Sub DeleteFileIfPresentAddins(ByVal filePath As String)
    On Error Resume Next
    If FileExistsAddins(filePath) Then Kill filePath
    On Error GoTo 0
End Sub

Private Sub CopyFileVerifiedAddins(ByVal sourcePath As String, ByVal targetPath As String)
    Dim parentPath As String

    parentPath = ParentFolderAddins(targetPath)
    If parentPath <> "" Then EnsureFolderRecursiveAddins parentPath
    DeleteFileIfPresentAddins targetPath
    FileCopy sourcePath, targetPath
End Sub

Private Function ParentFolderAddins(ByVal targetPath As String) As String
    Dim slashPos As Long

    targetPath = Trim$(Replace$(targetPath, "/", "\"))
    slashPos = InStrRev(targetPath, "\")
    If slashPos <= 0 Then Exit Function
    ParentFolderAddins = Left$(targetPath, slashPos - 1)
End Function

Private Function BuildTokenAddins() As String
    Randomize
    BuildTokenAddins = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 100000), "00000")
End Function

Private Function BuildWorkingFolderAddins(ByVal folderRole As String) As String
    Dim tempRoot As String

    tempRoot = NormalizeFolderPathAddins(Environ$("TEMP"))
    If tempRoot = "" Then tempRoot = "C:\Temp\"
    BuildWorkingFolderAddins = tempRoot & "invSys_addins_publish\" & BuildTokenAddins() & "\" & Trim$(folderRole) & "\"
End Function

Private Sub RollBackPublishedAddins(ByVal addinsRoot As String, _
                                    ByRef backupRoot As String, _
                                    ByRef changed() As Boolean, _
                                    ByRef hadTarget() As Boolean, _
                                    ByVal addinNames As Variant)
    Dim i As Long
    Dim targetPath As String
    Dim backupPath As String

    For i = LBound(addinNames) To UBound(addinNames)
        If changed(i) Then
            targetPath = addinsRoot & CStr(addinNames(i))
            backupPath = backupRoot & CStr(addinNames(i))
            If hadTarget(i) And FileExistsAddins(backupPath) Then
                CopyFileVerifiedAddins backupPath, targetPath
            Else
                DeleteFileIfPresentAddins targetPath
            End If
        End If
    Next i
End Sub

Private Function WriteTextFileAddins(ByVal filePath As String, ByVal textOut As String) As Boolean
    Dim fileNum As Integer
    Dim parentPath As String

    On Error GoTo FailWrite

    parentPath = ParentFolderAddins(filePath)
    If parentPath <> "" Then EnsureFolderRecursiveAddins parentPath

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textOut;
    Close #fileNum
    WriteTextFileAddins = True
    Exit Function

FailWrite:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
End Function

Private Function EscapeJsonAddins(ByVal textIn As String) As String
    textIn = Replace$(textIn, "\", "\\")
    textIn = Replace$(textIn, """", "\""")
    textIn = Replace$(textIn, vbCrLf, "\n")
    textIn = Replace$(textIn, vbCr, "\n")
    textIn = Replace$(textIn, vbLf, "\n")
    EscapeJsonAddins = textIn
End Function

Private Function JoinStringArrayAddins(ByRef values() As String, ByVal delimiter As String) As String
    Dim i As Long
    Dim outputText As String

    For i = LBound(values) To UBound(values)
        If Len(values(i)) > 0 Then
            If Len(outputText) > 0 Then outputText = outputText & delimiter
            outputText = outputText & values(i)
        End If
    Next i

    JoinStringArrayAddins = outputText
End Function
