Attribute VB_Name = "modTesterFiles"
Option Explicit
Option Private Module

Public Const TESTER_DEFAULT_WAREHOUSE_ID As String = "WH1"
Public Const TESTER_DEFAULT_STATION_ID As String = "R1"
Public Const TESTER_DEFAULT_OPERATOR_SUFFIX As String = ".Receiving.Operator.xlsm"
Public Const TESTER_SEED_SKU As String = "TEST-SKU-001"
Public Const TESTER_SEED_DESCRIPTION As String = "Test SKU for Confirm Writes"
Public Const TESTER_SEED_LOCATION As String = "A1"
Public Const TESTER_SEED_QTY As Double = 100#

Public Sub RemoveNonReceivingOperatorSheetsTesterSetup(ByVal wb As Workbook)
    Dim keepSheets As Object
    Dim i As Long
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    Set keepSheets = CreateObject("Scripting.Dictionary")
    keepSheets.CompareMode = vbTextCompare
    keepSheets("ReceivedTally") = True
    keepSheets("InventoryManagement") = True
    keepSheets("ReceivedLog") = True

    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If Not keepSheets.Exists(ws.Name) Then
            If wb.Worksheets.Count > 1 Then ws.Delete
        End If
    Next i
End Sub

Public Function ValidateTesterSetupSpec(ByRef spec As TesterProvisionSpec, _
                                         ByRef report As String) As Boolean
    If spec.UserId = "" Then
        report = "UserId is required."
        Exit Function
    End If
    If spec.PinHash = "" Then
        report = "PinHash is required."
        Exit Function
    End If
    If spec.WarehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If
    If spec.StationId = "" Then
        report = "StationId is required."
        Exit Function
    End If
    If spec.PathLocal = "" Then
        report = "PathLocal is required."
        Exit Function
    End If

    ValidateTesterSetupSpec = True
    report = "OK"
End Function

Public Sub NormalizeTesterSetupSpec(ByRef spec As TesterProvisionSpec)
    spec.UserId = Trim$(spec.UserId)
    spec.PinHash = Trim$(spec.PinHash)
    spec.WarehouseId = Trim$(spec.WarehouseId)
    spec.StationId = Trim$(spec.StationId)
    spec.PathLocal = Trim$(Replace$(spec.PathLocal, "/", "\"))
    spec.PathSharePointRoot = Trim$(Replace$(spec.PathSharePointRoot, "/", "\"))
End Sub

Public Sub ApplyTesterSetupDefaults(ByRef spec As TesterProvisionSpec)
    If spec.WarehouseId = "" Then spec.WarehouseId = TESTER_DEFAULT_WAREHOUSE_ID
    If spec.StationId = "" Then spec.StationId = TESTER_DEFAULT_STATION_ID
    If spec.PathLocal = "" Then spec.PathLocal = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(spec.WarehouseId, False)
    If spec.PathSharePointRoot = "" Then spec.PathSharePointRoot = DetectSharePointRoot(spec.WarehouseId)
    spec.PathLocal = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    spec.PathSharePointRoot = NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
End Sub

Public Function BuildReceivingOperatorPathTesterSetup(ByRef spec As TesterProvisionSpec) As String
    Dim rootPath As String

    rootPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    If rootPath = "" Then Exit Function
    BuildReceivingOperatorPathTesterSetup = rootPath & "\" & spec.WarehouseId & TESTER_DEFAULT_OPERATOR_SUFFIX
End Function

Public Function RuntimeArtifactsExistTesterSetup(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function

    RuntimeArtifactsExistTesterSetup = _
        FileExistsTesterSetup(rootPath & "\" & warehouseId & ".invSys.Config.xlsb") And _
        FileExistsTesterSetup(rootPath & "\" & warehouseId & ".invSys.Auth.xlsb") And _
        FileExistsTesterSetup(rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb")
End Function

Public Function IsUncPathTesterSetup(ByVal pathIn As String) As Boolean
    pathIn = NormalizeFolderPathTesterSetup(pathIn, False)
    IsUncPathTesterSetup = (Left$(pathIn, 2) = "\\")
End Function

Public Function CandidateLooksLikeSharePointRootTesterSetup(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    If rootPath = "" Then Exit Function
    If Not IsUsableLocalPathTesterSetup(rootPath) Then Exit Function

    If FolderExistsTesterSetup(rootPath & "\Addins") Then
        If FileExistsTesterSetup(rootPath & "\Addins\invSys.Core.xlam") _
           And FileExistsTesterSetup(rootPath & "\Addins\invSys.Operations.xlam") Then
            CandidateLooksLikeSharePointRootTesterSetup = True
            Exit Function
        End If
    End If

    If warehouseId <> "" Then
        If FileExistsTesterSetup(rootPath & "\TesterPackage\" & warehouseId & "\" & warehouseId & ".TesterBundle.zip") Then
            CandidateLooksLikeSharePointRootTesterSetup = True
        End If
    End If
End Function

Public Sub AddUniquePathTesterSetup(ByVal items As Collection, ByVal pathIn As String)
    Dim candidate As Variant
    pathIn = NormalizeFolderPathTesterSetup(pathIn, False)
    If pathIn = "" Then Exit Sub
    For Each candidate In items
        If StrComp(CStr(candidate), pathIn, vbTextCompare) = 0 Then Exit Sub
    Next candidate
    items.Add pathIn
End Sub

Public Function EnsureTemporaryAdminCapabilityTesterSetup(ByRef spec As TesterProvisionSpec, _
                                                           ByRef tempGranted As Boolean, _
                                                           ByRef report As String) As Boolean
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userRow As Long
    Dim capRow As Long
    Dim openedTransient As Boolean

    On Error GoTo FailEnsure

    authPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    Set wbAuth = OpenWorkbookForWriteTesterSetup(authPath, openedTransient, report)
    If wbAuth Is Nothing Then GoTo FailSoft
    If Not modAuth.EnsureAuthSchema(wbAuth, spec.WarehouseId, "svc_processor", report) Then GoTo FailSoft

    Set loUsers = FindTableByNameTesterSetup(wbAuth, "tblUsers")
    Set loCaps = FindTableByNameTesterSetup(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        report = "Auth tables were not available for temporary admin grant."
        GoTo FailSoft
    End If

    userRow = EnsureUserRowTesterSetup(loUsers, spec.UserId)
    SetTableCellTesterSetup loUsers, userRow, "UserId", spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "DisplayName", spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "Status", "Active"
    If SafeTrimTesterSetup(GetTableCellTesterSetup(loUsers, userRow, "PinHash")) = "" Then
        SetTableCellTesterSetup loUsers, userRow, "PinHash", spec.PinHash
    End If

    capRow = EnsureCapabilityRowTesterSetup(loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId)
    SetTableCellTesterSetup loCaps, capRow, "Status", "ACTIVE"
    wbAuth.Save

    tempGranted = True
    EnsureTemporaryAdminCapabilityTesterSetup = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    EnsureTemporaryAdminCapabilityTesterSetup = False
    If Len(report) = 0 Then report = "EnsureTemporaryAdminCapabilityTesterSetup failed."
    GoTo CleanExit

FailEnsure:
    report = "EnsureTemporaryAdminCapabilityTesterSetup failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbAuth, openedTransient
End Function

Public Sub RevokeAdminCapabilityTesterSetup(ByRef spec As TesterProvisionSpec)
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loCaps As ListObject
    Dim openedTransient As Boolean
    Dim report As String

    On Error Resume Next
    If Not modTesterCommands.CheckPhase() Then Exit Sub
    authPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    Set wbAuth = OpenWorkbookForWriteTesterSetup(authPath, openedTransient, report)
    If wbAuth Is Nothing Then GoTo CleanExit
    Set loCaps = FindTableByNameTesterSetup(wbAuth, "tblCapabilities")
    If loCaps Is Nothing Then GoTo CleanExit
    DeactivateCapabilityTesterSetup loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId
    wbAuth.Save
CleanExit:
    CloseWorkbookIfTransientTesterSetup wbAuth, openedTransient
    On Error GoTo 0
End Sub

Public Function TesterSeedAlreadyPresentTesterSetup(ByVal inventoryWb As Workbook) As Boolean
    Dim loSku As ListObject
    Dim rowIndex As Long
    Dim qtyVal As Variant
    Dim loInv As ListObject

    If inventoryWb Is Nothing Then Exit Function

    Set loSku = FindTableByNameTesterSetup(inventoryWb, "tblSkuBalance")
    If Not loSku Is Nothing Then
        rowIndex = FindRowByValueTesterSetup(loSku, "SKU", TESTER_SEED_SKU)
        If rowIndex > 0 Then
            qtyVal = GetTableCellTesterSetup(loSku, rowIndex, "QtyOnHand")
            If IsNumeric(qtyVal) Then
                TesterSeedAlreadyPresentTesterSetup = (CDbl(qtyVal) > 0)
                If TesterSeedAlreadyPresentTesterSetup Then Exit Function
            Else
                TesterSeedAlreadyPresentTesterSetup = True
                Exit Function
            End If
        End If
    End If

    Set loInv = FindTableByNameTesterSetup(inventoryWb, "invSys")
    If Not loInv Is Nothing Then
        rowIndex = FindRowByValueTesterSetup(loInv, "ITEM_CODE", TESTER_SEED_SKU)
        If rowIndex > 0 Then
            qtyVal = GetTableCellTesterSetup(loInv, rowIndex, "TOTAL INV")
            If IsNumeric(qtyVal) Then TesterSeedAlreadyPresentTesterSetup = (CDbl(qtyVal) > 0)
        End If
    End If
End Function

Public Function ReceivingWorkbookHasCanonicalSurfacesTesterSetup(ByVal wb As Workbook) As Boolean
    ReceivingWorkbookHasCanonicalSurfacesTesterSetup = _
        WorksheetHasTableTesterSetup(wb, "ReceivedTally", "ReceivedTally") And _
        WorksheetHasTableTesterSetup(wb, "InventoryManagement", "invSys") And _
        WorksheetHasTableTesterSetup(wb, "ReceivedLog", "ReceivedLog")
End Function

Public Function ReceivingWorkbookHasAliasSurfacesTesterSetup(ByVal wb As Workbook) As Boolean
    ReceivingWorkbookHasAliasSurfacesTesterSetup = _
        WorksheetExistsTesterSetup(wb, "tblReceiving") And _
        WorksheetExistsTesterSetup(wb, "tblReadModel") And _
        WorksheetExistsTesterSetup(wb, "tblStatus")
End Function

Public Function WorksheetHasTableTesterSetup(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    WorksheetHasTableTesterSetup = Not (ws.ListObjects(tableName) Is Nothing)
    On Error GoTo 0
End Function

Public Function WorksheetExistsTesterSetup(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExistsTesterSetup = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Function OpenWorkbookForWriteTesterSetup(ByVal workbookPath As String, _
                                                 ByRef openedTransient As Boolean, _
                                                 ByRef report As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then
        report = "Workbook path is required."
        Exit Function
    End If

    Set wb = FindOpenWorkbookByPathTesterSetup(workbookPath)
    If wb Is Nothing Then
        If Not FileExistsTesterSetup(workbookPath) Then
            report = "Workbook not found: " & workbookPath
            Exit Function
        End If
        Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wb Is Nothing
        If openedTransient Then modTesterCommands.RegisterOwned wb
    End If
    If wb Is Nothing Then
        report = "Workbook could not be opened: " & workbookPath
        Exit Function
    End If
    If wb.ReadOnly Then
        report = "Workbook is read-only: " & workbookPath
        If openedTransient Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
        End If
        Exit Function
    End If

    Set OpenWorkbookForWriteTesterSetup = wb
End Function

Public Sub CloseWorkbookIfTransientTesterSetup(ByVal wb As Workbook, ByVal openedTransient As Boolean)
    If Not openedTransient Then Exit Sub
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    If Not wb.ReadOnly Then
        If wb.Saved = False And modTesterCommands.CheckPhase() Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Public Function FindOpenWorkbookByPathTesterSetup(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathTesterSetup = wb
            Exit Function
        End If
    Next wb
End Function

Public Function FindTableByNameTesterSetup(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindTableByNameTesterSetup = ws.ListObjects(tableName)
        If Not FindTableByNameTesterSetup Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Public Function EnsureUserRowTesterSetup(ByVal lo As ListObject, ByVal userId As String) As Long
    EnsureUserRowTesterSetup = FindRowByValueTesterSetup(lo, "UserId", userId)
    If EnsureUserRowTesterSetup > 0 Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureUserRowTesterSetup = 1
    ElseIf lo.ListRows.Count = 1 And SafeTrimTesterSetup(GetTableCellTesterSetup(lo, 1, "UserId")) = "" Then
        EnsureUserRowTesterSetup = 1
    Else
        EnsureUserRowTesterSetup = lo.ListRows.Add.Index
    End If
End Function

Public Function EnsureCapabilityRowTesterSetup(ByVal lo As ListObject, _
                                                ByVal userId As String, _
                                                ByVal capability As String, _
                                                ByVal warehouseId As String, _
                                                ByVal stationId As String) As Long
    EnsureCapabilityRowTesterSetup = FindCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If EnsureCapabilityRowTesterSetup > 0 Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureCapabilityRowTesterSetup = 1
    ElseIf lo.ListRows.Count = 1 And SafeTrimTesterSetup(GetTableCellTesterSetup(lo, 1, "UserId")) = "" _
        And SafeTrimTesterSetup(GetTableCellTesterSetup(lo, 1, "Capability")) = "" Then
        EnsureCapabilityRowTesterSetup = 1
    Else
        EnsureCapabilityRowTesterSetup = lo.ListRows.Add.Index
    End If

    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "UserId", userId
    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "Capability", capability
    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "WarehouseId", warehouseId
    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "StationId", stationId
End Function

Public Sub EnsureCapabilityActiveTesterSetup(ByVal lo As ListObject, _
                                              ByVal userId As String, _
                                              ByVal capability As String, _
                                              ByVal warehouseId As String, _
                                              ByVal stationId As String)
    Dim rowIndex As Long

    rowIndex = EnsureCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Sub
    SetTableCellTesterSetup lo, rowIndex, "Status", "ACTIVE"
End Sub

Public Sub DeactivateCapabilityTesterSetup(ByVal lo As ListObject, _
                                            ByVal userId As String, _
                                            ByVal capability As String, _
                                            ByVal warehouseId As String, _
                                            ByVal stationId As String)
    Dim rowIndex As Long

    rowIndex = FindCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Sub
    SetTableCellTesterSetup lo, rowIndex, "Status", "INACTIVE"
End Sub

Public Function CapabilityIsActiveTesterSetup(ByVal lo As ListObject, _
                                               ByVal userId As String, _
                                               ByVal capability As String, _
                                               ByVal warehouseId As String, _
                                               ByVal stationId As String) As Boolean
    Dim rowIndex As Long

    rowIndex = FindCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Function
    CapabilityIsActiveTesterSetup = (StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, rowIndex, "Status")), "ACTIVE", vbTextCompare) = 0)
End Function

Public Function FindCapabilityRowTesterSetup(ByVal lo As ListObject, _
                                              ByVal userId As String, _
                                              ByVal capability As String, _
                                              ByVal warehouseId As String, _
                                              ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "UserId")), userId, vbTextCompare) = 0 _
           And StrComp(UCase$(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "Capability"))), UCase$(capability), vbTextCompare) = 0 _
           And StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 _
           And StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "StationId")), stationId, vbTextCompare) = 0 Then
            FindCapabilityRowTesterSetup = i
            Exit Function
        End If
    Next i
End Function

Public Function FindRowByValueTesterSetup(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long
    Dim columnIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    columnIndex = GetColumnIndexTesterSetup(lo, columnName)
    If columnIndex = 0 Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimTesterSetup(lo.DataBodyRange.Cells(i, columnIndex).Value), expectedValue, vbTextCompare) = 0 Then
            FindRowByValueTesterSetup = i
            Exit Function
        End If
    Next i
End Function

Public Function GetColumnIndexTesterSetup(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexTesterSetup = i
            Exit Function
        End If
    Next i
End Function

Public Function GetTableCellTesterSetup(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim columnIndex As Long

    columnIndex = GetColumnIndexTesterSetup(lo, columnName)
    If columnIndex = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    GetTableCellTesterSetup = lo.DataBodyRange.Cells(rowIndex, columnIndex).Value
End Function

Public Sub SetTableCellTesterSetup(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim columnIndex As Long

    columnIndex = GetColumnIndexTesterSetup(lo, columnName)
    If columnIndex = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, columnIndex).Value = valueOut
End Sub

Public Function NormalizeFolderPathTesterSetup(ByVal folderPath As String, Optional ByVal withTrailingSlash As Boolean = False) As String
    NormalizeFolderPathTesterSetup = modConfig.NormalizeFolderPathForRuntime(folderPath, withTrailingSlash)
End Function

Public Function FolderExistsTesterSetup(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Function
    If Not IsUsableLocalPathTesterSetup(folderPath) Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsTesterSetup = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsTesterSetup = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Public Function FileExistsTesterSetup(ByVal filePath As String) As Boolean
    Dim fso As Object

    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function
    If Not IsUsableLocalPathTesterSetup(filePath) Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsTesterSetup = fso.FileExists(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsTesterSetup = (Len(Dir$(filePath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Public Function RequiredTesterAddinNamesTesterSetup() As Variant
    RequiredTesterAddinNamesTesterSetup = Array( _
        "invSys.Core.xlam", _
        "invSys.Inventory.Domain.xlam", _
        "invSys.Designs.Domain.xlam", _
        "invSys.Operations.xlam", _
        "invSys.Admin.xlam")
End Function

Public Function ResolveCurrentAddinsRootTesterSetup() As String
    Dim candidate As String

    candidate = NormalizeFolderPathTesterSetup(ThisWorkbook.Path, False)
    If CandidateHasRequiredTesterAddinsTesterSetup(candidate) Then
        ResolveCurrentAddinsRootTesterSetup = candidate
        Exit Function
    End If

    candidate = NormalizeFolderPathTesterSetup(Environ$("USERPROFILE") & "\source\repos\invSys_fork\deploy\current", False)
    If CandidateHasRequiredTesterAddinsTesterSetup(candidate) Then
        ResolveCurrentAddinsRootTesterSetup = candidate
        Exit Function
    End If

    candidate = NormalizeFolderPathTesterSetup(modConfig.GetString("PathSharePointRoot", ""), False)
    If CandidateHasRequiredTesterAddinsTesterSetup(candidate & "\Addins") Then
        ResolveCurrentAddinsRootTesterSetup = NormalizeFolderPathTesterSetup(candidate & "\Addins", False)
    End If
End Function

Public Function CandidateHasRequiredTesterAddinsTesterSetup(ByVal folderPath As String) As Boolean
    Dim addinName As Variant
    Dim requiredAddins As Variant

    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Function
    If Not FolderExistsTesterSetup(folderPath) Then Exit Function

    requiredAddins = RequiredTesterAddinNamesTesterSetup()
    For Each addinName In requiredAddins
        If Not FileExistsTesterSetup(folderPath & "\" & CStr(addinName)) Then Exit Function
    Next addinName

    CandidateHasRequiredTesterAddinsTesterSetup = True
End Function

Public Function CopyFileIfNeededTesterSetup(ByVal sourcePath As String, ByVal targetPath As String) As Boolean
    sourcePath = Trim$(Replace$(sourcePath, "/", "\"))
    targetPath = Trim$(Replace$(targetPath, "/", "\"))
    If sourcePath = "" Or targetPath = "" Then Exit Function
    If Not FileExistsTesterSetup(sourcePath) Then Exit Function
    If FileExistsTesterSetup(targetPath) Then
        If SafeFileLenTesterSetup(sourcePath) = SafeFileLenTesterSetup(targetPath) Then Exit Function
        Kill targetPath
    End If
    EnsureFolderRecursiveTesterSetup GetParentFolderTesterSetup(targetPath)
    FileCopy sourcePath, targetPath
    CopyFileIfNeededTesterSetup = True
End Function

Public Function SafeFileLenTesterSetup(ByVal filePath As String) As Long
    On Error Resume Next
    SafeFileLenTesterSetup = FileLen(filePath)
    On Error GoTo 0
End Function

Public Sub EnsureFolderRecursiveTesterSetup(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Sub
    If FolderExistsTesterSetup(folderPath) Then Exit Sub

    parentPath = GetParentFolderTesterSetup(folderPath)
    If parentPath <> "" And Not FolderExistsTesterSetup(parentPath) Then EnsureFolderRecursiveTesterSetup parentPath
    MkDir folderPath
End Sub

Public Function GetParentFolderTesterSetup(ByVal pathIn As String) As String
    Dim slashPos As Long

    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    slashPos = InStrRev(pathIn, "\")
    If slashPos = 3 And Mid$(pathIn, 2, 2) = ":\" Then
        GetParentFolderTesterSetup = Left$(pathIn, 3)
    ElseIf slashPos > 1 Then
        GetParentFolderTesterSetup = Left$(pathIn, slashPos - 1)
    End If
End Function

Public Function IsUsableLocalPathTesterSetup(ByVal pathIn As String) As Boolean
    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    If InStr(1, pathIn, "://", vbTextCompare) > 0 Then Exit Function
    If Left$(pathIn, 8) = "https:\\" Then Exit Function
    If Left$(pathIn, 7) = "http:\\" Then Exit Function
    If InStr(pathIn, "*") > 0 Or InStr(pathIn, "?") > 0 Then Exit Function
    IsUsableLocalPathTesterSetup = True
End Function

Public Function SafeTrimTesterSetup(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimTesterSetup = Trim$(CStr(valueIn))
    On Error GoTo 0
End Function

Public Sub RestoreCoreRootOverrideTesterSetup(ByVal priorRootOverride As String)
    If Trim$(priorRootOverride) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRootOverride
    End If
End Sub
