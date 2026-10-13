Attribute VB_Name = "test_CreateWarehouse"
Option Explicit

Private mCheckNames() As String
Private mCheckResults() As String
Private mCheckDetails() As String
Private mCheckCount As Long

Private mWarehouseId As String
Private mStationId As String
Private mLocalRoot As String
Private mSharePointRoot As String
Private mOperatorRoot As String
Private mSummary As String

Public Function TestCreateWarehouse_EndToEndLifecycle() As Long
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim duplicateSpec As modWarehouseBootstrap.WarehouseSpec
    Dim templateRoot As String
    Dim duplicateRoot As String
    Dim validSpec As Boolean
    Dim existsBefore As Boolean
    Dim bootstrapOk As Boolean
    Dim publishOk As Boolean
    Dim duplicateExists As Boolean
    Dim duplicateRejected As Boolean
    Dim detail As String
    Dim duplicateReport As String

    On Error GoTo FailTest

    ResetCreateWarehouseEvidence

    mWarehouseId = "WHBOOT-E2E_01"
    mStationId = "ADM1"
    mLocalRoot = BuildCreateWarehouseTempRoot("local")
    mSharePointRoot = BuildCreateWarehouseTempRoot("share")
    templateRoot = BuildCreateWarehouseTempRoot("templates")
    duplicateRoot = BuildCreateWarehouseTempRoot("duplicate")

    spec.WarehouseId = mWarehouseId
    spec.WarehouseName = "Create Warehouse Integration"
    spec.StationId = mStationId
    spec.AdminUser = "admin.integration"
    spec.PathLocal = mLocalRoot
    spec.PathSharePoint = mSharePointRoot

    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot

    validSpec = modWarehouseBootstrap.ValidateWarehouseSpec(spec, detail)
    RecordCreateWarehouseCheck "WarehouseSpec.Valid", validSpec, detail
    If Not validSpec Then GoTo CleanExit

    existsBefore = modWarehouseBootstrap.WarehouseIdExists(spec.WarehouseId)
    RecordCreateWarehouseCheck "CollisionCheck.InitialClear", Not existsBefore, _
        "WarehouseIdExists=" & CStr(existsBefore)
    If existsBefore Then GoTo CleanExit

    bootstrapOk = modWarehouseBootstrap.BootstrapWarehouseLocal(spec)
    RecordCreateWarehouseCheck "Bootstrap.Local", bootstrapOk, modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    If Not bootstrapOk Then GoTo CleanExit

    ' Headless follow-up work binds its generated runtime explicitly. Bootstrap
    ' must not select a different interactive warehouse or sign out its caller.
    modRuntimeWorkbooks.SetCoreDataRootOverride spec.PathLocal

    RecordCreateWarehouseCheck "LocalStructure.Exists", _
        AssertLocalStructureCreateWarehouse(spec, detail), detail

    If Not RunD14CreateWarehouseChecks(spec) Then GoTo CleanExit

    RecordCreateWarehouseCheck "ConfigSeeded.Correctly", _
        AssertConfigSeededCreateWarehouse(spec, detail), detail

    publishOk = modWarehouseBootstrap.PublishInitialArtifacts(spec)
    RecordCreateWarehouseCheck "SharePointPublish.Initial", publishOk, modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    If Not publishOk Then GoTo CleanExit

    RecordCreateWarehouseCheck "SharePointArtifacts.Exists", _
        AssertSharePointArtifactsCreateWarehouse(spec, detail), detail

    duplicateExists = modWarehouseBootstrap.WarehouseIdExists(spec.WarehouseId)
    RecordCreateWarehouseCheck "CollisionCheck.DuplicateVisible", duplicateExists, _
        "WarehouseIdExists=" & CStr(duplicateExists)

    duplicateSpec = spec
    duplicateSpec.PathLocal = duplicateRoot
    duplicateRejected = Not modWarehouseBootstrap.BootstrapWarehouseLocal(duplicateSpec)
    duplicateReport = modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    RecordCreateWarehouseCheck "DuplicateRun.Rejected", _
        duplicateRejected And InStr(1, duplicateReport, "already exists", vbTextCompare) > 0, _
        duplicateReport

    If AllCreateWarehouseChecksPassed() Then
        mSummary = "Create warehouse lifecycle completed, SharePoint artifacts were published, and duplicate rejection was proven."
        TestCreateWarehouse_EndToEndLifecycle = 1
    Else
        mSummary = "One or more create warehouse lifecycle checks failed."
    End If

CleanExit:
    On Error Resume Next
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteCreateWarehouseFolderRecursive duplicateRoot
    DeleteCreateWarehouseFolderRecursive mOperatorRoot
    DeleteCreateWarehouseFolderRecursive mSharePointRoot
    DeleteCreateWarehouseFolderRecursive mLocalRoot
    DeleteCreateWarehouseFolderRecursive templateRoot
    On Error GoTo 0
    If mSummary = "" Then mSummary = "Create warehouse lifecycle did not complete."
    Exit Function

FailTest:
    RecordCreateWarehouseCheck "TestHarness.Exception", False, Err.Description
    mSummary = "Create warehouse lifecycle raised an unexpected exception."
    Resume CleanExit
End Function

Private Function RunD14CreateWarehouseChecks(ByRef spec As modWarehouseBootstrap.WarehouseSpec) As Boolean
    Dim wbInventory As Workbook
    Dim wbSnapshot As Workbook
    Dim wbOperator As Workbook
    Dim loEntities As ListObject
    Dim loSnapshot As ListObject
    Dim loOperator As ListObject
    Dim inventoryKeys As Object
    Dim snapshotKeys As Object
    Dim operatorKeys As Object
    Dim detail As String
    Dim headersOk As Boolean
    Dim rowsOk As Boolean
    Dim roundTripOk As Boolean
    Dim customOk As Boolean
    Dim operatorPath As String

    On Error GoTo FailCheck

    operatorPath = modWarehouseBootstrap.GetLastWarehouseOperatorWorkbookPath()
    Set wbInventory = Application.Workbooks.Open( _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb")
    Set wbSnapshot = Application.Workbooks.Open( _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb")
    Set wbOperator = Application.Workbooks.Open(operatorPath)

    Set loEntities = FindTableCreateWarehouse(wbInventory, "tblInventoryEntities")
    Set loSnapshot = FindTableCreateWarehouse(wbSnapshot, "tblInventorySnapshot")
    Set loOperator = FindTableCreateWarehouse(wbOperator, "invSys")

    headersOk = Not loEntities Is Nothing And Not loSnapshot Is Nothing And Not loOperator Is Nothing
    If Not headersOk Then
        detail = "Missing table(s): entities=" & CStr(Not loEntities Is Nothing) & _
                 "; snapshot=" & CStr(Not loSnapshot Is Nothing) & _
                 "; operator=" & CStr(Not loOperator Is Nothing)
    ElseIf Not TableHasHeadersCreateWarehouse(loEntities, _
        Array("System_Key", "SKU", "QtyOnHand", "Location", "Condition")) Then
        headersOk = False
        detail = "tblInventoryEntities is missing a required managed header."
    ElseIf Not TableHasHeadersCreateWarehouse(loSnapshot, _
        Array("System_Key", "SKU", "QtyOnHand", "Location", "Condition")) Then
        headersOk = False
        detail = "tblInventorySnapshot is missing a required managed header."
    ElseIf Not TableHasHeadersCreateWarehouse(loOperator, _
        Array("System_Key", "SKU", "QtyOnHand", "Location", "Condition", _
              "LastRefreshUTC", "SnapshotId", "SourceType", "IsStale")) Then
        headersOk = False
        detail = "operator invSys is missing a required managed header."
    ElseIf Not WorkbookTablesExcludeHeaderCreateWarehouse(wbInventory, "ROW") Then
        headersOk = False
        detail = "Inventory workbook still contains ROW in " & _
                 FirstTableWithHeaderCreateWarehouse(wbInventory, "ROW") & "."
    ElseIf Not WorkbookTablesExcludeHeaderCreateWarehouse(wbSnapshot, "ROW") Then
        headersOk = False
        detail = "Snapshot workbook still contains a ROW table header."
    ElseIf Not WorkbookTablesExcludeHeaderCreateWarehouse(wbOperator, "ROW") Then
        headersOk = False
        detail = "Operator workbook still contains a ROW table header."
    Else
        detail = "Inventory, snapshot, and operator tables contain required managed headers and no ROW header."
    End If
    RecordCreateWarehouseCheck "D14.GeneratedSchemas.ManagedHeadersNoROW", headersOk, detail
    If Not headersOk Then GoTo CleanExit

    Set inventoryKeys = CollectEntityKeysCreateWarehouse(loEntities, True, detail)
    rowsOk = Not inventoryKeys Is Nothing
    If rowsOk Then rowsOk = inventoryKeys.Count > 0
    RecordCreateWarehouseCheck "D14.Seed.UniqueKeysConditionGood", rowsOk, detail
    If Not rowsOk Then GoTo CleanExit

    Set snapshotKeys = CollectEntityKeysCreateWarehouse(loSnapshot, True, detail)
    Set operatorKeys = CollectEntityKeysCreateWarehouse(loOperator, True, detail)
    roundTripOk = DictionariesHaveSameKeysCreateWarehouse(inventoryKeys, snapshotKeys) _
                  And DictionariesHaveSameKeysCreateWarehouse(inventoryKeys, operatorKeys)
    detail = "Inventory entity keys must survive processor application, snapshot publication, and operator refresh."
    RecordCreateWarehouseCheck "D14.RoundTrip.PreservesSystemKey", roundTripOk, detail
    If Not roundTripOk Then GoTo CleanExit

    rowsOk = AssertRepeatedAdminSeedCreatesUniqueKeysCreateWarehouse( _
        spec, wbInventory, loEntities, detail)
    RecordCreateWarehouseCheck "D14.AdminSeed.RepeatedCreatesUniqueKeysNoMigration", rowsOk, detail
    If Not rowsOk Then GoTo CleanExit

    customOk = AssertOperatorCustomColumnSurvivesRefreshCreateWarehouse( _
        wbOperator, loOperator, spec.WarehouseId, inventoryKeys, detail)
    RecordCreateWarehouseCheck "D14.OperatorRefresh.PreservesCustomColumn", customOk, detail
    If Not customOk Then GoTo CleanExit

    CloseCreateWarehouseWorkbook wbOperator
    Set wbOperator = Application.Workbooks.Open(operatorPath)
    Set loOperator = FindTableCreateWarehouse(wbOperator, "invSys")
    customOk = OperatorCustomValuePresentCreateWarehouse( _
        loOperator, FirstDictionaryKeyCreateWarehouse(inventoryKeys), _
        "Custom_Local_Note", "PRESERVE-ME")
    detail = "Custom_Local_Note remained associated with its System_Key after save and reopen."
    RecordCreateWarehouseCheck "D14.OperatorReopen.PreservesCustomColumn", customOk, detail
    If Not customOk Then GoTo CleanExit

    RunD14CreateWarehouseChecks = True

CleanExit:
    CloseCreateWarehouseWorkbook wbOperator
    CloseCreateWarehouseWorkbook wbSnapshot
    CloseCreateWarehouseWorkbook wbInventory
    Exit Function

FailCheck:
    detail = Err.Description
    RecordCreateWarehouseCheck "D14.TestHarness.Exception", False, detail
    Resume CleanExit
End Function

Private Function AssertRepeatedAdminSeedCreatesUniqueKeysCreateWarehouse( _
    ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
    ByVal wbInventory As Workbook, _
    ByRef loEntities As ListObject, _
    ByRef detail As String) As Boolean

    Dim beforeKeys As Object
    Dim firstKeys As Object
    Dim secondKeys As Object
    Dim seedReport As String

    Set beforeKeys = CollectEntityKeysCreateWarehouse(loEntities, True, detail)
    If beforeKeys Is Nothing Then Exit Function

    If Not modAdminInventorySeed.SeedDemoInventoryForWarehouse( _
        spec.WarehouseId, spec.StationId, spec.AdminUser, seedReport) Then
        detail = seedReport
        Exit Function
    End If
    Set loEntities = FindTableCreateWarehouse(wbInventory, "tblInventoryEntities")
    Set firstKeys = CollectEntityKeysCreateWarehouse(loEntities, True, detail)
    If firstKeys Is Nothing Then Exit Function
    If firstKeys.Count <> 24 Then
        detail = "First Admin seed must ensure the complete 24-entity kit without duplicating bootstrap groups; before=" & _
                 CStr(beforeKeys.Count) & "; after=" & CStr(firstKeys.Count) & "."
        Exit Function
    End If

    If Not modAdminInventorySeed.SeedDemoInventoryForWarehouse( _
        spec.WarehouseId, spec.StationId, spec.AdminUser, seedReport) Then
        detail = seedReport
        Exit Function
    End If
    Set loEntities = FindTableCreateWarehouse(wbInventory, "tblInventoryEntities")
    Set secondKeys = CollectEntityKeysCreateWarehouse(loEntities, True, detail)
    If secondKeys Is Nothing Then Exit Function
    If secondKeys.Count <> firstKeys.Count Then
        detail = "Second Admin seed must be idempotent; first=" & _
                 CStr(firstKeys.Count) & "; second=" & CStr(secondKeys.Count) & "."
        Exit Function
    End If
    If InventoryLogContainsEventTypeCreateWarehouse(wbInventory, "MIGRATION_SEED") Then
        detail = "Supported bootstrap/Admin seed path called legacy MIGRATION_SEED."
        Exit Function
    End If

    detail = "Repeated Admin seed retained the first 24 collision-free System_Key values with Condition=GOOD and no migration event."
    AssertRepeatedAdminSeedCreatesUniqueKeysCreateWarehouse = True
End Function

Private Function InventoryLogContainsEventTypeCreateWarehouse(ByVal wb As Workbook, _
                                                              ByVal eventType As String) As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindTableCreateWarehouse(wb, "tblInventoryLog")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(GetCreateWarehouseTableValue( _
            lo, rowIndex, "EventType"))), eventType, vbTextCompare) = 0 Then
            InventoryLogContainsEventTypeCreateWarehouse = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FirstTableWithHeaderCreateWarehouse(ByVal wb As Workbook, _
                                                     ByVal headerName As String) As String
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If TableColumnIndexCreateWarehouse(lo, headerName) > 0 Then
                FirstTableWithHeaderCreateWarehouse = ws.Name & "!" & lo.Name
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FindTableCreateWarehouse(ByVal wb As Workbook, _
                                          ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindTableCreateWarehouse = ws.ListObjects(tableName)
        If Not FindTableCreateWarehouse Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function TableHasHeadersCreateWarehouse(ByVal lo As ListObject, _
                                                ByVal requiredHeaders As Variant) As Boolean
    Dim header As Variant

    If lo Is Nothing Then Exit Function
    For Each header In requiredHeaders
        If TableColumnIndexCreateWarehouse(lo, CStr(header)) = 0 Then Exit Function
    Next header
    TableHasHeadersCreateWarehouse = True
End Function

Private Function WorkbookTablesExcludeHeaderCreateWarehouse(ByVal wb As Workbook, _
                                                            ByVal prohibitedHeader As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    WorkbookTablesExcludeHeaderCreateWarehouse = True
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If TableColumnIndexCreateWarehouse(lo, prohibitedHeader) > 0 Then
                WorkbookTablesExcludeHeaderCreateWarehouse = False
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function CollectEntityKeysCreateWarehouse(ByVal lo As ListObject, _
                                                  ByVal requireGoodCondition As Boolean, _
                                                  ByRef detail As String) As Object
    Dim keys As Object
    Dim rowIndex As Long
    Dim systemKey As String
    Dim conditionValue As String

    If lo Is Nothing Then
        detail = "Required entity table was not found."
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        detail = lo.Name & " contained no seeded entity rows."
        Exit Function
    End If

    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare
    For rowIndex = 1 To lo.ListRows.Count
        systemKey = Trim$(CStr(GetCreateWarehouseTableValue(lo, rowIndex, "System_Key")))
        If systemKey = "" Then
            detail = lo.Name & " row " & CStr(rowIndex) & " has a blank System_Key."
            Exit Function
        End If
        If keys.Exists(systemKey) Then
            detail = lo.Name & " contains duplicate System_Key " & systemKey & "."
            Exit Function
        End If
        If requireGoodCondition Then
            conditionValue = UCase$(Trim$(CStr(GetCreateWarehouseTableValue(lo, rowIndex, "Condition"))))
            If conditionValue <> "GOOD" Then
                detail = lo.Name & " row " & CStr(rowIndex) & " Condition was not GOOD."
                Exit Function
            End If
        End If
        keys.Add systemKey, True
    Next rowIndex

    detail = lo.Name & " contains " & CStr(keys.Count) & " unique nonblank System_Key values with Condition=GOOD."
    Set CollectEntityKeysCreateWarehouse = keys
End Function

Private Function DictionariesHaveSameKeysCreateWarehouse(ByVal expectedKeys As Object, _
                                                         ByVal actualKeys As Object) As Boolean
    Dim key As Variant

    If expectedKeys Is Nothing Or actualKeys Is Nothing Then Exit Function
    If expectedKeys.Count <> actualKeys.Count Then Exit Function
    For Each key In expectedKeys.Keys
        If Not actualKeys.Exists(CStr(key)) Then Exit Function
    Next key
    DictionariesHaveSameKeysCreateWarehouse = True
End Function

Private Function AssertOperatorCustomColumnSurvivesRefreshCreateWarehouse( _
    ByVal wbOperator As Workbook, _
    ByVal loOperator As ListObject, _
    ByVal warehouseId As String, _
    ByVal inventoryKeys As Object, _
    ByRef detail As String) As Boolean

    Dim systemKey As String
    Dim rowIndex As Long
    Dim refreshReport As String

    If wbOperator Is Nothing Or loOperator Is Nothing Then Exit Function
    If inventoryKeys Is Nothing Or inventoryKeys.Count = 0 Then Exit Function

    systemKey = FirstDictionaryKeyCreateWarehouse(inventoryKeys)
    If TableColumnIndexCreateWarehouse(loOperator, "Custom_Local_Note") = 0 Then
        loOperator.ListColumns.Add loOperator.ListColumns.Count + 1
        loOperator.ListColumns(loOperator.ListColumns.Count).Name = "Custom_Local_Note"
    End If
    rowIndex = FindTableRowCreateWarehouse(loOperator, "System_Key", systemKey)
    If rowIndex = 0 Then
        detail = "System_Key was not found in operator table before custom-column refresh."
        Exit Function
    End If
    loOperator.DataBodyRange.Cells( _
        rowIndex, TableColumnIndexCreateWarehouse(loOperator, "Custom_Local_Note")).Value = "PRESERVE-ME"

    modRuntimeWorkbooks.SetCoreDataRootOverride mLocalRoot
    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook( _
        wbOperator, warehouseId, "LOCAL", refreshReport) Then
        detail = refreshReport
        Exit Function
    End If

    Set loOperator = FindTableCreateWarehouse(wbOperator, "invSys")
    AssertOperatorCustomColumnSurvivesRefreshCreateWarehouse = _
        OperatorCustomValuePresentCreateWarehouse( _
            loOperator, systemKey, "Custom_Local_Note", "PRESERVE-ME")
    If AssertOperatorCustomColumnSurvivesRefreshCreateWarehouse Then
        wbOperator.Save
        detail = "Custom_Local_Note survived refresh for System_Key " & systemKey & "."
    Else
        detail = "Custom_Local_Note was lost or moved to a different System_Key during refresh."
    End If
End Function

Private Function FirstDictionaryKeyCreateWarehouse(ByVal values As Object) As String
    Dim keys As Variant

    If values Is Nothing Then Exit Function
    If values.Count = 0 Then Exit Function
    keys = values.Keys
    FirstDictionaryKeyCreateWarehouse = CStr(keys(LBound(keys)))
End Function

Private Function OperatorCustomValuePresentCreateWarehouse( _
    ByVal lo As ListObject, _
    ByVal systemKey As String, _
    ByVal columnName As String, _
    ByVal expectedValue As String) As Boolean

    Dim rowIndex As Long
    Dim columnIndex As Long

    If lo Is Nothing Then Exit Function
    columnIndex = TableColumnIndexCreateWarehouse(lo, columnName)
    If columnIndex = 0 Then Exit Function
    rowIndex = FindTableRowCreateWarehouse(lo, "System_Key", systemKey)
    If rowIndex = 0 Then Exit Function
    OperatorCustomValuePresentCreateWarehouse = _
        (StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, columnIndex).Value)), _
                 expectedValue, vbBinaryCompare) = 0)
End Function

Private Function FindTableRowCreateWarehouse(ByVal lo As ListObject, _
                                             ByVal columnName As String, _
                                             ByVal matchValue As String) As Long
    Dim columnIndex As Long
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    columnIndex = TableColumnIndexCreateWarehouse(lo, columnName)
    If columnIndex = 0 Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, columnIndex).Value)), _
                   matchValue, vbTextCompare) = 0 Then
            FindTableRowCreateWarehouse = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function TableColumnIndexCreateWarehouse(ByVal lo As ListObject, _
                                                 ByVal columnName As String) As Long
    Dim columnIndex As Long

    If lo Is Nothing Then Exit Function
    For columnIndex = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(columnIndex).Name), Trim$(columnName), vbTextCompare) = 0 Then
            TableColumnIndexCreateWarehouse = columnIndex
            Exit Function
        End If
    Next columnIndex
End Function

Public Function GetCreateWarehouseContextPacked() As String
    GetCreateWarehouseContextPacked = _
        "WarehouseId=" & SafeCreateWarehouseText(mWarehouseId) & "|" & _
        "StationId=" & SafeCreateWarehouseText(mStationId) & "|" & _
        "LocalRoot=" & SafeCreateWarehouseText(mLocalRoot) & "|" & _
        "SharePointRoot=" & SafeCreateWarehouseText(mSharePointRoot) & "|" & _
        "Summary=" & SafeCreateWarehouseText(mSummary)
End Function

Public Function GetCreateWarehouseEvidenceRows() As String
    Dim i As Long

    For i = 1 To mCheckCount
        If Len(GetCreateWarehouseEvidenceRows) > 0 Then GetCreateWarehouseEvidenceRows = GetCreateWarehouseEvidenceRows & vbLf
        GetCreateWarehouseEvidenceRows = GetCreateWarehouseEvidenceRows & _
            mCheckNames(i) & vbTab & mCheckResults(i) & vbTab & mCheckDetails(i)
    Next i
End Function

Private Sub ResetCreateWarehouseEvidence()
    mCheckCount = 0
    Erase mCheckNames
    Erase mCheckResults
    Erase mCheckDetails
    mWarehouseId = vbNullString
    mStationId = vbNullString
    mLocalRoot = vbNullString
    mSharePointRoot = vbNullString
    mOperatorRoot = vbNullString
    mSummary = vbNullString
End Sub

Private Sub RecordCreateWarehouseCheck(ByVal checkName As String, _
                                       ByVal passed As Boolean, _
                                       ByVal detailText As String)
    mCheckCount = mCheckCount + 1
    ReDim Preserve mCheckNames(1 To mCheckCount)
    ReDim Preserve mCheckResults(1 To mCheckCount)
    ReDim Preserve mCheckDetails(1 To mCheckCount)

    mCheckNames(mCheckCount) = Trim$(checkName)
    mCheckResults(mCheckCount) = IIf(passed, "PASS", "FAIL")
    mCheckDetails(mCheckCount) = SafeCreateWarehouseText(detailText)
End Sub

Private Function AllCreateWarehouseChecksPassed() As Boolean
    Dim i As Long

    AllCreateWarehouseChecksPassed = (mCheckCount > 0)
    For i = 1 To mCheckCount
        If StrComp(mCheckResults(i), "PASS", vbTextCompare) <> 0 Then
            AllCreateWarehouseChecksPassed = False
            Exit Function
        End If
    Next i
End Function

Private Function AssertLocalStructureCreateWarehouse(ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
                                                     ByRef detailText As String) As Boolean
    Dim requiredPaths As Variant
    Dim item As Variant
    Dim operatorPath As String

    requiredPaths = Array( _
        spec.PathLocal, _
        spec.PathLocal & "\inbox", _
        spec.PathLocal & "\outbox", _
        spec.PathLocal & "\snapshots", _
        spec.PathLocal & "\config", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Config.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Auth.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".Outbox.Events.xlsb", _
        spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb")

    For Each item In requiredPaths
        If Not CreateWarehousePathExists(CStr(item)) Then
            detailText = "Missing path: " & CStr(item)
            Exit Function
        End If
    Next item

    operatorPath = modWarehouseBootstrap.GetLastWarehouseOperatorWorkbookPath()
    mOperatorRoot = GetParentFolderCreateWarehouse(operatorPath)
    If operatorPath = "" Or Not CreateWarehousePathExists(operatorPath) Then
        detailText = "Local receiving operator workbook was not created: " & operatorPath
        Exit Function
    End If
    If StrComp(Left$(operatorPath, Len(spec.PathLocal) + 1), spec.PathLocal & "\", vbTextCompare) = 0 Then
        detailText = "Receiving operator workbook was created under the warehouse hub: " & operatorPath
        Exit Function
    End If
    If Not CreateWarehousePathExists(mOperatorRoot & "\" & spec.WarehouseId & ".invSys.Config.xlsb") Then
        detailText = "Local operator config copy missing beside receiving workbook."
        Exit Function
    End If
    If Not CreateWarehousePathExists(mOperatorRoot & "\" & spec.WarehouseId & ".invSys.Auth.xlsb") Then
        detailText = "Local operator auth copy missing beside receiving workbook."
        Exit Function
    End If

    detailText = "All required runtime folders and seeded artifacts were created under " & spec.PathLocal
    AssertLocalStructureCreateWarehouse = True
End Function

Private Function AssertConfigSeededCreateWarehouse(ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
                                                   ByRef detailText As String) As Boolean
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject

    On Error GoTo FailAssert

    Set wbCfg = Application.Workbooks.Open(spec.PathLocal & "\" & spec.WarehouseId & ".invSys.Config.xlsb")
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")

    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "WarehouseId")), spec.WarehouseId, vbTextCompare) <> 0 Then
        detailText = "WarehouseId was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "WarehouseName")), spec.WarehouseName, vbTextCompare) <> 0 Then
        detailText = "WarehouseName was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "PathDataRoot")), spec.PathLocal, vbTextCompare) <> 0 Then
        detailText = "PathDataRoot was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loWh, 1, "PathSharePointRoot")), spec.PathSharePoint, vbTextCompare) <> 0 Then
        detailText = "PathSharePointRoot was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loSt, 1, "StationId")), spec.StationId, vbTextCompare) <> 0 Then
        detailText = "StationId row was not seeded correctly."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loSt, 1, "StationName")), spec.AdminUser, vbTextCompare) <> 0 Then
        detailText = "Admin user was not seeded into StationName."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetCreateWarehouseTableValue(loSt, 1, "RoleDefault")), "RECEIVE", vbTextCompare) <> 0 Then
        detailText = "RoleDefault was not seeded as RECEIVE."
        GoTo CleanExit
    End If

    detailText = "Config workbook seeded WarehouseId, WarehouseName, StationId, PathDataRoot, PathSharePointRoot, and RECEIVE defaults."
    AssertConfigSeededCreateWarehouse = True

CleanExit:
    CloseCreateWarehouseWorkbook wbCfg
    Exit Function

FailAssert:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function AssertSharePointArtifactsCreateWarehouse(ByRef spec As modWarehouseBootstrap.WarehouseSpec, _
                                                          ByRef detailText As String) As Boolean
    Dim discoveryPath As String
    Dim publishedConfigPath As String

    discoveryPath = spec.PathSharePoint & "\" & spec.WarehouseId & ".config.json"
    publishedConfigPath = spec.PathSharePoint & "\" & spec.WarehouseId & "\" & spec.WarehouseId & ".invSys.Config.xlsb"

    If Not CreateWarehousePathExists(discoveryPath) Then
        detailText = "Discovery artifact missing: " & discoveryPath
        Exit Function
    End If
    If Not CreateWarehousePathExists(publishedConfigPath) Then
        detailText = "Published config artifact missing: " & publishedConfigPath
        Exit Function
    End If

    detailText = "Discovery artifact and published config workbook exist under " & spec.PathSharePoint
    AssertSharePointArtifactsCreateWarehouse = True
End Function

Private Function GetCreateWarehouseTableValue(ByVal lo As ListObject, _
                                              ByVal rowIndex As Long, _
                                              ByVal columnName As String) As Variant
    Dim idx As Long

    idx = lo.ListColumns(columnName).Index
    GetCreateWarehouseTableValue = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function CreateWarehousePathExists(ByVal pathIn As String) As Boolean
    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function

    CreateWarehousePathExists = (Len(Dir$(pathIn, vbDirectory)) > 0)
    If Not CreateWarehousePathExists Then
        CreateWarehousePathExists = (Len(Dir$(pathIn, vbNormal)) > 0)
    End If
End Function

Private Function BuildCreateWarehouseTempRoot(ByVal leafName As String) As String
    BuildCreateWarehouseTempRoot = Environ$("TEMP") & "\invSys_createwarehouse_" & leafName & "_" & _
                                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Sub DeleteCreateWarehouseFolderRecursive(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Sub CloseCreateWarehouseWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function GetParentFolderCreateWarehouse(ByVal filePath As String) As String
    Dim sepPos As Long

    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function

    sepPos = InStrRev(filePath, "\")
    If sepPos > 1 Then GetParentFolderCreateWarehouse = Left$(filePath, sepPos - 1)
End Function

Private Function SafeCreateWarehouseText(ByVal textIn As String) As String
    SafeCreateWarehouseText = Replace$(Replace$(Trim$(textIn), vbCr, " "), vbLf, " ")
End Function
