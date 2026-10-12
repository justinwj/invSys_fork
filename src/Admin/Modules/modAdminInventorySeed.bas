Attribute VB_Name = "modAdminInventorySeed"
Option Explicit

Public Const DEMO_ACTION_SEED As String = "SEED"
Public Const DEMO_ACTION_DELETE As String = "DELETE"
Public Const DEMO_ACTION_UPLOAD As String = "UPLOAD"
Public Const DEMO_ACTION_UPLOAD_DATA_SET As String = "UPLOAD_DATASET"
Public Const DEMO_ACTION_DELETE_DATA_SET As String = "DELETE_DATASET"
Public Const R1_DATA_SET_ID As String = "R1_WORKFLOW_KIT"

Private Const DEMO_DATA_SET_RELATIVE_PATH As String = "admin\demo-inventory-data-sets"

Private mBuildActiveDemoGroups As Object
Private mBuildSkippedCount As Long

Public Function SeedDemoInventoryForWarehouse(ByVal warehouseId As String, _
                                              ByVal stationId As String, _
                                              ByVal userId As String, _
                                              ByRef report As String) As Boolean
    Dim activeGroups As Object
    Dim payloadItems As Collection
    Dim skippedCount As Long

    On Error GoTo FailSeed

    Set activeGroups = BuildActiveDemoGroupIndex(warehouseId, report)
    If activeGroups Is Nothing Then Exit Function
    Set payloadItems = BuildDemoInventoryPayload(activeGroups, skippedCount)
    If payloadItems.Count = 0 Then
        report = "Demo inventory already present.|Created=0|Skipped=" & CStr(skippedCount)
        SeedDemoInventoryForWarehouse = True
        Exit Function
    End If

    SeedDemoInventoryForWarehouse = QueueDemoCreateAndProcess( _
        warehouseId, stationId, userId, payloadItems, "Admin demo inventory seed", _
        "Demo inventory seeded.", skippedCount, report)
    Exit Function

FailSeed:
    report = "SeedDemoInventoryForWarehouse failed: " & Err.Description
End Function

Public Function DeleteDemoInventoryForWarehouse(ByVal warehouseId As String, _
                                                ByVal stationId As String, _
                                                ByVal userId As String, _
                                                ByRef report As String) As Boolean
    Dim inventoryWb As Workbook
    Dim entityTable As ListObject
    Dim payloadItems As Collection
    Dim item As Object
    Dim rowIndex As Long
    Dim systemKey As String
    Dim sku As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim qtyOnHand As Double
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim retryReport As String
    Dim processedCount As Long
    Dim eventApplied As Boolean, retryApplied As Boolean
    Dim inboxReport As String

    On Error GoTo FailDelete

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    If inventoryWb Is Nothing Then report = "Inventory workbook not found.": Exit Function
    Set entityTable = FindTableByNameSeed(inventoryWb, "tblInventoryEntities")
    If entityTable Is Nothing Then report = "Inventory entity projection was not found.": Exit Function

    Set payloadItems = New Collection
    If Not entityTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To entityTable.ListRows.Count
            sku = TableCellTextSeed(entityTable, rowIndex, "SKU")
            qtyOnHand = TableCellNumberSeed(entityTable, rowIndex, "QtyOnHand")
            If IsDemoSkuSeed(sku) And qtyOnHand > 0 Then
                systemKey = TableCellTextSeed(entityTable, rowIndex, "System_Key")
                If systemKey = "" Then
                    report = "Active demo inventory is missing System_Key identity."
                    Exit Function
                End If
                locationValue = TableCellTextSeed(entityTable, rowIndex, "Location")
                conditionValue = UCase$(TableCellTextSeed(entityTable, rowIndex, "Condition"))
                If conditionValue = "" Then conditionValue = "GOOD"
                Set item = CreateObject("Scripting.Dictionary")
                item.CompareMode = vbTextCompare
                item("System_Key") = systemKey
                item("SKU") = sku
                item("ITEM_CODE") = sku
                item("Qty") = -qtyOnHand
                item("Location") = locationValue
                item("Condition") = conditionValue
                item("IoType") = "ADJUST"
                item("Note") = "Admin delete demo inventory"
                payloadItems.Add item
            End If
        Next rowIndex
    End If

    If payloadItems.Count = 0 Then
        report = "No active demo inventory was found.|Depleted=0"
        DeleteDemoInventoryForWarehouse = True
        Exit Function
    End If
    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If Not modRoleEventWriter.QueueAdminInventoryAdjustEvent( _
            warehouseId, stationId, userId, payloadJson, "Admin delete demo inventory", _
            0, eventIdOut, queueError, "") Then
        report = "Delete event could not be queued: " & queueError
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport, eventIdOut, eventApplied)
    If processedCount < 1 And InStr(1, batchReport, "Poison=0", vbTextCompare) > 0 Then
        processedCount = modProcessor.RunBatch(warehouseId, 0, retryReport, eventIdOut, retryApplied)
        eventApplied = eventApplied Or retryApplied
        batchReport = batchReport & "|Retry=" & retryReport
    End If
    If Not eventApplied Then
        report = "Delete event was queued; application was not confirmed. " & batchReport
        Exit Function
    End If

    report = "Demo inventory deleted from active stock.|Depleted=" & _
             CStr(payloadItems.Count) & "|Applied=" & CStr(processedCount) & _
             "|Processor=" & batchReport
    DeleteDemoInventoryForWarehouse = True
    Exit Function

FailDelete:
    report = "DeleteDemoInventoryForWarehouse failed: " & Err.Description
End Function

Public Function UploadDemoInventoryForWarehouse(ByVal warehouseId As String, _
                                                ByVal stationId As String, _
                                                ByVal userId As String, _
                                                ByVal csvPath As String, _
                                                ByRef report As String) As Boolean
    Dim activeGroups As Object
    Dim payloadItems As Collection
    Dim skippedCount As Long

    On Error GoTo FailUpload

    csvPath = Trim$(csvPath)
    If csvPath = "" Then report = "Choose a demo inventory CSV file.": Exit Function
    If LCase$(Right$(csvPath, 4)) <> ".csv" Then report = "Demo inventory upload requires a .csv file.": Exit Function
    If Len(Dir$(csvPath, vbNormal)) = 0 Then report = "Demo inventory CSV was not found.": Exit Function

    Set activeGroups = BuildActiveDemoGroupIndex(warehouseId, report)
    If activeGroups Is Nothing Then Exit Function
    Set payloadItems = LoadDemoInventoryCsv(csvPath, activeGroups, skippedCount, report)
    If payloadItems Is Nothing Then Exit Function
    If payloadItems.Count = 0 Then
        report = "Uploaded demo inventory is already present.|Created=0|Skipped=" & CStr(skippedCount)
        UploadDemoInventoryForWarehouse = True
        Exit Function
    End If

    UploadDemoInventoryForWarehouse = QueueDemoCreateAndProcess( _
        warehouseId, stationId, userId, payloadItems, "Admin demo inventory upload", _
        "Demo inventory uploaded.", skippedCount, report)
    Exit Function

FailUpload:
    report = "UploadDemoInventoryForWarehouse failed: " & Err.Description
End Function

Public Function DemoInventoryDataSetLibraryPath(ByVal runtimeRoot As String) As String
    runtimeRoot = Trim$(runtimeRoot)
    Do While Right$(runtimeRoot, 1) = "\" Or Right$(runtimeRoot, 1) = "/"
        runtimeRoot = Left$(runtimeRoot, Len(runtimeRoot) - 1)
    Loop
    If runtimeRoot = "" Then Exit Function
    DemoInventoryDataSetLibraryPath = runtimeRoot & "\" & DEMO_DATA_SET_RELATIVE_PATH
End Function

Public Function ListDemoInventoryDataSets(ByVal runtimeRoot As String) As Collection
    Dim results As Collection
    Dim libraryPath As String
    Dim fileName As String

    Set results = New Collection
    libraryPath = DemoInventoryDataSetLibraryPath(runtimeRoot)
    If libraryPath <> "" And Len(Dir$(libraryPath, vbDirectory)) > 0 Then
        fileName = Dir$(libraryPath & "\*.csv", vbNormal)
        Do While fileName <> ""
            results.Add Array("Uploaded: " & fileName, libraryPath & "\" & fileName)
            fileName = Dir$()
        Loop
    End If
    Set ListDemoInventoryDataSets = results
End Function

Public Function ImportDemoInventoryDataSet(ByVal runtimeRoot As String, _
                                           ByVal sourcePath As String, _
                                           ByRef storedPath As String, _
                                           ByRef report As String) As Boolean
    Dim emptyGroups As Object
    Dim validatedRows As Collection
    Dim skippedCount As Long
    Dim libraryPath As String
    Dim sourceName As String
    Dim fso As Object

    On Error GoTo FailImport

    storedPath = ""
    sourcePath = Trim$(sourcePath)
    If sourcePath = "" Then report = "Choose a demo inventory CSV data set.": Exit Function
    If LCase$(Right$(sourcePath, 4)) <> ".csv" Then report = "Demo inventory data sets must be .csv files.": Exit Function
    If Len(Dir$(sourcePath, vbNormal)) = 0 Then report = "The selected demo inventory data set was not found.": Exit Function

    Set emptyGroups = CreateObject("Scripting.Dictionary")
    emptyGroups.CompareMode = vbTextCompare
    Set validatedRows = LoadDemoInventoryCsv(sourcePath, emptyGroups, skippedCount, report, False)
    If validatedRows Is Nothing Then Exit Function
    If validatedRows.Count = 0 Then report = "The demo inventory data set contains no rows.": Exit Function

    libraryPath = DemoInventoryDataSetLibraryPath(runtimeRoot)
    If libraryPath = "" Then report = "The selected warehouse runtime root is unavailable.": Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Trim$(runtimeRoot)) Then report = "The selected warehouse runtime root is unavailable.": Exit Function
    If Not EnsureDemoDataSetLibrarySeed(runtimeRoot, libraryPath, report) Then Exit Function

    sourceName = fso.GetFileName(sourcePath)
    storedPath = libraryPath & "\" & sourceName
    If fso.FileExists(storedPath) Then
        report = "A data set named " & sourceName & " already exists. Delete it first or rename the source file."
        storedPath = ""
        Exit Function
    End If
    fso.CopyFile sourcePath, storedPath, False
    report = "Demo inventory data set uploaded.|DataSet=" & sourceName
    ImportDemoInventoryDataSet = True
    Exit Function

FailImport:
    storedPath = ""
    report = "ImportDemoInventoryDataSet failed: " & Err.Description
End Function

Public Function DeleteDemoInventoryDataSet(ByVal runtimeRoot As String, _
                                           ByVal selectedPath As String, _
                                           ByRef report As String) As Boolean
    Dim libraryPath As String
    Dim normalizedLibrary As String
    Dim normalizedSelected As String
    Dim fso As Object

    On Error GoTo FailDeleteDataSet

    If Trim$(selectedPath) = "" Or StrComp(Trim$(selectedPath), R1_DATA_SET_ID, vbTextCompare) = 0 Then
        report = "The built-in R1 Workflow Kit cannot be deleted."
        Exit Function
    End If
    libraryPath = DemoInventoryDataSetLibraryPath(runtimeRoot)
    If libraryPath = "" Then report = "The selected warehouse runtime root is unavailable.": Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    normalizedLibrary = fso.GetAbsolutePathName(libraryPath)
    normalizedSelected = fso.GetAbsolutePathName(Trim$(selectedPath))
    If StrComp(Left$(normalizedSelected, Len(normalizedLibrary) + 1), _
               normalizedLibrary & "\", vbTextCompare) <> 0 Then
        report = "Only an uploaded data set in the selected warehouse library can be deleted."
        Exit Function
    End If
    If LCase$(Right$(normalizedSelected, 4)) <> ".csv" Then
        report = "Only an uploaded CSV data set can be deleted."
        Exit Function
    End If
    If Not fso.FileExists(normalizedSelected) Then report = "The selected uploaded data set was not found.": Exit Function

    fso.DeleteFile normalizedSelected, False
    report = "Demo inventory data set deleted.|DataSet=" & fso.GetFileName(normalizedSelected)
    DeleteDemoInventoryDataSet = True
    Exit Function

FailDeleteDataSet:
    report = "DeleteDemoInventoryDataSet failed: " & Err.Description
End Function

Public Function DescribeDemoInventoryDataSetsForAutomation(ByVal runtimeRoot As String) As String
    Dim dataSets As Collection
    Set dataSets = ListDemoInventoryDataSets(runtimeRoot)
    DescribeDemoInventoryDataSetsForAutomation = "OK|R1=Protected|Uploaded=" & CStr(dataSets.Count)
End Function

Public Function DescribeDemoInventoryStateForAutomation(ByVal warehouseId As String) As String
    Dim inventoryWb As Workbook
    Dim entityTable As ListObject
    Dim groups As Object
    Dim keys As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim systemKey As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim qtyOnHand As Double
    Dim entityCount As Long
    Dim activeCount As Long

    On Error GoTo FailDescribe

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    If inventoryWb Is Nothing Then DescribeDemoInventoryStateForAutomation = "FAIL|Inventory workbook not found.": Exit Function
    Set entityTable = FindTableByNameSeed(inventoryWb, "tblInventoryEntities")
    If entityTable Is Nothing Then DescribeDemoInventoryStateForAutomation = "FAIL|Entity projection not found.": Exit Function
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare

    If Not entityTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To entityTable.ListRows.Count
            sku = TableCellTextSeed(entityTable, rowIndex, "SKU")
            If IsDemoSkuSeed(sku) Then
                entityCount = entityCount + 1
                systemKey = TableCellTextSeed(entityTable, rowIndex, "System_Key")
                If systemKey <> "" Then keys(systemKey) = True
                qtyOnHand = TableCellNumberSeed(entityTable, rowIndex, "QtyOnHand")
                If qtyOnHand > 0 Then
                    activeCount = activeCount + 1
                    locationValue = TableCellTextSeed(entityTable, rowIndex, "Location")
                    conditionValue = UCase$(TableCellTextSeed(entityTable, rowIndex, "Condition"))
                    If conditionValue = "" Then conditionValue = "GOOD"
                    groups(DemoGroupKeySeed(sku, locationValue, conditionValue)) = True
                End If
            End If
        Next rowIndex
    End If

    DescribeDemoInventoryStateForAutomation = "OK|Entities=" & CStr(entityCount) & _
        "|UniqueKeys=" & CStr(keys.Count) & "|Active=" & CStr(activeCount) & _
        "|ActiveGroups=" & CStr(groups.Count)
    Exit Function

FailDescribe:
    DescribeDemoInventoryStateForAutomation = "FAIL|" & Err.Description
End Function

Private Function FindOpenWorkbookByPathSeed(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    fullPath = Trim$(fullPath)
    If fullPath = "" Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathSeed = wb
            Exit Function
        End If
    Next wb
End Function

Private Function QueueDemoCreateAndProcess(ByVal warehouseId As String, _
                                           ByVal stationId As String, _
                                           ByVal userId As String, _
                                           ByVal payloadItems As Collection, _
                                           ByVal eventNote As String, _
                                           ByVal successText As String, _
                                           ByVal skippedCount As Long, _
                                           ByRef report As String) As Boolean
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim retryReport As String
    Dim processedCount As Long
    Dim eventApplied As Boolean, retryApplied As Boolean
    Dim inboxReport As String
    Dim productionInboxPath As String
    Dim productionInbox As Workbook
    Dim productionInboxWasOpen As Boolean

    On Error GoTo FailQueue

    If payloadItems Is Nothing Then report = "Demo inventory payload was not supplied.": Exit Function
    If payloadItems.Count = 0 Then report = "Demo inventory payload was empty.": Exit Function
    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Demo inventory payload was empty."
        Exit Function
    End If

    productionInboxPath = modRoleEventWriter.ResolveInboxWorkbookPath( _
        "INVENTORY_CREATE", warehouseId, stationId, queueError)
    If productionInboxPath = "" Then
        report = "Production inbox path could not be resolved: " & queueError
        Exit Function
    End If
    Set productionInbox = FindOpenWorkbookByPathSeed(productionInboxPath)
    productionInboxWasOpen = Not productionInbox Is Nothing
    If productionInbox Is Nothing Then
        Set productionInbox = modRoleEventWriter.OpenInboxWorkbook( _
            "INVENTORY_CREATE", warehouseId, stationId, queueError)
    End If
    If productionInbox Is Nothing Then
        report = "Production inbox could not be opened: " & queueError
        Exit Function
    End If

    If Not modRoleEventWriter.QueueInventoryCreateEvent( _
            warehouseId, stationId, userId, payloadJson, eventNote, _
            0, productionInbox, eventIdOut, queueError, "") Then
        report = "Demo inventory event could not be queued: " & queueError
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport, eventIdOut, eventApplied)
    If processedCount < 1 And InStr(1, batchReport, "Poison=0", vbTextCompare) > 0 Then
        processedCount = modProcessor.RunBatch(warehouseId, 0, retryReport, eventIdOut, retryApplied)
        eventApplied = eventApplied Or retryApplied
        batchReport = batchReport & "|Retry=" & retryReport
    End If
    If Not eventApplied Then
        report = "Demo inventory event was queued; application was not confirmed. " & batchReport
        GoTo CleanExit
    End If

    report = successText & "|Created=" & CStr(payloadItems.Count) & _
             "|Skipped=" & CStr(skippedCount) & "|Applied=" & _
             CStr(processedCount) & "|Processor=" & batchReport
    QueueDemoCreateAndProcess = True

CleanExit:
    On Error Resume Next
    If Not productionInboxWasOpen Then productionInbox.Close SaveChanges:=True
    On Error GoTo 0
    Exit Function

FailQueue:
    report = "QueueDemoCreateAndProcess failed: " & Err.Description
    Resume CleanExit
End Function

Private Function BuildActiveDemoGroupIndex(ByVal warehouseId As String, _
                                           ByRef report As String) As Object
    Dim inventoryWb As Workbook
    Dim entityTable As ListObject
    Dim groups As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim locationValue As String
    Dim conditionValue As String

    On Error GoTo FailIndex

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    If inventoryWb Is Nothing Then
        report = "Inventory workbook not found."
        Exit Function
    End If
    Set entityTable = FindTableByNameSeed(inventoryWb, "tblInventoryEntities")
    If entityTable Is Nothing Then
        report = "Inventory entity projection was not found."
        Exit Function
    End If

    If Not entityTable.DataBodyRange Is Nothing Then
        For rowIndex = 1 To entityTable.ListRows.Count
            sku = TableCellTextSeed(entityTable, rowIndex, "SKU")
            If IsDemoSkuSeed(sku) _
               And TableCellNumberSeed(entityTable, rowIndex, "QtyOnHand") > 0 Then
                locationValue = TableCellTextSeed(entityTable, rowIndex, "Location")
                conditionValue = UCase$(TableCellTextSeed(entityTable, rowIndex, "Condition"))
                If conditionValue = "" Then conditionValue = "GOOD"
                groups(DemoGroupKeySeed(sku, locationValue, conditionValue)) = True
            End If
        Next rowIndex
    End If
    Set BuildActiveDemoGroupIndex = groups
    Exit Function

FailIndex:
    report = "BuildActiveDemoGroupIndex failed: " & Err.Description
End Function

Private Function LoadDemoInventoryCsv(ByVal csvPath As String, _
                                      ByVal activeGroups As Object, _
                                      ByRef skippedCount As Long, _
                                      ByRef report As String, _
                                      Optional ByVal createPayloadItems As Boolean = True) As Collection
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim values As Variant
    Dim headers As Object
    Dim fileGroups As Object
    Dim rows As Collection
    Dim item As Object
    Dim rowIndex As Long
    Dim itemCode As String
    Dim itemName As String
    Dim qtyValue As Variant
    Dim qty As Double
    Dim uomValue As String
    Dim locationValue As String
    Dim conditionValue As String
    Dim descriptionValue As String
    Dim categoryValue As String
    Dim vendorValue As String
    Dim groupKey As String

    On Error GoTo FailLoad

    Set sourceWb = Application.Workbooks.Open( _
        Filename:=csvPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Local:=True)
    Set sourceWs = sourceWb.Worksheets(1)
    values = sourceWs.UsedRange.Value2
    If Not IsArray(values) Then
        report = "Demo inventory CSV must contain a header row and at least one data row."
        GoTo InvalidLoad
    End If
    If UBound(values, 1) < 2 Then
        report = "Demo inventory CSV must contain at least one data row."
        GoTo InvalidLoad
    End If

    Set headers = BuildCsvHeaderIndexSeed(values)
    If Not CsvHasRequiredHeadersSeed(headers, report) Then GoTo InvalidLoad
    Set fileGroups = CreateObject("Scripting.Dictionary")
    fileGroups.CompareMode = vbTextCompare
    Set rows = New Collection

    For rowIndex = 2 To UBound(values, 1)
        itemCode = Trim$(CsvValueSeed(values, rowIndex, headers, "ITEM_CODE"))
        itemName = Trim$(CsvValueSeed(values, rowIndex, headers, "ITEM"))
        uomValue = Trim$(CsvValueSeed(values, rowIndex, headers, "UOM"))
        locationValue = Trim$(CsvValueSeed(values, rowIndex, headers, "LOCATION"))
        conditionValue = UCase$(Trim$(CsvValueSeed(values, rowIndex, headers, "CONDITION")))
        If conditionValue = "" Then conditionValue = "GOOD"
        If itemCode = "" And itemName = "" And uomValue = "" And locationValue = "" Then GoTo NextCsvRow
        If Not IsDemoSkuSeed(itemCode) Then
            report = "CSV row " & CStr(rowIndex) & " ITEM_CODE must begin with DEMO-."
            GoTo InvalidLoad
        End If
        If itemName = "" Or uomValue = "" Or locationValue = "" Then
            report = "CSV row " & CStr(rowIndex) & " requires ITEM, UOM, and LOCATION."
            GoTo InvalidLoad
        End If
        qtyValue = CsvRawValueSeed(values, rowIndex, headers, "QTY")
        If Not IsNumeric(qtyValue) Or CDbl(qtyValue) <= 0 Then
            report = "CSV row " & CStr(rowIndex) & " QTY must be greater than zero."
            GoTo InvalidLoad
        End If
        qty = CDbl(qtyValue)
        groupKey = DemoGroupKeySeed(itemCode, locationValue, conditionValue)
        If fileGroups.Exists(groupKey) Then
            report = "CSV row " & CStr(rowIndex) & " duplicates an item/location/condition group in the file."
            GoTo InvalidLoad
        End If
        fileGroups.Add groupKey, True
        If Not activeGroups Is Nothing Then
            If activeGroups.Exists(groupKey) Then
                skippedCount = skippedCount + 1
                GoTo NextCsvRow
            End If
        End If

        descriptionValue = Trim$(CsvValueSeed(values, rowIndex, headers, "DESCRIPTION"))
        categoryValue = Trim$(CsvValueSeed(values, rowIndex, headers, "CATEGORY"))
        vendorValue = Trim$(CsvValueSeed(values, rowIndex, headers, "VENDOR"))
        If createPayloadItems Then
            Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
                modRoleEventWriter.CreateSystemKey(), itemCode, qty, locationValue, _
                conditionValue, "", "Admin demo inventory CSV upload")
            item("ITEM_CODE") = itemCode
            item("ITEM") = itemName
            item("UOM") = uomValue
            item("DESCRIPTION") = descriptionValue
            item("CATEGORY") = categoryValue
            If vendorValue <> "" Then item("VENDOR(s)") = vendorValue
        Else
            Set item = CreateObject("Scripting.Dictionary")
            item.CompareMode = vbTextCompare
            item("ITEM_CODE") = itemCode
        End If
        rows.Add item
NextCsvRow:
    Next rowIndex

    Set LoadDemoInventoryCsv = rows
InvalidLoad:
    On Error Resume Next
    sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

FailLoad:
    report = "Demo inventory CSV could not be read: " & Err.Description
    Resume InvalidLoad
End Function

Private Function BuildCsvHeaderIndexSeed(ByVal values As Variant) As Object
    Dim headers As Object
    Dim columnIndex As Long
    Dim normalized As String

    Set headers = CreateObject("Scripting.Dictionary")
    headers.CompareMode = vbTextCompare
    For columnIndex = 1 To UBound(values, 2)
        normalized = NormalizeCsvHeaderSeed(CStr(values(1, columnIndex)))
        If normalized = "VENDOR(S)" Or normalized = "VENDORS" Then normalized = "VENDOR"
        If normalized <> "" And Not headers.Exists(normalized) Then headers.Add normalized, columnIndex
    Next columnIndex
    Set BuildCsvHeaderIndexSeed = headers
End Function

Private Function CsvHasRequiredHeadersSeed(ByVal headers As Object, _
                                           ByRef report As String) As Boolean
    Dim required As Variant
    Dim headerName As Variant

    required = Array("ITEM_CODE", "ITEM", "QTY", "UOM", "LOCATION")
    For Each headerName In required
        If Not headers.Exists(CStr(headerName)) Then
            report = "Demo inventory CSV is missing required header " & CStr(headerName) & "."
            Exit Function
        End If
    Next headerName
    CsvHasRequiredHeadersSeed = True
End Function

Private Function CsvValueSeed(ByVal values As Variant, ByVal rowIndex As Long, _
                              ByVal headers As Object, ByVal headerName As String) As String
    Dim rawValue As Variant

    rawValue = CsvRawValueSeed(values, rowIndex, headers, headerName)
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    CsvValueSeed = CStr(rawValue)
End Function

Private Function CsvRawValueSeed(ByVal values As Variant, ByVal rowIndex As Long, _
                                 ByVal headers As Object, ByVal headerName As String) As Variant
    If headers Is Nothing Then Exit Function
    If Not headers.Exists(headerName) Then Exit Function
    CsvRawValueSeed = values(rowIndex, CLng(headers(headerName)))
End Function

Private Function NormalizeCsvHeaderSeed(ByVal headerText As String) As String
    headerText = UCase$(Trim$(headerText))
    headerText = Replace$(headerText, " ", "_")
    headerText = Replace$(headerText, "-", "_")
    Do While InStr(1, headerText, "__", vbBinaryCompare) > 0
        headerText = Replace$(headerText, "__", "_")
    Loop
    NormalizeCsvHeaderSeed = headerText
End Function

Private Function EnsureDemoDataSetLibrarySeed(ByVal runtimeRoot As String, _
                                              ByVal libraryPath As String, _
                                              ByRef report As String) As Boolean
    Dim fso As Object
    Dim adminPath As String

    On Error GoTo FailEnsureLibrary

    Set fso = CreateObject("Scripting.FileSystemObject")
    runtimeRoot = fso.GetAbsolutePathName(Trim$(runtimeRoot))
    libraryPath = fso.GetAbsolutePathName(Trim$(libraryPath))
    If StrComp(Left$(libraryPath, Len(runtimeRoot) + 1), runtimeRoot & "\", vbTextCompare) <> 0 Then
        report = "The demo inventory data set library is outside the selected warehouse runtime."
        Exit Function
    End If
    adminPath = runtimeRoot & "\admin"
    If Not fso.FolderExists(adminPath) Then fso.CreateFolder adminPath
    If Not fso.FolderExists(libraryPath) Then fso.CreateFolder libraryPath
    EnsureDemoDataSetLibrarySeed = True
    Exit Function

FailEnsureLibrary:
    report = "The demo inventory data set library could not be prepared: " & Err.Description
End Function

Private Function FindTableByNameSeed(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByNameSeed = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function TableColumnIndexSeed(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), columnName, vbTextCompare) = 0 Then
            TableColumnIndexSeed = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function TableCellTextSeed(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                   ByVal columnName As String) As String
    Dim columnIndex As Long
    Dim valueIn As Variant

    columnIndex = TableColumnIndexSeed(lo, columnName)
    If columnIndex = 0 Or lo.DataBodyRange Is Nothing Then Exit Function
    valueIn = lo.DataBodyRange.Cells(rowIndex, columnIndex).Value2
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    TableCellTextSeed = Trim$(CStr(valueIn))
End Function

Private Function TableCellNumberSeed(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                     ByVal columnName As String) As Double
    Dim valueIn As Variant
    Dim columnIndex As Long

    columnIndex = TableColumnIndexSeed(lo, columnName)
    If columnIndex = 0 Or lo.DataBodyRange Is Nothing Then Exit Function
    valueIn = lo.DataBodyRange.Cells(rowIndex, columnIndex).Value2
    If IsNumeric(valueIn) Then TableCellNumberSeed = CDbl(valueIn)
End Function

Private Function DemoGroupKeySeed(ByVal itemCode As String, _
                                  ByVal locationValue As String, _
                                  ByVal conditionValue As String) As String
    DemoGroupKeySeed = UCase$(Trim$(itemCode)) & Chr$(30) & _
                       UCase$(Trim$(locationValue)) & Chr$(30) & _
                       UCase$(Trim$(conditionValue))
End Function

Private Function IsDemoSkuSeed(ByVal sku As String) As Boolean
    IsDemoSkuSeed = (StrComp(Left$(Trim$(sku), 5), "DEMO-", vbTextCompare) = 0)
End Function

Private Function EnsureDemoStationInboxes(ByVal warehouseId As String, _
                                          ByVal stationId As String, _
                                          ByRef report As String) As Boolean
    Dim inboxPath As String
    Dim stepReport As String

    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "RECEIVE", "", inboxPath, stepReport) Then
        report = stepReport
        Exit Function
    End If
    inboxPath = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "SHIP", "", inboxPath, stepReport) Then
        report = stepReport
        Exit Function
    End If
    inboxPath = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "PRODUCTION", "", inboxPath, stepReport) Then
        report = stepReport
        Exit Function
    End If

    EnsureDemoStationInboxes = True
End Function

Private Function BuildDemoInventoryPayload(ByVal activeGroups As Object, _
                                           ByRef skippedCount As Long) As Collection
    Dim rows As Collection

    Set mBuildActiveDemoGroups = activeGroups
    mBuildSkippedCount = 0
    Set rows = New Collection
    AddDemoInventoryItem rows, "DEMO-RAW-BLACK-TEA", "Black Tea", 5000, "lbs", _
        "Loose black tea for brewing.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-FILTERED-WATER", "Filtered Water", 20000, "lbs", _
        "Filtered brewing water.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-CARDAMOM", "Cardamom (Decorticated)", 500, "lbs", _
        "Cardamom for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-BLACK-PEPPER", "Black Pepper (Whole)", 300, "lbs", _
        "Black pepper for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-NUTMEG", "Nutmeg (Ground)", 250, "lbs", _
        "Ground nutmeg for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-GINGER", "Ginger (Ground)", 250, "lbs", _
        "Ground ginger for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-CITRIC-ACID", "Citric Acid", 120, "lbs", _
        "Citric acid ingredient.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-CASSIA-OIL", "Cassia Oil 340139", 80, "lbs", _
        "Cassia oil for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-LEMON-OIL", "Lemon Oil (5x) 34013", 80, "lbs", _
        "Lemon oil for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-ORANGE-OIL", "Orange Oil (Cold Press)", 80, "lbs", _
        "Orange oil for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-SUGAR-WHITE", "Pure Cane Sugar White Granulated", 8000, "lbs", _
        "White granulated cane sugar.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-SUGAR-CLOUDY", "Pure Cane Sugar Cloudy White Granulated", 6000, "lbs", _
        "Cloudy white granulated cane sugar.", "raw"
    AddDemoInventoryItem rows, "DEMO-WIP-BREWED-BLACK-TEA", "Brewed Black Tea", 1200, "lbs", _
        "Intermediate brewed tea.", "wip"
    AddDemoInventoryItem rows, "DEMO-WIP-CHAI-SPICE-BLEND", "Classic Chai Spice Blend", 600, "lbs", _
        "Intermediate chai spice blend.", "wip"
    AddDemoInventoryItem rows, "DEMO-RAW-BROWN-COLOR", "Brown Color 10.5g", 100, "lbs", _
        "Brown color ingredient.", "raw"
    AddDemoInventoryItem rows, "DEMO-FG-CLASSIC-CHAI", "Black Scottie Chai Classic Concentrate", 400, "gal", _
        "Finished good concentrate for shipping.", "shippable"
    AddDemoInventoryItem rows, "DEMO-FG-12PACK-CASE", "Classic Chai 12-Pack Case", 120, "each", _
        "Finished case pack.", "shippable"
    AddDemoInventoryItem rows, "DEMO-FG-SAMPLE-BOX", "Black Scottie Sample Box", 80, "each", _
        "Sample assortment box.", "shippable"
    AddDemoInventoryItem rows, "DEMO-PKG-TIN", "Classic Chai Tin", 1000, "each", _
        "Packaging component for receiving and box-building tests.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-SHIPPING-CARTON", "Shipping Carton Blank", 1000, "each", _
        "Corrugated carton blank for Box Designer and Box Maker tests.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-CASE-DIVIDER", "Case Divider", 2000, "each", _
        "Divider insert for shipping-box alternatives.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-SHIPPING-LABEL", "Shipping Label", 5000, "each", _
        "Adhesive label consumed when making a shipping box.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-PACKING-TAPE", "Packing Tape Strip", 5000, "each", _
        "Premeasured packing-tape strip for shipping assembly tests.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-VOID-FILL", "Void Fill Sheet", 3000, "each", _
        "Protective void fill for shipping-box alternatives.", "packaging.ship"
    skippedCount = mBuildSkippedCount
    Set mBuildActiveDemoGroups = Nothing
    Set BuildDemoInventoryPayload = rows
End Function

Private Sub AddDemoInventoryItem(ByVal rows As Collection, _
                                 ByVal itemCode As String, _
                                 ByVal itemName As String, _
                                 ByVal qty As Double, _
                                 ByVal uom As String, _
                                 ByVal description As String, _
                                 ByVal category As String)
    Dim item As Object
    Dim groupKey As String

    groupKey = DemoGroupKeySeed(itemCode, "CLEARVIEW", "GOOD")
    If Not mBuildActiveDemoGroups Is Nothing Then
        If mBuildActiveDemoGroups.Exists(groupKey) Then
            mBuildSkippedCount = mBuildSkippedCount + 1
            Exit Sub
        End If
    End If

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), itemCode, qty, "CLEARVIEW", _
        "GOOD", "", "Admin R1 workflow demo seed")
    item("ITEM_CODE") = itemCode
    item("ITEM") = itemName
    item("UOM") = uom
    item("DESCRIPTION") = description
    item("CATEGORY") = category
    rows.Add item
End Sub
