Attribute VB_Name = "modTesterProvisioning"
Option Explicit


Private mLastTesterSetupReport As String
Private mLastTesterOperatorWorkbookPath As String
Private mLastTesterSharePointRoot As String
Private mTesterSharePointRootOverride As String

Public Type TesterProvisionSpec
    UserId As String
    PinHash As String
    WarehouseId As String
    StationId As String
    PathLocal As String
    PathSharePointRoot As String
End Type

Public Function SetupTesterStation(ByRef spec As TesterProvisionSpec) As Boolean
    Dim report As String
    Dim rootPath As String
    Dim configPath As String
    Dim inboxPath As String
    Dim runtimeExists As Boolean
    Dim seedReport As String
    Dim authReport As String
    Dim operatorReport As String
    Dim configReport As String
    Dim sharePointReport As String
    Dim operatorPath As String
    Dim priorRootOverride As String
    Dim runtimeArtifactsExist As Boolean
    Dim runtimeCreated As Boolean

    On Error GoTo FailSetup

    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    ResetTesterSetupState
    NormalizeTesterSetupSpec spec
    If Not modTesterCommands.IsWriting() Then ApplyTesterSetupDefaults spec

    If Not ValidateTesterSetupSpec(spec, report) Then GoTo FailSoft

    rootPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    If rootPath = "" Then
        report = "PathLocal could not be resolved."
        GoTo FailSoft
    End If

    spec.PathLocal = rootPath
    configPath = rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    operatorPath = BuildReceivingOperatorPathTesterSetup(spec)
    mLastTesterOperatorWorkbookPath = operatorPath

    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    If Not modTesterCommands.EnterProvisioningRoot(rootPath) Then GoTo FailSoft

    runtimeExists = FolderExistsTesterSetup(rootPath)
    runtimeArtifactsExist = RuntimeArtifactsExistTesterSetup(rootPath, spec.WarehouseId)
    If Not runtimeExists And IsUncPathTesterSetup(rootPath) Then
        report = "Warehouse hub path is not accessible: " & rootPath
        GoTo FailSoft
    End If
    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    If Not runtimeArtifactsExist Then
        If Not CreateTesterRuntimeArtifactsTesterSetup(spec, rootPath, report) Then
            GoTo FailSoft
        End If
        runtimeExists = True
        runtimeArtifactsExist = True
        runtimeCreated = True
    End If
    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    EnsureFolderRecursiveTesterSetup rootPath & "\auth"

    If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, Environ$("COMPUTERNAME"), _
                                              rootPath & "\inbox\", "RECEIVE", configPath, rootPath, report) Then
        GoTo FailSoft
    End If
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "RECEIVE", configPath, inboxPath, report) Then
        GoTo FailSoft
    End If
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "SHIP", configPath, inboxPath, report) Then
        GoTo FailSoft
    End If
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "PRODUCTION", configPath, inboxPath, report) Then
        GoTo FailSoft
    End If

    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    If Not SeedTesterScenarioInventory(spec, seedReport) Then
        report = seedReport
        GoTo FailSoft
    End If

    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    If Not ProvisionTesterAuth(spec, authReport) Then
        report = authReport
        GoTo FailSoft
    End If

    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    If Not CreateOrVerifyReceivingWorkbook(spec, operatorPath, operatorReport) Then
        report = operatorReport
        GoTo FailSoft
    End If

    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    If Not UpdateConfigSharePointRoot(spec, configReport) Then
        report = configReport
        GoTo FailSoft
    End If
    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    If Not EnsureTesterSharePointPackage(spec, sharePointReport) Then
        report = sharePointReport
        GoTo FailSoft
    End If

    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    report = "OK|WarehouseId=" & spec.WarehouseId & _
             "|StationId=" & spec.StationId & _
             "|UserId=" & spec.UserId & _
             "|Runtime=" & IIf(runtimeCreated, "CREATED", "EXISTING") & _
             "|Inbox=" & inboxPath & _
             "|Seed=" & seedReport & _
             "|Auth=" & authReport & _
             "|Operator=" & operatorPath & _
             "|SharePoint=" & spec.PathSharePointRoot & _
             "|SharePointPackage=" & sharePointReport
    SetupTesterStation = True
    GoTo CleanExit

FailSoft:
    SetupTesterStation = False
    If Len(report) = 0 Then report = "SetupTesterStation failed."
    LogDiagnosticEvent "TESTER-SETUP", modTesterCommands.DiagnosticText(report)
    GoTo CleanExit

FailSetup:
    report = "SetupTesterStation failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastTesterSetupReport = report
    mLastTesterSharePointRoot = spec.PathSharePointRoot
    modTesterCommands.LeaveProvisioningRoot priorRootOverride
End Function

Public Function ProvisionTesterAuth(ByRef spec As TesterProvisionSpec, _
                                    Optional ByRef report As String = "") As Boolean
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userRow As Long
    Dim capabilitiesReady As Boolean
    Dim hashReady As Boolean
    Dim openedTransient As Boolean
    Dim displayName As String

    On Error GoTo FailProvision

    NormalizeTesterSetupSpec spec
    If Not modTesterCommands.IsWriting() Then ApplyTesterSetupDefaults spec

    authPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    Set wbAuth = OpenWorkbookForWriteTesterSetup(authPath, openedTransient, report)
    If wbAuth Is Nothing Then GoTo FailSoft
    If Not modAuth.EnsureAuthSchema(wbAuth, spec.WarehouseId, "svc_processor", report) Then GoTo FailSoft

    Set loUsers = FindTableByNameTesterSetup(wbAuth, "tblUsers")
    Set loCaps = FindTableByNameTesterSetup(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        report = "Auth tables were not available for tester provisioning."
        GoTo FailSoft
    End If

    userRow = EnsureUserRowTesterSetup(loUsers, spec.UserId)
    displayName = spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "UserId", spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "DisplayName", displayName
    SetTableCellTesterSetup loUsers, userRow, "Status", "Active"

    If SafeTrimTesterSetup(GetTableCellTesterSetup(loUsers, userRow, "PinHash")) = "" Then
        SetTableCellTesterSetup loUsers, userRow, "PinHash", spec.PinHash
    End If

    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "RECEIVE_POST", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "RECEIVE_VIEW", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "SHIP_POST", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "PROD_POST", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "READMODEL_REFRESH", spec.WarehouseId, spec.StationId
    DeactivateCapabilityTesterSetup loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId

    hashReady = (SafeTrimTesterSetup(GetTableCellTesterSetup(loUsers, userRow, "PinHash")) <> "")
    capabilitiesReady = _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "RECEIVE_POST", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "RECEIVE_VIEW", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "SHIP_POST", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "PROD_POST", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "READMODEL_REFRESH", spec.WarehouseId, spec.StationId) And _
        Not CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId)

    wbAuth.Save
    ProvisionTesterAuth = (hashReady And capabilitiesReady)
    If ProvisionTesterAuth Then
        report = "OK"
    Else
        report = "Tester auth verification failed after provisioning."
    End If
    GoTo CleanExit

FailSoft:
    ProvisionTesterAuth = False
    If Len(report) = 0 Then report = "ProvisionTesterAuth failed."
    LogDiagnosticEvent "TESTER-SETUP", modTesterCommands.DiagnosticText(report)
    GoTo CleanExit

FailProvision:
    report = "ProvisionTesterAuth failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbAuth, openedTransient
End Function

Public Function VerifyReceivingWorkbook(ByVal pathIn As String, _
                                        Optional ByRef report As String = "") As Boolean
    Dim wb As Workbook
    Dim openedTransient As Boolean

    On Error GoTo FailVerify

    pathIn = Trim$(pathIn)
    If pathIn = "" Then
        report = "Receiving workbook path is required."
        GoTo FailSoft
    End If
    If Not FileExistsTesterSetup(pathIn) Then
        report = "Receiving workbook not found: " & pathIn
        GoTo FailSoft
    End If

    Set wb = FindOpenWorkbookByPathTesterSetup(pathIn)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(Filename:=pathIn, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        report = "Receiving workbook could not be opened."
        GoTo FailSoft
    End If

    If ReceivingWorkbookHasCanonicalSurfacesTesterSetup(wb) Or ReceivingWorkbookHasAliasSurfacesTesterSetup(wb) Then
        VerifyReceivingWorkbook = True
        report = "OK"
        GoTo CleanExit
    End If

    report = "Receiving workbook is missing required sheets or tables."
    GoTo FailSoft

FailSoft:
    VerifyReceivingWorkbook = False
    If Len(report) = 0 Then report = "VerifyReceivingWorkbook failed."
    LogDiagnosticEvent "TESTER-SETUP", modTesterCommands.DiagnosticText(report)
    GoTo CleanExit

FailVerify:
    report = "VerifyReceivingWorkbook failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wb, openedTransient
End Function

Public Function DetectSharePointRoot(Optional ByVal warehouseId As String = "WH1") As String
    Dim candidate As String
    Dim searchRoots As Collection
    Dim rootPath As Variant
    Dim fso As Object
    Dim folderObj As Object
    Dim rootFolder As Object

    On Error GoTo FailDetect
    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then warehouseId = TESTER_DEFAULT_WAREHOUSE_ID

    If Trim$(mTesterSharePointRootOverride) <> "" Then
        DetectSharePointRoot = NormalizeFolderPathTesterSetup(mTesterSharePointRootOverride, False)
        Exit Function
    End If

    candidate = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    If CandidateLooksLikeSharePointRootTesterSetup(candidate, warehouseId) Then
        DetectSharePointRoot = NormalizeFolderPathTesterSetup(candidate, False)
        Exit Function
    End If

    Set searchRoots = New Collection
    AddUniquePathTesterSetup searchRoots, Environ$("OneDriveCommercial")
    AddUniquePathTesterSetup searchRoots, Environ$("OneDrive")
    AddUniquePathTesterSetup searchRoots, Environ$("USERPROFILE")

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each rootPath In searchRoots
        candidate = NormalizeFolderPathTesterSetup(CStr(rootPath), False)
        If CandidateLooksLikeSharePointRootTesterSetup(candidate, warehouseId) Then
            DetectSharePointRoot = candidate
            Exit Function
        End If

        If candidate <> "" And FolderExistsTesterSetup(candidate) Then
            On Error Resume Next
            Set rootFolder = fso.GetFolder(candidate)
            On Error GoTo FailDetect
            If Not rootFolder Is Nothing Then
                For Each folderObj In rootFolder.SubFolders
                    If CandidateLooksLikeSharePointRootTesterSetup(folderObj.Path, warehouseId) Then
                        DetectSharePointRoot = NormalizeFolderPathTesterSetup(folderObj.Path, False)
                        Exit Function
                    End If
                Next folderObj
            End If
            Set rootFolder = Nothing
        End If
    Next rootPath
    Exit Function

FailDetect:
    DetectSharePointRoot = vbNullString
End Function

Public Function GetLastTesterSetupReport() As String
    GetLastTesterSetupReport = mLastTesterSetupReport
End Function

Public Function GetLastTesterOperatorWorkbookPath() As String
    GetLastTesterOperatorWorkbookPath = mLastTesterOperatorWorkbookPath
End Function

Public Function GetLastTesterSharePointRoot() As String
    GetLastTesterSharePointRoot = mLastTesterSharePointRoot
End Function

Public Sub SetTesterSharePointRootOverride(ByVal rootPath As String)
    mTesterSharePointRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearTesterSharePointRootOverride()
    mTesterSharePointRootOverride = vbNullString
End Sub

Private Function SeedTesterScenarioInventory(ByRef spec As TesterProvisionSpec, _
                                             ByRef report As String) As Boolean
    Dim inventoryWb As Workbook
    Dim inventoryPath As String
    Dim openedTransient As Boolean
    Dim seedPayload As Object
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim tempAdminGranted As Boolean
    Dim tempAdminReport As String

    On Error GoTo FailSeed

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(spec.WarehouseId, Nothing)
    modTesterCommands.RegisterOwned inventoryWb
    If inventoryWb Is Nothing Then
        inventoryPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb"
        Set inventoryWb = OpenWorkbookForWriteTesterSetup(inventoryPath, openedTransient, report)
        If inventoryWb Is Nothing Then
            If Len(report) = 0 Then report = "Canonical inventory workbook could not be resolved."
            GoTo FailSoft
        End If
        If Not EnsureInventorySchemaBridge(inventoryWb, report) Then GoTo FailSoft
    End If

    If TesterSeedAlreadyPresentTesterSetup(inventoryWb) Then
        SeedTesterScenarioInventory = True
        report = "SKIPPED"
        GoTo CleanExit
    End If

    If Not EnsureTemporaryAdminCapabilityTesterSetup(spec, tempAdminGranted, tempAdminReport) Then
        report = tempAdminReport
        GoTo FailSoft
    End If

    If Not modConfig.LoadConfig(spec.WarehouseId, spec.StationId) Then
        report = "Config load failed: " & modConfig.Validate()
        GoTo FailSoft
    End If
    If Not modAuth.LoadAuth(spec.WarehouseId) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        GoTo FailSoft
    End If

    Set seedPayload = modRoleEventWriter.CreatePayloadItem(1, TESTER_SEED_SKU, TESTER_SEED_QTY, TESTER_SEED_LOCATION, "Tester setup seed", "IMPORT")
    seedPayload("Description") = TESTER_SEED_DESCRIPTION
    seedPayload("Item_Code") = TESTER_SEED_SKU
    seedPayload("Item") = TESTER_SEED_SKU
    seedPayload("UOM") = "EA"
    payloadJson = modRoleEventWriter.BuildPayloadJson(seedPayload)

    If Not modRoleEventWriter.QueueMigrationSeedEvent(spec.WarehouseId, spec.StationId, spec.UserId, payloadJson, "", "TESTER_SETUP_SEED", 0, Nothing, eventIdOut, queueError, "") Then
        report = "QueueMigrationSeedEvent failed: " & queueError
        GoTo FailSoft
    End If

    If modProcessor.RunBatch(spec.WarehouseId, 0, batchReport) < 1 Then
        If Not TesterSeedAlreadyPresentTesterSetup(inventoryWb) Then
            report = "Processor did not apply the tester seed. " & batchReport
            GoTo FailSoft
        End If
    End If

    If Not TesterSeedAlreadyPresentTesterSetup(inventoryWb) Then
        report = "Tester seed SKU was not present after processor run."
        GoTo FailSoft
    End If

    SeedTesterScenarioInventory = True
    report = "SEEDED"
    GoTo CleanExit

FailSoft:
    SeedTesterScenarioInventory = False
    If Len(report) = 0 Then report = "SeedTesterScenarioInventory failed."
    LogDiagnosticEvent "TESTER-SETUP", modTesterCommands.DiagnosticText(report)
    GoTo CleanExit

FailSeed:
    report = "SeedTesterScenarioInventory failed: " & Err.Description
    Resume FailSoft

CleanExit:
    RevokeAdminCapabilityTesterSetup spec
    CloseWorkbookIfTransientTesterSetup inventoryWb, openedTransient
End Function

Private Function CreateOrVerifyReceivingWorkbook(ByRef spec As TesterProvisionSpec, _
                                                 ByVal operatorPath As String, _
                                                 ByRef report As String) As Boolean
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim refreshReport As String
    Dim parentFolder As String
    Dim prevEvents As Boolean
    Dim prevDisplayAlerts As Boolean

    On Error GoTo FailCreate
    prevEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts

    parentFolder = GetParentFolderTesterSetup(operatorPath)
    If parentFolder = "" Then
        report = "Operator workbook parent folder could not be resolved."
        GoTo FailSoft
    End If
    EnsureFolderRecursiveTesterSetup parentFolder

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = FindOpenWorkbookByPathTesterSetup(operatorPath)
    If wb Is Nothing And FileExistsTesterSetup(operatorPath) Then
        Set wb = Application.Workbooks.Open(Filename:=operatorPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        openedTransient = Not wb Is Nothing
    End If
    If openedTransient Then modTesterCommands.RegisterOwned wb
    If wb Is Nothing Then
        report = "Operator workbook could not be created."
        GoTo FailSoft
    End If

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo FailSoft
    RemoveNonReceivingOperatorSheetsTesterSetup wb
    Call modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wb, spec.WarehouseId, "LOCAL", refreshReport)

    If Trim$(wb.FullName) = "" Then
        wb.SaveAs Filename:=operatorPath, FileFormat:=52
    ElseIf StrComp(wb.FullName, operatorPath, vbTextCompare) <> 0 Then
        wb.SaveAs Filename:=operatorPath, FileFormat:=52
    Else
        wb.Save
    End If

    If Not VerifyReceivingWorkbook(operatorPath, report) Then GoTo FailSoft

    mLastTesterOperatorWorkbookPath = operatorPath
    CreateOrVerifyReceivingWorkbook = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    CreateOrVerifyReceivingWorkbook = False
    If Len(report) = 0 Then report = "CreateOrVerifyReceivingWorkbook failed."
    LogDiagnosticEvent "TESTER-SETUP", modTesterCommands.DiagnosticText(report)
    GoTo CleanExit

FailCreate:
    report = "CreateOrVerifyReceivingWorkbook failed: " & Err.Description
    Resume FailSoft

CleanExit:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEvents
    CloseWorkbookIfTransientTesterSetup wb, openedTransient
End Function

Private Function UpdateConfigSharePointRoot(ByRef spec As TesterProvisionSpec, _
                                            ByRef report As String) As Boolean
    Dim configPath As String
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim rowIndex As Long
    Dim openedTransient As Boolean

    On Error GoTo FailUpdate

    configPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    Set wbCfg = OpenWorkbookForWriteTesterSetup(configPath, openedTransient, report)
    If wbCfg Is Nothing Then GoTo FailSoft
    If Not modConfig.EnsureConfigSchema(wbCfg, spec.WarehouseId, spec.StationId, report) Then GoTo FailSoft

    Set loWh = FindTableByNameTesterSetup(wbCfg, "tblWarehouseConfig")
    If loWh Is Nothing Or loWh.DataBodyRange Is Nothing Then
        report = "Warehouse config table not available."
        GoTo FailSoft
    End If

    rowIndex = FindRowByValueTesterSetup(loWh, "WarehouseId", spec.WarehouseId)
    If rowIndex = 0 Then rowIndex = 1

    SetTableCellTesterSetup loWh, rowIndex, "WarehouseId", spec.WarehouseId
    If SafeTrimTesterSetup(GetTableCellTesterSetup(loWh, rowIndex, "WarehouseName")) = "" Then
        SetTableCellTesterSetup loWh, rowIndex, "WarehouseName", spec.WarehouseId
    End If
    SetTableCellTesterSetup loWh, rowIndex, "PathDataRoot", NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    SetTableCellTesterSetup loWh, rowIndex, "PathSharePointRoot", NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
    wbCfg.Save

    UpdateConfigSharePointRoot = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    UpdateConfigSharePointRoot = False
    If Len(report) = 0 Then report = "UpdateConfigSharePointRoot failed."
    LogDiagnosticEvent "TESTER-SETUP", modTesterCommands.DiagnosticText(report)
    GoTo CleanExit

FailUpdate:
    report = "UpdateConfigSharePointRoot failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbCfg, openedTransient
End Function

Private Function CreateTesterRuntimeArtifactsTesterSetup(ByRef spec As TesterProvisionSpec, _
                                                         ByVal rootPath As String, _
                                                         ByRef report As String) As Boolean
    Dim configPath As String
    Dim authPath As String
    Dim inventoryPath As String
    Dim outboxPath As String
    Dim snapshotPath As String
    Dim capabilityOut As String
    Dim wbInventory As Workbook
    Dim wbOutbox As Workbook
    Dim wbConfig As Workbook
    Dim configOpenedTransient As Boolean
    Dim inventoryOpenedTransient As Boolean
    Dim outboxOpenedTransient As Boolean

    On Error GoTo FailCreate

    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    If rootPath = "" Then
        report = "Warehouse hub path could not be resolved."
        Exit Function
    End If
    If IsUncPathTesterSetup(rootPath) And Not FolderExistsTesterSetup(rootPath) Then
        report = "Warehouse hub path is not accessible: " & rootPath
        Exit Function
    End If

    EnsureFolderRecursiveTesterSetup rootPath
    EnsureFolderRecursiveTesterSetup rootPath & "\inbox"
    EnsureFolderRecursiveTesterSetup rootPath & "\outbox"
    EnsureFolderRecursiveTesterSetup rootPath & "\snapshots"
    EnsureFolderRecursiveTesterSetup rootPath & "\config"
    If Not modTesterCommands.CheckPhase() Then GoTo FailSoft
    EnsureFolderRecursiveTesterSetup rootPath & "\auth"

    configPath = rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    authPath = rootPath & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    inventoryPath = rootPath & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb"
    outboxPath = rootPath & "\" & spec.WarehouseId & ".Outbox.Events.xlsb"
    snapshotPath = rootPath & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb"

    If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, spec.UserId, rootPath & "\inbox\", "ADMIN", configPath, rootPath, report) Then GoTo FailSoft
    configOpenedTransient = (FindOpenWorkbookByPathTesterSetup(configPath) Is Nothing)
    Set wbConfig = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(spec.WarehouseId, spec.StationId, rootPath, report)
    If configOpenedTransient Then modTesterCommands.RegisterOwned wbConfig
    If wbConfig Is Nothing Then GoTo FailSoft

    If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.UserId, spec.UserId, "ADMIN", authPath, "svc_processor", capabilityOut, report) Then GoTo FailSoft

    inventoryOpenedTransient = (FindOpenWorkbookByPathTesterSetup(inventoryPath) Is Nothing)
    Set wbInventory = ResolveInventoryWorkbookBridge(spec.WarehouseId)
    If inventoryOpenedTransient Then modTesterCommands.RegisterOwned wbInventory
    If wbInventory Is Nothing Then
        report = "Inventory workbook not resolved."
        GoTo FailSoft
    End If
    If Not EnsureInventorySchemaBridge(wbInventory, report) Then GoTo FailSoft
    If Not wbInventory.ReadOnly Then wbInventory.Save

    If Not GenerateWarehouseSnapshot(spec.WarehouseId, wbInventory, snapshotPath, Nothing, report) Then GoTo FailSoft

    outboxOpenedTransient = (FindOpenWorkbookByPathTesterSetup(outboxPath) Is Nothing)
    Set wbOutbox = ResolveOutboxWorkbook(spec.WarehouseId, Nothing, True)
    If outboxOpenedTransient Then modTesterCommands.RegisterOwned wbOutbox
    If wbOutbox Is Nothing Then
        report = "Outbox workbook not resolved."
        GoTo FailSoft
    End If
    If Not EnsureOutboxSchema(wbOutbox, report) Then GoTo FailSoft
    If Not wbOutbox.ReadOnly Then wbOutbox.Save

    CreateTesterRuntimeArtifactsTesterSetup = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    If Len(report) = 0 Then report = "CreateTesterRuntimeArtifacts failed."
    GoTo CleanExit

FailCreate:
    report = "CreateTesterRuntimeArtifacts failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbOutbox, outboxOpenedTransient
    CloseWorkbookIfTransientTesterSetup wbInventory, inventoryOpenedTransient
    CloseWorkbookIfTransientTesterSetup wbConfig, configOpenedTransient
End Function

Private Function EnsureTesterSharePointPackage(ByRef spec As TesterProvisionSpec, _
                                               ByRef report As String) As Boolean
    Dim sharePointRoot As String
    Dim addinsRoot As String
    Dim sourceAddinsRoot As String
    Dim requiredAddins As Variant
    Dim addinName As Variant
    Dim sourcePath As String
    Dim targetPath As String
    Dim copiedCount As Long

    On Error GoTo FailPackage

    sharePointRoot = NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
    If sharePointRoot = "" Then
        report = "SKIPPED|PathSharePointRoot not configured."
        EnsureTesterSharePointPackage = True
        Exit Function
    End If

    EnsureFolderRecursiveTesterSetup sharePointRoot
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\Addins"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\Events"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\Snapshots"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\TesterPackage"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\TesterPackage\" & spec.WarehouseId

    sourceAddinsRoot = ResolveCurrentAddinsRootTesterSetup()
    If sourceAddinsRoot = "" Then
        report = "Current add-ins folder could not be resolved for tester package publish."
        Exit Function
    End If

    addinsRoot = sharePointRoot & "\Addins"
    requiredAddins = RequiredTesterAddinNamesTesterSetup()
    For Each addinName In requiredAddins
        sourcePath = sourceAddinsRoot & "\" & CStr(addinName)
        targetPath = addinsRoot & "\" & CStr(addinName)
        If Not FileExistsTesterSetup(sourcePath) Then
            report = "Source add-in missing: " & sourcePath
            Exit Function
        End If
        If CopyFileIfNeededTesterSetup(sourcePath, targetPath) Then copiedCount = copiedCount + 1
    Next addinName

    report = "OK|Addins=" & addinsRoot & "|Copied=" & CStr(copiedCount)
    EnsureTesterSharePointPackage = True
    Exit Function

FailPackage:
    report = "EnsureTesterSharePointPackage failed: " & Err.Description
End Function

Private Sub ResetTesterSetupState()
    mLastTesterSetupReport = vbNullString
    mLastTesterOperatorWorkbookPath = vbNullString
    mLastTesterSharePointRoot = vbNullString
End Sub

' Explicit headless fixture entry; interactive callers use modTesterCommands.
Public Function SetupValues(ByRef userId As String, ByRef credential As String, ByRef warehouse As String, _
    ByRef station As String, ByRef root As String, ByRef publishRoot As String) As Boolean
    Dim spec As TesterProvisionSpec
    spec.UserId = userId: spec.PinHash = credential: spec.WarehouseId = warehouse
    spec.StationId = station: spec.PathLocal = root: spec.PathSharePointRoot = publishRoot
    SetupValues = SetupTesterStation(spec)
    userId = spec.UserId: credential = spec.PinHash: warehouse = spec.WarehouseId
    station = spec.StationId: root = spec.PathLocal: publishRoot = spec.PathSharePointRoot
End Function
