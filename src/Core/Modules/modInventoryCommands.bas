Attribute VB_Name = "modInventoryCommands"
Option Explicit

' D8-INV-01 interactive facade. Handles are transient and grant no headless authority.
Private mForms As cCapturedFormStore

Public Function Capture(ByVal workbookName As String) As String
    Dim book As Workbook, target As WarehouseTarget, entry As Object
    Dim context As String, key As String, report As String
    On Error GoTo Failed
    Set book = Application.Workbooks(workbookName)
    If book.IsAddin Or book.Path = "" Then Exit Function
    context = modAuthUserCommands.CaptureContext()
    If context = "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If modAggregationFiles.SameFile(book.FullName, target.ConfigPath) Or _
       modAggregationFiles.SameFile(book.FullName, target.AuthPath) Then Exit Function
    If LCase$(book.Name) Like "*.invsys.data.*" Or _
       LCase$(book.Name) Like "*.invsys.snapshot.*" Or _
       LCase$(book.Name) = "invsys.global.inventorysnapshot.xlsb" Then Exit Function
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then Exit Function
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "Context", context: entry.Add "Workbook", book
    entry.Add "WorkbookPath", book.FullName
    entry.Add "Warehouse", target.WarehouseId: entry.Add "Station", target.StationId
    entry.Add "User", modAuth.GetCurrentUserId(): entry.Add "Root", target.RuntimeRoot
    entry.Add "Config", target.ConfigPath: entry.Add "Auth", target.AuthPath
    entry.Add "Feed", modConfig.GetString("PathDataRoot", "")
    entry.Add "Expired", False
    If mForms Is Nothing Then Set mForms = New cCapturedFormStore
    key = modSystemIdentity.NewId(): mForms.Add key, entry
    If Not CheckForm(key, report) Then Release key: Exit Function
    Capture = key
Failed:
End Function

' Declared cross-package values: a primitive handle and a three-string array.
Public Function SourceValues(ByVal key As String) As Variant
    Dim report As String, entry As Object
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key)
    SourceValues = Array(CStr(entry("Warehouse")), CStr(entry("Station")), CStr(entry("User")))
End Function

Public Function CheckForm(ByVal key As String, ByRef report As String) As Boolean
    Dim entry As Object, original As Workbook, book As Workbook, found As Boolean
    Dim root As String
    report = "The invSys session or inventory source changed. Reopen Add/Edit Inventory."
    On Error GoTo Refused
    If mForms Is Nothing Then Exit Function
    Set entry = mForms.Item(key)
    If entry Is Nothing Then Exit Function
    If entry("Expired") Then Exit Function
    If Not modAuth.IsSignedIn() Then
        report = "Current invSys user is not signed in."
        GoTo Refused
    End If
    If entry("Context") <> modAuthUserCommands.CaptureContext() Then GoTo Refused
    Set original = entry("Workbook")
    For Each book In Application.Workbooks
        If book Is original Then found = True
    Next book
    If Not found Then GoTo Refused
    If Not modAggregationFiles.SameFile(original.FullName, CStr(entry("WorkbookPath"))) Then GoTo Refused
    root = modRuntimeWorkbooks.GetCoreDataRootOverride()
    If root <> "" Then
        If StrComp(modAggregationFiles.Canonical(root), _
                   modAggregationFiles.Canonical(CStr(entry("Root"))), vbTextCompare) <> 0 Then GoTo Refused
    End If
    If modAggregationFiles.Dirty(CStr(entry("Config"))) Or _
       modAggregationFiles.Dirty(CStr(entry("Auth"))) Then GoTo Refused
    If Not modAuth.LoadAuth(CStr(entry("Warehouse"))) Then GoTo Refused
    If Not modAuth.CanPerform("ADMIN_MAINT", modAuth.GetCurrentUserId(), _
            CStr(entry("Warehouse")), CStr(entry("Station")), "INVENTORY_COMMAND") Then
        report = "Current user lacks ADMIN_MAINT for the captured warehouse."
        GoTo Refused
    End If
    If Not modConfig.LoadConfig(CStr(entry("Warehouse")), CStr(entry("Station"))) Then GoTo Refused
    If modConfig.GetString("PathDataRoot", "") <> entry("Feed") Then GoTo Refused
    If entry("Context") <> modAuthUserCommands.CaptureContext() Then GoTo Refused
    report = "": CheckForm = True
    Exit Function
Refused:
    If Not entry Is Nothing Then entry("Expired") = True
End Function

Public Sub Release(ByVal key As String)
    If Not mForms Is Nothing Then mForms.Remove key
End Sub
