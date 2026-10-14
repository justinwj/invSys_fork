Attribute VB_Name = "modTesterOwnership"
Option Explicit
Option Private Module

Public Const RECEIPT_NAME As String = "invSys.TesterOwnership.json"

' Receipt data is local provenance, never an authorization grant or activity payload.
Public Function Resolve(ByVal context As String, ByRef spec As TesterProvisionSpec, ByVal reserve As Boolean) As Object
    Dim root As String, publish As String, path As String, record As Object, fso As Object
    Dim source As Object
    On Error GoTo Rejected
    Set source = modTesterAuthority.Entry(context)
    If source Is Nothing Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = CanonicalRoot(spec.PathLocal): publish = CanonicalRoot(spec.PathSharePointRoot)
    If root = "" Then Exit Function
    If spec.PathSharePointRoot <> "" And publish = "" Then Exit Function
    If Not SafeIdentity(spec.WarehouseId) Or Not SafeIdentity(spec.StationId) Then Exit Function
    If Overlaps(root, CStr(source("SourceRoot"))) Then Exit Function
    If publish <> "" Then
        If Overlaps(root, publish) Or Overlaps(publish, CStr(source("SourceRoot"))) Then Exit Function
    End If
    spec.PathLocal = root: spec.PathSharePointRoot = publish
    path = root & "\" & RECEIPT_NAME
    If fso.FileExists(path) Then
        Set record = ReadReceipt(path)
        If record Is Nothing Then Exit Function
        If Not Matches(record, source, spec) Then Exit Function
        If publish <> "" Then
            If Not SameReceipt(record, publish & "\" & RECEIPT_NAME) Then
                If reserve Then
                    If Not EmptyRoot(publish) Then Exit Function
                    If Not modTesterAuthority.Check(context) Then Exit Function
                    WriteNewReceipt publish, record
                ElseIf Not PublicationAlreadyClean(record) Then
                    Exit Function
                End If
            End If
        End If
    Else
        If Not reserve Or Not EmptyRoot(root) Then Exit Function
        If publish <> "" Then
            If Not EmptyRoot(publish) Then Exit Function
        End If
        Set record = CreateObject("Scripting.Dictionary")
        record.Add "Version", CLng(1)
        record.Add "ReservationId", modSystemIdentity.NewId()
        record.Add "Root", root: record.Add "PublishRoot", publish
        record.Add "Warehouse", spec.WarehouseId: record.Add "Station", spec.StationId
        record.Add "SourceWarehouse", CStr(source("SourceWarehouse"))
        record.Add "CreatorUser", CStr(source("CreatorUser"))
        record.Add "CreatorGeneration", CLng(source("CreatorGeneration"))
        record.Add "Artifacts", Artifacts(spec)
        record.Add "Directories", Directories(spec)
        ' Durable reservation precedes every workbook, directory tree and publication write.
        If Not modTesterAuthority.Check(context) Then Exit Function
        WriteNewReceipt root, record
        If publish <> "" Then
            If Not modTesterAuthority.Check(context) Then Exit Function
            WriteNewReceipt publish, record
        End If
    End If
    Set Resolve = record
Rejected:
End Function

Public Function Matches(ByVal record As Object, ByVal source As Object, ByRef spec As TesterProvisionSpec) As Boolean
    On Error GoTo Rejected
    If record.Count <> 11 Then Exit Function
    If record("Version") <> 1 Or Len(CStr(record("ReservationId"))) <> 36 Then Exit Function
    If record("SourceWarehouse") <> source("SourceWarehouse") Then Exit Function
    If VarType(record("CreatorUser")) <> vbString Or record("CreatorUser") = "" Then Exit Function
    If VarType(record("CreatorGeneration")) <> vbLong Then Exit Function
    If record("CreatorGeneration") < 0 Then Exit Function
    If record("Root") <> spec.PathLocal Or record("PublishRoot") <> spec.PathSharePointRoot Then Exit Function
    If record("Warehouse") <> spec.WarehouseId Or record("Station") <> spec.StationId Then Exit Function
    If modTrainingJson.EncodeObject(ArtifactMap(record("Artifacts"))) <> modTrainingJson.EncodeObject(ArtifactMap(Artifacts(spec))) Then Exit Function
    If modTrainingJson.EncodeObject(ArtifactMap(record("Directories"))) <> modTrainingJson.EncodeObject(ArtifactMap(Directories(spec))) Then Exit Function
    Matches = True
Rejected:
End Function

Private Function ArtifactMap(ByVal items As Collection) As Object
    Dim result As Object, item As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each item In items
        result.Add CStr(item), True
    Next item
    Set ArtifactMap = result
End Function

Public Function Unchanged(ByVal record As Object, Optional ByVal cleaning As Boolean = False) As Boolean
    Dim item As Variant
    On Error GoTo Rejected
    If CanonicalRoot(CStr(record("Root"))) <> record("Root") Then Exit Function
    If Not SameReceipt(record, CStr(record("Root")) & "\" & RECEIPT_NAME) Then Exit Function
    If record("PublishRoot") <> "" Then
        If CanonicalRoot(CStr(record("PublishRoot"))) <> record("PublishRoot") Then Exit Function
        If Not SameReceipt(record, CStr(record("PublishRoot")) & "\" & RECEIPT_NAME) Then
            If Not cleaning Or Not PublicationAlreadyClean(record) Then Exit Function
        End If
    End If
    For Each item In record("Directories")
        If CanonicalRoot(CStr(item)) <> item Then Exit Function
    Next item
    Unchanged = True
Rejected:
End Function

Private Function PublicationAlreadyClean(ByVal record As Object) As Boolean
    Dim fso As Object, item As Variant, root As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    root = CStr(record("PublishRoot"))
    If fso.FileExists(root & "\" & RECEIPT_NAME) Then Exit Function
    For Each item In record("Artifacts")
        If Left$(CStr(item), Len(root) + 1) = root & "\" Then
            If fso.FileExists(CStr(item)) Then Exit Function
        End If
    Next item
    PublicationAlreadyClean = True
End Function

Private Function SameReceipt(ByVal record As Object, ByVal path As String) As Boolean
    Dim other As Object
    Set other = ReadReceipt(path)
    If other Is Nothing Then Exit Function
    SameReceipt = (modTrainingJson.EncodeObject(other) = modTrainingJson.EncodeObject(record))
End Function

Private Function ReadReceipt(ByVal path As String) As Object
    Dim fso As Object, stream As Object, text As String
    On Error GoTo Rejected
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    If fso.GetFile(path).Size > 32768 Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False)
    text = stream.ReadAll: stream.Close
    Set ReadReceipt = modTrainingJson.DecodeObject(text)
Rejected:
End Function

Private Sub WriteNewReceipt(ByVal root As String, ByVal record As Object)
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(root) Then fso.CreateFolder root
    ' CreateTextFile(overwrite=False) reserves without replacing foreign ownership.
    Set stream = fso.CreateTextFile(root & "\" & RECEIPT_NAME, False, False)
    stream.Write modTrainingJson.EncodeObject(record): stream.Close
    If Not SameReceipt(record, root & "\" & RECEIPT_NAME) Then Err.Raise 5
End Sub

Public Function Artifacts(ByRef spec As TesterProvisionSpec) As Collection
    Dim items As New Collection, suffix As Variant, name As Variant
    For Each suffix In Array(".invSys.Config.xlsb", ".invSys.Auth.xlsb", ".invSys.Data.Inventory.xlsb", _
        ".invSys.Snapshot.Inventory.xlsb", ".Outbox.Events.xlsb", ".Receiving.Operator.xlsm")
        items.Add spec.PathLocal & "\" & spec.WarehouseId & CStr(suffix)
    Next suffix
    For Each name In Array("Receiving", "Shipping", "Production")
        items.Add spec.PathLocal & "\inbox\invSys.Inbox." & CStr(name) & "." & spec.StationId & ".xlsb"
    Next name
    items.Add spec.PathLocal & "\invSys.Publish.log"
    If spec.PathSharePointRoot <> "" Then
        For Each name In RequiredTesterAddinNamesTesterSetup()
            items.Add spec.PathSharePointRoot & "\Addins\" & CStr(name)
        Next name
    End If
    Set Artifacts = items
End Function

Public Function Directories(ByRef spec As TesterProvisionSpec) As Collection
    Dim items As New Collection, suffix As Variant
    For Each suffix In Array("auth", "config", "snapshots", "outbox", "inbox")
        items.Add spec.PathLocal & "\" & CStr(suffix)
    Next suffix
    If spec.PathSharePointRoot <> "" Then
        items.Add spec.PathSharePointRoot & "\TesterPackage\" & spec.WarehouseId
        For Each suffix In Array("TesterPackage", "Events", "Snapshots", "Addins")
            items.Add spec.PathSharePointRoot & "\" & CStr(suffix)
        Next suffix
    End If
    Set Directories = items
End Function

Public Function CanonicalRoot(ByVal path As String) As String
    Dim fso As Object, canonical As String, current As String
    On Error GoTo Rejected
    path = Replace$(Trim$(path), "/", "\")
    If path = "" Or InStr(path, "..") Or InStr(path, "~") Then Exit Function
    If InStr(path, "*") Or InStr(path, "?") Then Exit Function
    If Not (Mid$(path, 2, 2) = ":\" Or Left$(path, 2) = "\\") Then Exit Function
    If InStr(3, path, ":") Then Exit Function
    Do While Right$(path, 1) = "\": path = Left$(path, Len(path) - 1): Loop
    Set fso = CreateObject("Scripting.FileSystemObject")
    canonical = fso.GetAbsolutePathName(path)
    If StrComp(path, canonical, vbTextCompare) <> 0 Then Exit Function
    If fso.GetParentFolderName(canonical) = "" Then Exit Function
    current = canonical
    Do While current <> ""
        If fso.FileExists(current) Then Exit Function
        If fso.FolderExists(current) Then
            If (fso.GetFolder(current).Attributes And 1024) <> 0 Then Exit Function
        End If
        current = fso.GetParentFolderName(current)
    Loop
    CanonicalRoot = canonical
Rejected:
End Function

Private Function EmptyRoot(ByVal path As String) As Boolean
    Dim fso As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        EmptyRoot = fso.FolderExists(fso.GetParentFolderName(path))
    Else
        Set folder = fso.GetFolder(path)
        EmptyRoot = (folder.Files.Count = 0 And folder.SubFolders.Count = 0)
    End If
End Function

Private Function Overlaps(ByVal first As String, ByVal second As String) As Boolean
    first = LCase$(first) & "\": second = LCase$(second) & "\"
    Overlaps = (Left$(first, Len(second)) = second Or Left$(second, Len(first)) = first)
End Function

Private Function SafeIdentity(ByVal value As String) As Boolean
    Dim index As Long, ch As String
    If value = "" Or value <> Trim$(value) Then Exit Function
    If value = "." Or value = ".." Or Right$(value, 1) = "." Then Exit Function
    For index = 1 To Len(value)
        ch = Mid$(value, index, 1)
        If AscW(ch) >= 0 And AscW(ch) < 32 Then Exit Function
        If InStr(1, "\/:*?""<>|", ch, vbBinaryCompare) > 0 Then Exit Function
    Next index
    SafeIdentity = True
End Function
