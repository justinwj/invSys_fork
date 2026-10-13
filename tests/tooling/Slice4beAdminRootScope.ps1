# Generated discovery inputs; real scope assembly, remembered preference read and directory reader.
function Install-AdminRootScope($Probe,$Console,[string]$OriginalRoots) {
    $Probe.InsertLines($Probe.CountOfDeclarationLines+1,@'
Private mScopeActive As Boolean, mScopeBeforeWrite As Boolean
Private mScopeAmbient As String, mScopeLateRoot As String, mScopeLatePath As String
Private mScopeRoots As String, mScopeBook As Workbook
'@)
    $Probe.AddFromString(@'
Public Sub ArmRootReadScope(ByVal root As String, ByVal lateRoot As String, ByVal latePath As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, latePath, vbTextCompare) = 0 Then Err.Raise 5, , "Late root read fixture already open"
    Next wb
    mScopeActive = True: mScopeBeforeWrite = False: mScopeAmbient = root
    mScopeLateRoot = lateRoot: mScopeLatePath = latePath: mScopeRoots = ""
End Sub
Public Function RootReadScopeActive() As Boolean
    RootReadScopeActive = mScopeActive
End Function
Public Function RootScopeInputs() As Object
    Dim known As New Collection
    mScopeBeforeWrite = (mRootWrites = 0)
    known.Add mScopeAmbient & "\"
    Set RootScopeInputs = modAdminDirectoryScope.Capture(mScopeAmbient, mScopeAmbient, mScopeAmbient, known, mScopeAmbient)
End Function
Public Sub OpenLateRootScope()
    If Not mScopeActive Then Exit Sub
    Set mScopeBook = Application.Workbooks.Open(Filename:=mScopeLatePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    mScopeAmbient = mScopeLateRoot
End Sub
Public Sub ObservedRootScope(ByVal roots As Collection)
    Dim item As Variant
    mScopeRoots = ""
    For Each item In roots
        If mScopeRoots <> "" Then mScopeRoots = mScopeRoots & "|"
        mScopeRoots = mScopeRoots & CStr(item)
    Next item
End Sub
Public Function RootScopeMatches(ByVal expected As String) As Boolean
    RootScopeMatches = (StrComp(mScopeRoots, expected, vbBinaryCompare) = 0)
End Function
Public Function RootScopeFaultReached() As Boolean
    RootScopeFaultReached = mScopeBeforeWrite And Not mScopeBook Is Nothing And mScopeAmbient = mScopeLateRoot
End Function
Public Function RootScopeOnlyWarehouse(ByVal warehouseId As String) As Boolean
    Dim lo As ListObject, row As ListRow, value As String, count As Long
    Set lo = Directory()
    For Each row In lo.ListRows
        value = CStr(row.Range.Cells(1, lo.ListColumns("WarehouseId").Index).Value2)
        If value <> "" Then
            If value <> warehouseId Then Exit Function
            count = count + 1
        End If
    Next row
    RootScopeOnlyWarehouse = (count = 1)
End Function
Public Sub CloseRootReadScope()
    If Not mScopeBook Is Nothing Then mScopeBook.Close False
    Set mScopeBook = Nothing: mScopeActive = False
End Sub
'@)
    if(-not $OriginalRoots){throw 'Original packaged discovery method unavailable.'}
    $Console.AddFromString($OriginalRoots.Replace('GetWarehouseDirectoryScanRoots','OriginalRootScanForTest'))
    $start=$Console.ProcStartLine('GetWarehouseDirectoryScanRoots',0);$count=$Console.ProcCountLines('GetWarehouseDirectoryScanRoots',0)
    $Console.DeleteLines($start,$count)
    $Console.InsertLines($start,@'
Public Function GetWarehouseDirectoryScanRoots(Optional ByVal inputs As Object = Nothing) As Collection
    If TestDirectoryColumns.RootReadScopeActive() Then
        Set GetWarehouseDirectoryScanRoots = OriginalRootScanForTest(inputs)
        TestDirectoryColumns.ObservedRootScope GetWarehouseDirectoryScanRoots
    Else
        Set GetWarehouseDirectoryScanRoots = TestDirectoryColumns.ReadRoots()
    End If
End Function
'@)
    $declarations=@(for($line=1;$line -le $Console.CountOfLines;$line++){if($Console.Lines($line,1) -match '^Public Function CaptureWarehouseDirectoryInputs\('){$line}})
    if($declarations.Count -ne 1){throw 'Root capture declaration is not unique.'}
    $Console.InsertLines($declarations[0]+1,@'
    If TestDirectoryColumns.RootReadScopeActive() Then
        Set CaptureWarehouseDirectoryInputs = TestDirectoryColumns.RootScopeInputs()
        Exit Function
    End If
'@)
    $observer=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminActionObservation').CodeModule
    $start=$observer.ProcStartLine('BeginObservation',0);$count=$observer.ProcCountLines('BeginObservation',0)
    $source=$observer.Lines($start,$count)
    $changed=$source.Replace('    context = modActivity.CaptureContext()',"    If controlId = `"ADMIN_WAREHOUSE_ROOT_ADD`" Then TestDirectoryColumns.OpenLateRootScope`r`n    context = modActivity.CaptureContext()")
    if($source -ceq $changed){throw 'Root observation setup boundary unavailable.'}
    $observer.DeleteLines($start,$count);$observer.InsertLines($start,$changed)
}

function Test-AdminRootScope($Fixture,$Other) {
    $available=[bool](Invoke 'RootReadScopeAvailable')
    Check 'WarehouseRoot.Scope.Available' $available
    if(-not $available){return}
    SelectTarget $Fixture
    if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Root scope policy prerequisite unavailable.'}
    $base=Join-Path $runRoot 'root-discovery-scope'
    $ambient=Join-Path $base 'ambient';$remembered=Join-Path $base 'remembered';$added=Join-Path $base 'added';$late=Join-Path $base 'late'
    foreach($path in @($ambient,$remembered,$added,$late)){New-Item -ItemType Directory -Path $path|Out-Null}
    Copy-Item -LiteralPath $Fixture.Config -Destination (Join-Path $ambient 'ScopeA.invSys.Config.xlsb')
    Copy-Item -LiteralPath $Other.Config -Destination (Join-Path $late 'ScopeB.invSys.Config.xlsb')
    [void](Invoke 'SetRootPreference' @($remembered))
    $authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root;$scopePins=Pins $base;$before=@(Get-Slice4beActivityFiles $Fixture)
    [void](Invoke 'ArmRootReadScope' @($ambient,$late,$Other.Config))
    try{
        $facts=([string](Invoke 'InvokeRoot' @($added,$false,'',$false))) -split '\|'
        Check 'WarehouseRoot.Scope.ActualLateConfigAndAmbientChange' ([bool](Invoke 'RootScopeFaultReached'))
        Check 'WarehouseRoot.Scope.FrozenAmbientAndPostWritePreference' ([bool](Invoke 'RootScopeMatches' @((@($ambient,$added,$remembered) -join '|'))))
        Check 'WarehouseRoot.Scope.OwnerOnceAndOriginalFeedback' ($facts.Count -eq 14 -and $facts[0] -ceq '1' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq '1' -and $facts[4] -ceq '1' -and $facts[8] -ceq 'True' -and $facts[11] -ceq '1' -and [int](Invoke 'RootTrackingNotices') -eq 0)
        Check 'WarehouseRoot.Scope.LateWarehouseExcluded' ([bool](Invoke 'RootScopeOnlyWarehouse' @($Fixture.Warehouse)))
        $preserved=$true;foreach($part in @('Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
        Check 'WarehouseRoot.Scope.CustomColumnsPreserved' $preserved
        Check 'WarehouseRoot.Scope.SavedSourcesAndHistoryPreserved' ((Retained $authority) -and (Retained $otherAuthority) -and (Same $scopePins (Pins $base)))
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_WAREHOUSE_ROOT_ADD')
        Check 'WarehouseRoot.Scope.HonestOwnerPair' ($records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($records|Where-Object OutcomeCode -CEQ 'REMEMBERED_REFRESHED').Count -eq 1)
    }finally{
        [void](Invoke 'CloseRootReadScope')
        [void](Invoke 'SetRootPreference' @($Fixture.Root))
        SelectTarget $Fixture
        [void](Invoke 'Invoke')
    }
}
