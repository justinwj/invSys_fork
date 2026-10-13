# D18: actual View Warehouses dispatch; emit assertions, never entered values.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminDirectoryBoundaries.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminDirectoryGuide.ps1')
function Install-AdminDirectoryRecordingProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Item('TestDirectoryColumns').CodeModule
    $probe.InsertLines(2,'Private mLatePath As String, mLateRoot As String, mLateBook As Workbook')
    $probe.AddFromString(@'
Public Function HostWindowFacts() As String
    Dim wb As Workbook, found As Boolean
    For Each wb In Application.Workbooks
        If wb Is mHost Then found = True
    Next wb
    HostWindowFacts = CStr(found) & "|" & CStr(mHost.Windows.Count) & "|" & CStr(mHost Is ThisWorkbook)
End Function
Public Sub ArmLateReadScope(ByVal path As String, ByVal root As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Err.Raise 5, , "Late read-scope fixture already open"
    Next wb
    mLatePath = path: mLateRoot = root
End Sub
Public Sub OpenLateReadScope()
    If mLatePath = "" Then Exit Sub
    Set mLateBook = Application.Workbooks.Open(Filename:=mLatePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    mRoot = mLateRoot
End Sub
Public Function LateReadScopeObserved() As Boolean
    LateReadScopeObserved = Not mLateBook Is Nothing
End Function
Public Sub CloseLateReadScope(ByVal originalRoot As String)
    If Not mLateBook Is Nothing Then mLateBook.Close False
    Set mLateBook = Nothing: mLatePath = "": mLateRoot = "": mRoot = originalRoot
End Sub
Public Sub ProtectDirectory(ByVal value As Boolean)
    If value Then
        mHost.Worksheets("WarehouseDirectory").Name = "Saved_Directory"
        mHost.Protect Structure:=True
    Else
        mHost.Unprotect
        On Error Resume Next
        mHost.Worksheets("Saved_Directory").Name = "WarehouseDirectory"
        On Error GoTo 0
    End If
End Sub
'@)
    if(@($admin.VBComponents|Where-Object Name -CEQ 'modAdminDirectoryActivity').Count){
        $observer=$admin.VBComponents.Item('modAdminActionObservation').CodeModule
        $start=$observer.ProcStartLine('BeginObservation',0);$count=$observer.ProcCountLines('BeginObservation',0)
        $source=$observer.Lines($start,$count)
        $changed=$source.Replace('    context = modActivity.CaptureContext()',"    If controlId = `"ADMIN_DIRECTORY_OPEN`" Then TestDirectoryColumns.OpenLateReadScope`r`n    context = modActivity.CaptureContext()")
        if($source -ceq $changed){throw 'Directory optional-observation seam unavailable.'}
        $observer.DeleteLines($start,$count);$observer.InsertLines($start,$changed)
    }
    $core=$packages['invSys.Core.xlam'].VBProject
    $gate=$core.VBComponents.Add(1);$gate.Name='TestDirectoryActivity'
    $gate.CodeModule.AddFromString(@'
Option Explicit
Public DenialMessages As Long
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function SetPolicy(ByVal mode As String) As Boolean
    Dim context As String, version As Long, wire As String, notice As String
    Dim model As Object, row As Variant, users As Collection, user As Object
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "Capture")
    For Each row In model("Controls")
        If row("ControlId") = "ADMIN_DIRECTORY_OPEN" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    SetPolicy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), notice)
End Function
Public Function Definition(ByVal version As Long) As String
    Dim record As Object
    Set record = modActivityCatalog.Control("ADMIN_DIRECTORY_OPEN", version)
    If Not record Is Nothing Then Definition = modTrainingJson.EncodeObject(record)
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(38)
        Set before = modActivityCatalog.Control(CStr(id), 38)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
Public Function DenialCount() As Long
    DenialCount = DenialMessages
End Function
'@)
    $core.VBComponents.Item('modRoleUiAccess').CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    TestDirectoryActivity.DenialMessages = TestDirectoryActivity.DenialMessages + 1
    MsgBox = vbOK
End Function
'@)
    Install-AdminDirectoryBoundaries $gate.CodeModule $probe
    if($AdminDirectoryGuidePaths){
        . (Join-Path $PSScriptRoot 'Slice4beAdminVerifyGuide.ps1')
        Install-AdminVerifyGuideProbe
    }
}

function Test-AdminDirectoryRecording($Fixture,$Other) {
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestDirectoryActivity.'+$Name) $Values}
    function Invoke([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestDirectoryColumns.'+$Name) $Values}
    $catalogVersion=[int](Probe 'CurrentCatalogVersion')
    function Pins([string]$Root){$result=@{};Get-ChildItem -LiteralPath $Root -Recurse -File|ForEach-Object {$s=[IO.File]::Open($_.FullName,'Open','Read','ReadWrite');try{$result[$_.FullName]=(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}};return $result}
    function Retained($Before){foreach($path in $Before.Keys){if(-not(Test-Path -LiteralPath $path)){return $false};$s=[IO.File]::Open($path,'Open','Read','ReadWrite');try{$hash=(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()};if($hash -cne $Before[$path]){return $false}};return $true}
    SelectTarget $Fixture
    # Establish visible Excel before creating the guide's captured host. The prior
    # preservation fixture has closed its last visible workbook at this point.
    if($AdminDirectoryGuidePaths){$excel.Visible=$true}
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add();$otherCount=$otherBook.Worksheets.Count
    $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Other directory workbook canary'
    try{
        if(-not [bool](Invoke 'Prepare' @($hostBook.Name,$Fixture.Root))){throw 'Directory recording fixture preparation failed; not behavioral RED.'}
        if([string](Invoke 'Invoke') -cne '1|True|True|1|0|0' -or -not [bool](Invoke 'AddCustomColumns')){throw 'Directory recording owner fixture unavailable.'}
        $wire=[string](Probe 'Definition' @(39));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
        Check 'Directory.Recording.RegisteredDefinition' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_WAREHOUSE_DIRECTORY' -and $definition.Class -ceq 'Navigation' -and $definition.Caption -ceq 'View Warehouses' -and $definition.Capability -ceq 'ADMIN_MAINT')
        Check 'Directory.Recording.AbsentFrom38' ([string](Probe 'Definition' @(38)) -ceq '')
        Check 'Directory.Recording.PriorDefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
        $cases=@('Opened','Protected','Disabled','ControlOff','UserOff','Reader','SignedOut')
        if($Phase -ceq 'GREEN'){$cases+='LoggingReadScope'}
        foreach($case in $cases){
            SelectTarget $Fixture
            $mode=if($case -cin @('Disabled','ControlOff','UserOff')){$case}else{'Enabled'}
            if(-not [bool](Probe 'SetPolicy' @($mode))){throw 'Directory recording policy fixture unavailable.'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root
            if($case -ceq 'Reader'){SelectTarget $Fixture 'config-reader'}
            if($case -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            $denied=$case -cin @('Reader','SignedOut');$denials=[int](Probe 'DenialCount')
            if($case -ceq 'LoggingReadScope'){[void](Invoke 'ArmLateReadScope' @($Other.Config,$Other.Root))}
            [void](Invoke 'ProtectDirectory' @(($case -ceq 'Protected')))
            try{$values=([string](Invoke 'Invoke')) -split '\|'}finally{[void](Invoke 'ProtectDirectory' @($false))}
            if($values.Count -ne 6){throw 'Directory callback facts unavailable.'}
            $opened=-not $denied -and $case -cne 'Protected';$prefix='Directory.Recording.'+$case+'.'
            Check ($prefix+'OwnerOnceOrDenied') ($values[0] -ceq $(if($denied){'0'}else{'1'}))
            Check ($prefix+'ActualPresentation') ($values[1] -ceq [string]$opened -and $values[2] -ceq [string]$opened)
            Check ($prefix+'ExistingFeedback') ($values[3] -ceq $(if($opened){'1'}else{'0'}) -and $values[4] -ceq $(if($case -ceq 'Protected'){'1'}else{'0'}) -and $values[5] -ceq '0')
            Check ($prefix+'PermissionGate') (([int](Probe 'DenialCount')-$denials) -eq $(if($denied){1}else{0}))
            Check ($prefix+'NoTrackingWarning') ([int](Invoke 'TrackingNotices') -eq 0)
            if($case -ceq 'LoggingReadScope'){
                Check ($prefix+'ActualOptionalReadScopeChanged') ([bool](Invoke 'LateReadScopeObserved'))
                Check ($prefix+'OriginalDirectoryReadScopeRetained') ([bool](Invoke 'Preserved' @('Managed')))
                [void](Invoke 'CloseLateReadScope' @($Fixture.Root))
            }
            $preserved=$true;foreach($part in @('Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
            Check ($prefix+'LocalAndOtherWorkbookPreserved') ($preserved -and $otherBook.Worksheets.Count -eq $otherCount -and [string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other directory workbook canary')
            Check ($prefix+'SavedAuthorityAndHistoryPreserved') ((Retained $authority) -and (Retained $otherAuthority))
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_DIRECTORY_OPEN')
            if($denied -or $mode -cne 'Enabled'){Check ($prefix+'NoFabricatedAction') ($records.Count -eq 0);continue}
            $expected=if($opened){'OPENED'}else{'FAILED'}
            $requests=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($records|Where-Object OutcomeCode -CEQ $expected)
            $pair=$records.Count -eq 2 -and $requests.Count -eq 1 -and $outcomes.Count -eq 1
            Check ($prefix+'ExactCorrelatedPair') ($pair -and $requests[0].ActivityId -ceq $outcomes[0].ActivityId -and $requests[0].RecordId -cne $outcomes[0].RecordId)
            $owned=$pair;$safe=$pair;$conclusion=$pair
            foreach($record in $records){
                $owned=$owned -and $record.OwnerId -ceq 'ADMIN_WAREHOUSE_DIRECTORY' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.StationId -ceq 'S1' -and $record.CatalogVersion -eq $catalogVersion -and $record.DataEffect -ceq 'Unknown'
                $json=$record|ConvertTo-Json -Depth 20 -Compress
                $safe=$safe -and @($record.SourceEventRefs).Count -eq 0
                foreach($value in @($Fixture.Root,$Other.Root,$Fixture.Secret,'Keep this local note')){$safe=$safe -and $json.IndexOf($value,[StringComparison]::OrdinalIgnoreCase) -lt 0}
                $conclusion=$conclusion -and [bool](Probe 'Concludes' @($json)) -eq ($record.OutcomeCode -ceq 'OPENED')
            }
            Check ($prefix+'CapturedOwnerContextAndUnknownEffect') $owned
            Check ($prefix+'NoPayloadOrSourceReferences') $safe
            Check ($prefix+'HonestCommandConclusion') $conclusion
        }
        if($Phase -ceq 'GREEN'){Test-AdminDirectoryBoundaries $Fixture $Other $hostBook $otherBook}
        if($AdminDirectoryGuidePaths){Test-AdminDirectoryGuide $Fixture $hostBook}
    }finally{[void](Invoke 'CloseLateReadScope' @($Fixture.Root));SelectTarget $Fixture;$hostBook.Close($false);$otherBook.Close($false)}
}
