# Public modal launcher proof. Observers never fabricate an opening result or event.
function Install-AdminSettingsEntryProbe($TestModule,$FormCode) {
    # Observe activation after the real Activate body, before the established test driver dismisses it.
    $anchor='If TestD5Commands.DriveDefaultClose(Me) Then Exit Sub'
    $hits=@(for($line=1;$line -le $FormCode.CountOfLines;$line++){if($FormCode.Lines($line,1).Trim() -ceq $anchor){$line}})
    if($hits.Count -ne 1){throw 'Modal Close probe boundary changed; not behavioral RED.'}
    $FormCode.DeleteLines($hits[0],1)
    $start=$FormCode.ProcStartLine('UserForm_Activate',0);$end=$start+$FormCode.ProcCountLines('UserForm_Activate',0)
    $ends=@(for($line=$start;$line -lt $end;$line++){if($FormCode.Lines($line,1).Trim() -ceq 'End Sub'){$line}})
    if($ends.Count -ne 1){throw 'Activate observation boundary changed; not behavioral RED.'}
    $FormCode.InsertLines($ends[0],"    TestD5Commands.EntryActivated Me.Visible`r`n    If TestD5Commands.DriveDefaultClose(Me) Then Exit Sub")
    $TestModule.CodeModule.AddFromString(@'
Private mEntryVisibleCount As Long
Public Sub EntryActivated(ByVal visible As Boolean)
    If visible Then mEntryVisibleCount = mEntryVisibleCount + 1
End Sub
Public Function EntryVisibleCount() As Long
    EntryVisibleCount = mEntryVisibleCount
End Function
'@)
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1);$probe.Name='TestAdminEntry'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function EnablePolicy(Optional ByVal mode As String = "Enabled") As Boolean
    Dim context As String, version As Long, request As String, report As String
    Dim model As Object, row As Variant, user As Object, users As Collection
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    If mode <> "Defaults" And mode <> "Disabled" Then
        For Each row In model("Controls")
            row("Collect") = Not (mode = "ControlsOff" And Left$(row("ControlId"), 20) = "ADMIN_SETTINGS_PAGE_")
            If mode = "OpenOff" And row("ControlId") = "ADMIN_SETTINGS_OPEN" Then row("Collect") = False
        Next row
    End If
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    request = modTrainingJson.EncodeObject(model)
    EnablePolicy = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function HistoricalPolicyIds() As String
    HistoricalPolicyIds = Join(modActivityCatalog.ControlIds(32), vbLf)
End Function
Public Function ExistingPolicyReadable() As Boolean
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String
    ExistingPolicyReadable = modActivityPolicy.ReadPolicy(modNasConnection.GetCurrentTarget(), "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice)
    ExistingPolicyReadable = ExistingPolicyReadable And collect
End Function
Public Function Definition(ByVal id As String, ByVal version As Long) As String
    Dim value As Object
    Set value = modActivityCatalog.Control(id, version)
    If Not value Is Nothing Then Definition = modTrainingJson.EncodeObject(value)
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
}

function Test-AdminSettingsEntry($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAdminEntry.'+$Method) $Values}
    SelectTarget $Fixture
    if(-not [bool](Probe 'EnablePolicy')){throw 'Explicit recording policy fixture unavailable; not behavioral RED.'}
    $id='ADMIN_SETTINGS_OPEN';$before=@(Files)
    $configHash=(Get-FileHash $Fixture.Config).Hash
    $otherPins=@{};foreach($file in Get-ChildItem $Other.Root -Recurse -File){$otherPins[$file.FullName]=(Get-FileHash $file.FullName).Hash}
    Test-AdminSettingsDefaultClose $Fixture
    $visible=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.EntryVisibleCount')
    if($visible -ne 2){throw 'Two real visible modal activations required; not observation RED.'}
    Check 'AdminEntry.TwoVisiblePublicLauncherActivations' ($visible -eq 2)
    $wire=[string](Probe 'Definition' @($id,33));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check 'AdminEntry.Catalog.OpenNavigation' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $definition.Class -ceq 'Navigation' -and $definition.Role -ceq 'Admin' -and $definition.Capability -ceq 'ADMIN_MAINT')
    Check 'AdminEntry.Catalog.AbsentFromPriorVersion' ([string](Probe 'Definition' @($id,32)) -ceq '')
    $records=@(Files|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ $id)
    $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($records|Where-Object OutcomeCode -CEQ 'OPENED')
    Check 'AdminEntry.OnePairPerVisibleInvocation' ($records.Count -eq 4 -and $attempts.Count -eq 2 -and $outcomes.Count -eq 2)
    $correlated=$attempts.Count -eq 2;$context=$records.Count -eq 4;$facts=$context;$safe=$context;$terminal=$context
    foreach($attempt in $attempts){$last=@($outcomes|Where-Object ActivityId -CEQ $attempt.ActivityId);$correlated=$correlated -and $last.Count -eq 1;if($last.Count -eq 1){$correlated=$correlated -and $attempt.RecordId -cne $last[0].RecordId}}
    foreach($record in $records){
        $context=$context -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $record.CatalogVersion -eq 33
        $positive=$record.OutcomeCode -ceq 'OPENED'
        $facts=$facts -and $record.Severity -ceq 'Info' -and $record.DataEffect -ceq $(if($positive){'Unchanged'}else{'Unknown'}) -and $record.EventCode -ceq ($id+'_'+$record.OutcomeCode)
        if($positive){$facts=$facts -and $record.UserMessage -ceq 'Settings form opened. No settings save or Domain application is asserted.'}
        $json=ConvertTo-Json $record -Depth 20 -Compress
        $safe=$safe -and @($record.SourceEventRefs).Count -eq 0
        foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'PinHash')){$safe=$safe -and -not $json.Contains($value)}
        $terminal=$terminal -and [bool](Probe 'Concludes' @($json)) -eq $positive
    }
    Check 'AdminEntry.ExactDistinctCorrelatedPairs' ($correlated -and @($attempts.ActivityId|Select-Object -Unique).Count -eq 2)
    Check 'AdminEntry.CapturedContext' $context
    Check 'AdminEntry.PresentationOnlyOwnerFacts' $facts
    Check 'AdminEntry.NoEnteredValuesOrSourceReferences' $safe
    Check 'AdminEntry.ExactLocalConclusion' $terminal
    $before=@(Files);[void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings');[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    Check 'AdminEntry.InternalConstructorAndCleanupExcluded' (@(Files).Count -eq $before.Count)
    Check 'AdminEntry.SavedConfigPreserved' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
    $same=$true;foreach($file in $otherPins.Keys){$same=$same -and (Get-FileHash $file).Hash -ceq $otherPins[$file]}
    Check 'AdminEntry.OtherWarehousePreserved' $same
}
