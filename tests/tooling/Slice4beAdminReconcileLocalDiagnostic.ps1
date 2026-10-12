# Observes current local eligibility; no new recording policy or native-input claim.
# Only generated fixtures and disposable, unsaved test modules are touched.
function Install-ReconcileLocalDiagnostic {
    $project=$packages['invSys.Core.xlam'].VBProject
    $project.VBComponents.Item('modActivityStore').CodeModule.AddFromString(@'
Public Function LocalDiagnosticRootForTest(ByVal target As WarehouseTarget, ByVal create As Boolean) As Boolean
    LocalDiagnosticRootForTest = (StoreRoot(target, create) <> "")
End Function
'@)
    $module=$project.VBComponents.Add(1);$module.Name='TestReconcileLocal'
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Function PreparePolicyAndRoots() As Boolean
    Dim context As String, version As Long, wire As String, report As String, model As Object
    Dim target As WarehouseTarget
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(True)
    model("ViewerActionPathCaptureEnabled") = True
    wire = modTrainingJson.EncodeObject(model)
    If Not modTrackingPolicySettings.SavePolicy(context, version, wire, report) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    PreparePolicyAndRoots = modActivityStore.LocalDiagnosticRootForTest(target, True) And _
        (modRecordingJournal.JournalRoot(target, True) <> "")
End Function
Public Function ResolveExplicitLocal(ByVal root As String, ByVal secret As String) As Boolean
    Dim target As WarehouseTarget, status As NasStatusCode, signed As String
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.SetCoreDataRootOverride root
    If Not modNasConnection.ResolveWarehouseTarget(target, status) Then Exit Function
    If status <> NAS_OK Or target.SourceType <> WH_SOURCE_LOCAL Then Exit Function
    If StrComp(target.HubRoot, root, vbTextCompare) <> 0 Or StrComp(target.RuntimeRoot, root, vbTextCompare) <> 0 Then Exit Function
    signed = modAuth.SignInCurrentTargetForAutomation("config-admin", secret, "")
    secret = ""
    ResolveExplicitLocal = (Left$(signed, 3) = "OK|" And modAuth.IsSignedIn())
End Function
Public Function ObserveBoundary() As String
    Dim target As WarehouseTarget, version As Long, collect As Boolean, visible As Boolean
    Dim notice As String, capture As Boolean, eligible As Boolean, policy As Boolean
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    policy = modActivityPolicy.ReadPolicy(target, "ADMIN_RECONCILE_INSPECT", version, collect, visible, notice, Nothing, capture, eligible)
    ObserveBoundary = CStr(target.SourceType) & "|" & CStr(modAuth.IsSignedIn()) & "|" & _
        CStr(modNasConnection.IsCurrentTargetAllowed(False)) & "|" & CStr(modNasConnection.IsCurrentTargetAllowed(True)) & "|" & _
        CStr(modActivity.CaptureContext() <> "") & "|" & CStr(modActivityStore.LocalDiagnosticRootForTest(target, False)) & "|" & _
        CStr(modRecordingJournal.JournalRoot(target, False) <> "") & "|" & CStr(policy And collect And capture And eligible)
End Function
'@)
}

function Test-ReconcileLocalDiagnostic($Fixture,$Events,$State) {
    function TrainingPins {
        $result=@{};$root=Join-Path $Fixture.Root 'Training'
        if(Test-Path -LiteralPath $root){Get-ChildItem -LiteralPath $root -Recurse -File|ForEach-Object {$result[$_.FullName]=Hash $_.FullName}}
        return $result
    }
    function Preserved($Before,$After){foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}};return $true}
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestReconcileLocal.PreparePolicyAndRoots')){throw 'Local diagnostic policy/root fixture unavailable; not product RED.'}
    $policyHash=Hash $Fixture.Config
    $observations=@()
    try{
        foreach($case in @('NASFixture','ExplicitLocal','FallbackFixture')){
            if($case -ceq 'ExplicitLocal'){
                if(-not [bool](Run 'invSys.Core.xlam' 'TestReconcileLocal.ResolveExplicitLocal' @($Fixture.Root,$Fixture.Secret))){throw 'Actual explicit local resolution failed; not product RED.'}
            }elseif($case -ceq 'FallbackFixture'){
                if(-not [bool](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetSourceTypeForTest' @(3))){throw 'Fallback fixture unavailable.'}
            }
            $values=([string](Run 'invSys.Core.xlam' 'TestReconcileLocal.ObserveBoundary')) -split '\|'
            if($values.Count -ne 8){throw 'Local diagnostic envelope unavailable.'}
            $expectedType=if($case -ceq 'NASFixture'){0}elseif($case -ceq 'ExplicitLocal'){2}else{3}
            $nas=$case -ceq 'NASFixture'
            $prefix='Reconcile.LocalDiagnostic.'+$case+'.'
            Check ($prefix+'ExactTargetKindAndLiveBusinessAccess') ($values[0] -ceq [string]$expectedType -and $values[1] -ceq 'True' -and $values[2] -ceq 'True')
            Check ($prefix+'EnabledPolicyStillReadable') ($values[7] -ceq 'True')
            Check ($prefix+'ObservedNASContextAndStoreBoundary') (($values[3..6] -join '|') -ceq ((@([string]$nas)*4) -join '|'))
            $before=TrainingPins;$inventoryHash=Hash $State.InventoryPath;$rows=$Events.ListRows.Count
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareAuthorityCase' @('LocalSignedIn'))){throw 'Reconcile callback fixture unavailable.'}
            try{
                $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeCallback')
                $notices=[int](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.TrackingNoticeCount')
                $after=TrainingPins
                Check ($prefix+'ActualCallbackQueuesExactlyOnce') ($called -and $Events.ListRows.Count -eq $rows+1)
                $row=$Events.ListRows.Item($Events.ListRows.Count)
                $items=ConvertFrom-Json -InputObject ([string]$row.Range.Cells.Item(1,$Events.ListColumns.Item('PayloadJson').Index).Value2)
                Check ($prefix+'ExactEntityCorrection') ($items.Count -eq 1 -and $items[0].System_Key -ceq $State.KeyA)
                Check ($prefix+'InventoryAndCustomValuesPreserved') ((Hash $State.InventoryPath) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.CanaryPreserved'))
                Check ($prefix+'PriorTrainingEvidencePreserved') (Preserved $before $after)
                Check ($prefix+'ObservedRecordingAvailability') $(if($nas){$after.Count -gt $before.Count}else{$after.Count -eq $before.Count})
                Check ($prefix+'PolicyUnchanged') ((Hash $Fixture.Config) -ceq $policyHash)
                $observations+=[pscustomobject]@{Case=$case;SourceType=[int]$values[0];SignedIn=[bool]::Parse($values[1]);BusinessAllowed=[bool]::Parse($values[2]);NASAllowed=[bool]::Parse($values[3]);ContextAvailable=[bool]::Parse($values[4]);ActivityStoreAvailable=[bool]::Parse($values[5]);JournalStoreAvailable=[bool]::Parse($values[6]);PolicyEnabled=[bool]::Parse($values[7]);Queued=$Events.ListRows.Count-$rows;NewTrainingFiles=$after.Count-$before.Count;TrackingNotices=$notices;PolicyAccepted=$false}
            }finally{[void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RestoreAuthorityFixture')}
        }
        $observations|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'local-recording-diagnostic.json')
    }finally{SelectTarget $Fixture}
}
