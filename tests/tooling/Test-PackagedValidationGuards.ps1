[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_phase6_packaged_xlams.ps1'),[ref]$tokens,[ref]$errors)
if($errors){throw 'Validator does not parse; not a behavioral guard failure.'}
# Load function definitions only: never execute the validator or create Excel.
foreach($definition in $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true)){
    . ([scriptblock]::Create($definition.Extent.Text))
}
$root=Join-Path $repo ('reports/runtime/packaged-validator-guards/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed}); Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
$mockExcel=[pscustomobject]@{}
$mockExcel | Add-Member ScriptMethod Run {param($Macro,$Argument) $script:dispatches++; return 'CALLED'}
$script:DesktopMonitorReport=Join-Path $root 'desktop.json'
foreach($case in @('Healthy','Error5','Locked','MouseLocked','Stale','Malformed')){
    $state=[ordered]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Error5Samples=0;LockSamples=0;MouseLockSamples=0}
    switch($case){
        Error5 {$state.Error5Samples=1}
        Locked {$state.LockSamples=1}
        MouseLocked {$state.MouseLockSamples=1}
        Stale {$state.UTC=[DateTimeOffset]::UtcNow.AddMinutes(-2).ToString('o')}
    }
    if($case -eq 'Malformed'){'{'|Set-Content $DesktopMonitorReport}else{$state|ConvertTo-Json|Set-Content $DesktopMonitorReport}
    foreach($wrapper in @('Run-WorkbookMacro','Run-WorkbookMacro1')){
        $script:dispatches=0; $script:PackagedDesktopStopped=$false; $refused=$false
        try {
            if($wrapper -eq 'Run-WorkbookMacro'){Run-WorkbookMacro $mockExcel 'fixture.xlam' 'Fixture.Action'}
            else {[void](Run-WorkbookMacro1 $mockExcel 'fixture.xlam' 'Fixture.Action' 'fixture')}
        } catch {$refused=$true}
        $expected=if($case -eq 'Healthy'){(-not $refused) -and $script:dispatches -eq 1}else{$refused -and $script:dispatches -eq 0}
        Check ($case+'.'+$wrapper) $expected
    }
}

# Fake HWND/process and plain objects exercise the real cleanup function.
# No native API, real process query, process termination or COM host is used.
Add-Type @'
using System;
public static class PackagedValidationProcess {
    public static uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId) {processId=4242;return 1;}
}
'@
$script:stopCalls=0; $script:waitCalls=0; $script:fakeExited=$false
$script:fakeProcess=[pscustomobject]@{HasExited=$false}
$script:fakeProcess | Add-Member ScriptMethod WaitForExit {
    param($Milliseconds)
    $script:waitCalls++
    if($script:waitCalls -ge 2){$script:fakeExited=$true;$this.HasExited=$true;return $true}
    return $false
}
function Get-Process {param($Id,$ErrorAction) if(-not $script:fakeExited){return $script:fakeProcess}}
function Stop-Process {param($Id) $script:stopCalls++;$script:fakeExited=$true;$script:fakeProcess.HasExited=$true}
function Release-ComObject {param($Obj)}
$cleanupExcel=[pscustomobject]@{Hwnd=1;Workbooks=@();EnableEvents=$true;DisplayAlerts=$true}
$cleanupExcel | Add-Member ScriptMethod Quit {}
$targetRoot=$root; $deployPath=$root; $closureRoot=$root
$resultRows=[Collections.Generic.List[object]]::new()
$script:PackagedDesktopStopped=$false
Close-PackagedValidationSession -Excel $cleanupExcel -Phase 'Mock'
$facts=Get-Content -Raw (Join-Path $root 'Mock.json')|ConvertFrom-Json
Check 'SlowExit.NoForcedTermination' ($script:stopCalls -eq 0 -and -not $facts.TerminationRequested)
Check 'SlowExit.WaitedForNaturalClosure' ($script:waitCalls -ge 2 -and $facts.UnassistedExit -and $facts.ExitObserved)

# WorkbookBeforeClose is responsible for releasing role forms and cancelling
# scheduled snapshot callbacks. Model delivery through the actual cleanup call.
$script:fakeExited=$false; $script:waitCalls=0; $script:closeHandlerDelivered=$false
$script:fakeProcess.HasExited=$false
$closeBook=[pscustomobject]@{Path=$root;FullName=(Join-Path $root 'operator.xlsx');IsAddin=$false}
$closeBook | Add-Member ScriptMethod Close {
    param($SaveChanges)
    if($cleanupExcel.EnableEvents){$script:closeHandlerDelivered=$true}
    $cleanupExcel.Workbooks=@()
}
$cleanupExcel.Workbooks=@($closeBook)
Close-PackagedValidationSession -Excel $cleanupExcel -Phase 'CloseEvents'
Check 'CloseEvents.OperatorHandlerDelivered' $script:closeHandlerDelivered

# Bounded fake stall: the third wait throws so the pre-fix harness cannot hang.
# The real implementation must terminate the captured process object, record
# assisted recovery and throw before a replacement Excel can be created.
$script:fakeExited=$false; $script:stopCalls=0; $script:waitCalls=0
$script:fakeProcess=[pscustomobject]@{HasExited=$false}
$script:fakeProcess | Add-Member ScriptMethod WaitForExit {
    param($Milliseconds)
    $script:waitCalls++
    if($script:waitCalls -ge 3){throw 'TEST_STALL_BOUND'}
    return $false
}
$script:capturedRecovery=$false
$script:recoveryRecordedBeforeStop=$false
function Stop-Process {
    param($InputObject,$ErrorAction)
    $script:capturedRecovery=[object]::ReferenceEquals($InputObject,$script:fakeProcess)
    $script:recoveryRecordedBeforeStop=Test-Path (Join-Path $root 'Stalled-recovery-requested.json')
    $script:stopCalls++; $script:fakeExited=$true; $script:fakeProcess.HasExited=$true
}
$refused=$false
try {Close-PackagedValidationSession -Excel $cleanupExcel -Phase 'Stalled' -ExitWaitSeconds 0}
catch {$refused=$_.Exception.Message -like '*normal exit*'}
$facts=Get-Content -Raw (Join-Path $root 'Stalled.json')|ConvertFrom-Json
Check 'Stalled.RecoveryUsesCapturedProcess' ($script:stopCalls -eq 1 -and $script:capturedRecovery)
Check 'Stalled.AssistedExitNeverQualifies' ($facts.TerminationRequested -and -not $facts.UnassistedExit -and $facts.ExitObserved)
Check 'Stalled.RefusesRestart' $refused
Check 'Stalled.RecoveryRecordedBeforeTermination' $script:recoveryRecordedBeforeStop
$checks.ToArray()|ConvertTo-Json|Set-Content (Join-Path $root 'results.json')
$failed=($checks|Where-Object {-not $_.Passed}|Measure-Object).Count
Write-Output ('GUARDS: '+($checks.Count-$failed)+' passed, '+$failed+' failed; '+$root)
if($failed){exit 1}
