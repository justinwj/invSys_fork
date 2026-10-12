[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-admin-reconcile-01',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$NativePresentation,[switch]$CheckRecording)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before Reconcile identity validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/admin-reconcile-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
$testPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Admin reconcile controller: '+$root)
try{
    $nativeArgs=@();if($NativePresentation){$nativeArgs+='-AdminReconcileNativePresentation'}
    if($CheckRecording){if(-not $NativePresentation){throw 'Recording gate requires native handlers.'};$nativeArgs+='-AdminReconcileRecording'}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -AdminReconcileIdentityOnly -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @nativeArgs *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object{(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object{(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;ExcelClosed=$true;FullAcceptance=$false;NativeInputAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $testsSame)){throw 'Reconcile identity preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 16
exit $code
