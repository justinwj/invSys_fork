[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-admin-console-03',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$Recording,[switch]$GuidePaths,[switch]$RootWriteDiagnostic,[switch]$RootRecording,[switch]$RootGuidePaths)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated Admin directory columns gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/admin-directory-columns-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter '*.ps1'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$testPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Admin directory columns controller: '+$root)
try{
    $extra=@();if($Recording -or $GuidePaths){$extra+='-AdminDirectoryRecording'};if($GuidePaths){$extra+='-AdminDirectoryGuidePaths'}
    if($RootWriteDiagnostic -or $RootRecording -or $RootGuidePaths){$extra+='-AdminWarehouseRootWriteDiagnostic'}
    if($RootRecording -or $RootGuidePaths){$extra+='-AdminWarehouseRootRecording'}
    if($RootGuidePaths){$extra+='-AdminWarehouseRootGuidePaths'}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -AdminDirectoryColumnsOnly -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object {(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;ExcelClosed=$true;FullAcceptance=$false;NativeInputAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $testsSame)){throw 'Admin directory columns preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 14
exit $code
