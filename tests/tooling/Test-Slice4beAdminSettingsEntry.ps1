[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-run-scale-success-02',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$Navigation,[switch]$NavigationFailures,[switch]$EntryFailures)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($NavigationFailures -and -not $Navigation){throw 'Navigation failures require the baseline navigation gate.'}
if($EntryFailures -and $Navigation){throw 'Entry failures require the public-launcher baseline.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before Settings entry validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/admin-entry-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Admin entry controller: '+$root)
try{
    $extra=@();if($Navigation){$extra+='-CheckAdminSettingsNavigation'}
    if($NavigationFailures){$extra+='-CheckAdminSettingsNavigationFailures'}
    if($EntryFailures){$extra+='-CheckAdminSettingsEntryFailures'}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -AdminSettingsCloseOnly -CheckAdminSettingsEntry -CheckTrackingSettings -CheckTrackingPolicy -CheckDetailProfile -CheckActionPathPreference -CaptureEvidence -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object{(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Navigation=[bool]$Navigation;NavigationFailures=[bool]$NavigationFailures;EntryFailures=[bool]$EntryFailures;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true;FullAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same)){throw 'Settings entry preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 8
exit $code
