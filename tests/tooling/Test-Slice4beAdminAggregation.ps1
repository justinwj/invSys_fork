[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-tester-recording-02',[ValidateSet('RED','GREEN')][string]$Phase='RED',[string]$DesktopMonitorReport='',[switch]$PathComparison,[switch]$BoundaryMatrix,[switch]$WriteBoundariesOnly)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated Aggregation owner gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/admin-aggregation-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter '*.ps1'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$testPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1;$worker=$null
Write-Output ('Aggregation owner controller: '+$root)
try {
    $workerArgs=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',
        (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1'),'-DeployRoot',$DeployRoot,
        '-Phase',$Phase,'-AdminAggregationOnly','-CompileEvaluationProbesForTest',
        '-WaitForExcelReadyForTest','-ExcelReadyReadLimitForTest','40')
    if($DesktopMonitorReport){$workerArgs+=@('-DesktopMonitorReport',$DesktopMonitorReport)}
    if($PathComparison){$workerArgs+='-AdminAggregationPathComparison'}
    if($BoundaryMatrix){$workerArgs+='-AdminAggregationBoundaryMatrix'}
    if($WriteBoundariesOnly){$workerArgs+='-AdminAggregationBoundaryMatrix';$workerArgs+='-AdminAggregationWriteBoundariesOnly'}
    $quotedArgs=($workerArgs|ForEach-Object {'"'+$_+'"'}) -join ' '
    $launch=[Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME 'powershell.exe'))
    $launch.Arguments=$quotedArgs;$launch.WorkingDirectory=$PWD.Path
    $launch.UseShellExecute=$false;$launch.CreateNoWindow=$true
    $launch.RedirectStandardOutput=$true;$launch.RedirectStandardError=$true
    $outputFile=[IO.FileStream]::new((Join-Path $PWD.Path (Join-Path $root 'worker.log')),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,$true)
    $errorFile=[IO.FileStream]::new((Join-Path $PWD.Path (Join-Path $root 'worker-error.log')),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,$true)
    try{
        $worker=[Diagnostics.Process]::Start($launch)
        $outputCopy=$worker.StandardOutput.BaseStream.CopyToAsync($outputFile,256)
        $errorCopy=$worker.StandardError.BaseStream.CopyToAsync($errorFile,256)
        $worker.WaitForExit();$outputCopy.Wait();$errorCopy.Wait()
        if(-not $worker.HasExited){throw 'Aggregation worker termination was not confirmed.'}
        $code=$worker.ExitCode
    }finally{$outputFile.Dispose();$errorFile.Dispose()}
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $worker
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;ExcelClosed=$true;FullAcceptance=$false;NativeInputAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $testsSame)){throw 'Aggregation diagnostic preservation failed.'}
}
# Delayed native audit: process exit alone does not prove clean native shutdown.
Start-Sleep -Seconds 20
$audit=[DateTimeOffset]::UtcNow
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors|Where-Object {$_.Message -match '(?i)excel\.exe'})
if(@($queryErrors|Where-Object FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*').Count){throw 'Aggregation native audit unavailable.'}
$closed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0
[pscustomobject]@{StartUTC=$start.ToString('o');AuditUTC=$audit.ToString('o');ExcelEventCount=$events.Count;ExcelClosed=$closed}|ConvertTo-Json|Set-Content (Join-Path $root 'native-audit.json')
if($events.Count -or -not $closed){throw 'Aggregation native closure did not pass.'}
Get-Content (Join-Path $root 'worker.log') -Tail 14
exit $code
