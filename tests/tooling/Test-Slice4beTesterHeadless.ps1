[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-tester-owner-08')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the packaged headless regression.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/tester-headless' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$fixtureRoot=[IO.Path]::GetFullPath((Join-Path $env:TEMP ('invSys_tester_headless_'+[guid]::NewGuid().ToString('N'))))
if(-not $fixtureRoot.StartsWith([IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Fixture root is outside Temp.'}
$pins=@(Get-ChildItem $DeployRoot -Filter '*.xlam' -File|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$start=[DateTimeOffset]::UtcNow;$excel=$null;$rows=@();$passed=$false
Write-Output ('Headless controller: '+$root)
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    $books=@{}
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){$books[$name]=$excel.Workbooks.Open((Join-Path (Resolve-Path $DeployRoot).Path $name),0,$true)}
    $project=$books['invSys.Admin.xlam'].VBProject
    # Existing five-case regression, imported unsaved. Never emit its raw details.
    $helpers=$project.VBComponents.Add(1);$helpers.Name='TestPhase2Helpers'
    $helpers.CodeModule.AddFromString(@'
Option Explicit
Public Function GetRowValue(ByVal table As ListObject, ByVal row As Long, ByVal header As String) As Variant
    GetRowValue = table.ListRows(row).Range.Cells(1, table.ListColumns(header).Index).Value2
End Function
'@)
    $test=$project.VBComponents.Import((Join-Path $PWD 'tests/integration/test_TesterSetup.bas'))
    $module=$test.CodeModule
    $first=$module.ProcStartLine('BuildTesterSetupTempRoot',0);$count=$module.ProcCountLines('BuildTesterSetupTempRoot',0)
    $module.DeleteLines($first,$count)
    $module.InsertLines($first,('Private Function BuildTesterSetupTempRoot(ByVal suffix As String) As String'+"`r`n"+'BuildTesterSetupTempRoot = "'+$fixtureRoot.Replace('"','""')+'\" & suffix'+"`r`nEnd Function"))
    $first=$module.ProcStartLine('CleanupTesterSetupRoot',0);$count=$module.ProcCountLines('CleanupTesterSetupRoot',0)
    $module.DeleteLines($first,$count)
    # Leave disposable evidence in the verified root; no recursive deletion during this gate.
    $module.InsertLines($first,"Private Sub CleanupTesterSetupRoot(ByVal rootPath As String)`r`nEnd Sub")
    $module.AddFromString(@'
Public Function FixtureNormalizationForTest(ByVal root As String) As Boolean
    Dim spec As modTesterSetup.TesterSetupSpec, completed As Boolean
    spec.UserId = " fixture-user ": spec.WarehouseId = " FIXTURE "
    spec.PathLocal = " " & Replace$(root, "\", "/") & " "
    spec.PathSharePointRoot = root & "\publication"
    completed = modTesterSetup.SetupTesterStation(spec)
    FixtureNormalizationForTest = Not completed And spec.UserId = "fixture-user" And _
        spec.WarehouseId = "FIXTURE" And spec.StationId = "R1" And spec.PathLocal = root
End Function
'@)
    $excel.VBE.ActiveVBProject=$project;$module.CodePane.Show()
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($compile.Enabled){$compile.Execute()}
    if($excel.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Headless test adapter did not compile.'}
    $passed=([int]$excel.Run("'invSys.Admin.xlam'!test_TesterSetup.TestTesterSetup_EndToEnd") -eq 1)
    $packed=[string]$excel.Run("'invSys.Admin.xlam'!test_TesterSetup.GetTesterSetupEvidenceRows")
    foreach($line in ($packed -split "`r?`n")){if($line){$fields=$line -split "`t",3;$rows+=[pscustomobject]@{Check=$fields[0];Result=$fields[1]};Write-Output ($fields[0]+': '+$fields[1])}}
    $packed=''
    $normalized=[bool]$excel.Run("'invSys.Admin.xlam'!test_TesterSetup.FixtureNormalizationForTest",$fixtureRoot)
    $rows+=[pscustomobject]@{Check='Headless.FailedSetupReturnsNormalizedSpec';Result=$(if($normalized){'PASS'}else{'FAIL'})}
    Write-Output ('Headless.FailedSetupReturnsNormalizedSpec: '+$rows[-1].Result)
    $passed=$passed -and $normalized
}finally{
    if($null -ne $excel){foreach($book in @($excel.Workbooks)){$book.Close($false)};$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');Passed=$passed;Checks=$rows;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root 'result.json')
    if(-not($restored -and $same)){throw 'Headless regression preservation failed.'}
}
if(-not $passed -or $rows.Count -ne 6 -or @($rows|Where-Object Result -ne PASS).Count){throw 'Headless regression failed; inspect sanitized case results.'}
