[CmdletBinding()]
param([ValidateSet('RED','GREEN')][string]$Phase='RED')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated native dismissal testing.'}
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$root=Join-Path $repo ('reports/runtime/live-role-dismissal-owner/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$source=Join-Path $repo 'tools/validate_phase6_live_role_workflows.ps1'
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Helper source does not parse; not behavioral RED.'}
$functions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Start-ExcelEnterDismissal'},$true))
if($functions.Count -ne 1){throw 'Actual dismissal helper unavailable; not behavioral RED.'}
. ([scriptblock]::Create($functions[0].Extent.Text))
$fixture=Join-Path $root 'sentinel.ps1'
@'
param([string]$Root)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type 'using System;using System.Runtime.InteropServices;public static class SentinelWindow{[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int n);}'
$script:count=0
$form=New-Object Windows.Forms.Form
$form.Text='Microsoft Excel';$form.Width=530;$form.Height=170
$label=New-Object Windows.Forms.Label
$label.Text='invSys test window only. No Excel or workbook is open.'
$label.AutoSize=$true;$label.Left=15;$label.Top=15;$form.Controls.Add($label)
$button=New-Object Windows.Forms.Button
$button.Text='OK';$button.Width=180;$button.Left=15;$button.Top=55
$button.Add_Click({$script:count++;[IO.File]::WriteAllText((Join-Path $Root 'count.txt'),[string]$script:count)})
$form.Controls.Add($button);$form.AcceptButton=$button
$form.Add_Shown({[void][SentinelWindow]::ShowWindow($form.Handle,5);$button.Focus();[IO.File]::WriteAllText((Join-Path $Root 'handle.txt'),[string]$form.Handle.ToInt64());[IO.File]::WriteAllText((Join-Path $Root 'count.txt'),'0');[IO.File]::WriteAllText((Join-Path $Root 'ready.txt'),[string]$PID)})
$timer=New-Object Windows.Forms.Timer;$timer.Interval=100
$deadline=[DateTimeOffset]::UtcNow.AddSeconds(55)
$timer.Add_Tick({if((Test-Path (Join-Path $Root 'stop.txt')) -or [DateTimeOffset]::UtcNow -gt $deadline){$form.Close()}})
$timer.Start()
try{[Windows.Forms.Application]::Run($form)}finally{$timer.Stop();$timer.Dispose();$form.Dispose()}
'@ | Set-Content -LiteralPath $fixture -Encoding UTF8
$process=$null;$job=$null;$rows=@();$started=[DateTimeOffset]::UtcNow
function Count { [int][IO.File]::ReadAllText((Join-Path $root 'count.txt')) }
try{
 $process=Start-Process powershell.exe -ArgumentList @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',('"'+$fixture+'"'),'-Root',('"'+$root+'"')) -WindowStyle Hidden -PassThru
 $deadline=[DateTimeOffset]::UtcNow.AddSeconds(15)
 while(-not(Test-Path (Join-Path $root 'ready.txt'))){if($process.HasExited -or [DateTimeOffset]::UtcNow -gt $deadline){throw 'Native sentinel unavailable; not behavioral RED.'};Start-Sleep -Milliseconds 100}
 if([int][IO.File]::ReadAllText((Join-Path $root 'ready.txt')) -ne $process.Id){throw 'Sentinel identity mismatch.'}
 $process.Refresh()
 [pscustomobject]@{ProcessId=$process.Id;MainWindowHandle=$process.MainWindowHandle.ToInt64();ShownFormHandle=[long][IO.File]::ReadAllText((Join-Path $root 'handle.txt'));TitleMatches=($process.MainWindowTitle -ceq 'Microsoft Excel')}|ConvertTo-Json|Tee-Object (Join-Path $root 'sentinel-window.json')
 $before=Count
 $job=Start-ExcelEnterDismissal -Seconds 3 -ExcelProcessId $process.Id
 if(-not(Wait-Job $job -Timeout 20)){throw 'Positive control timed out; not behavioral RED.'}
 Receive-Job $job -ErrorAction Stop|Out-Null;Remove-Job $job;$job=$null
 $after=Count
 if($after -le $before){throw 'Owned input positive control failed; not behavioral RED.'}
 $rows+=[pscustomobject]@{Check='OwnedProcessReceivesInput';Passed=$true}
 $missing=2147483000
 if(Get-Process -Id $missing -ErrorAction SilentlyContinue){throw 'Expected absent process exists.'}
 $before=Count
 $job=Start-ExcelEnterDismissal -Seconds 3 -ExcelProcessId $missing
 if(-not(Wait-Job $job -Timeout 20)){throw 'Missing-owner check timed out; not behavioral RED.'}
 Receive-Job $job -ErrorAction Stop|Out-Null;Remove-Job $job;$job=$null
 $rows+=[pscustomobject]@{Check='MissingOwnerDoesNotTargetDifferentProcess';Passed=((Count) -eq $before)}
}finally{
 if($null -ne $job){Stop-Job $job -ErrorAction SilentlyContinue;Remove-Job $job -Force -ErrorAction SilentlyContinue}
 [IO.File]::WriteAllText((Join-Path $root 'stop.txt'),'stop')
 if($null -ne $process -and -not $process.WaitForExit(5000)){throw 'Sentinel did not close normally.'}
 $rows|ConvertTo-Json|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
 [pscustomobject]@{StartUTC=$started.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');SourceSha256=(Get-FileHash $source).Hash;ActualHelperInvoked=$true;ExcelStarted=$false;SentinelClosed=($null -ne $process -and $process.HasExited);ForcedTermination=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
}
Write-Output ('Dismissal owner evidence: '+$root)
$rows|Format-Table -AutoSize
if(@($rows|Where-Object{-not $_.Passed}).Count){exit 1}
