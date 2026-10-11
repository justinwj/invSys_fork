# Produce an ignored diagnostic copy; preserve the original chain and its actions.
param([Parameter(Mandatory=$true)][string]$OutputDirectory,[string]$LateFixtureReceipt='')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$root=if([IO.Path]::IsPathRooted($OutputDirectory)){[IO.Path]::GetFullPath($OutputDirectory)}else{[IO.Path]::GetFullPath((Join-Path $repo $OutputDirectory))}
$allowed=Join-Path $repo 'reports/runtime'
if(-not $root.StartsWith($allowed+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $root)){throw 'A new ignored runtime directory is required.'}
. (Join-Path $PSScriptRoot 'Slice4beShutdownTrace.ps1')
$sourcePath=Join-Path $repo 'tools/validate_release1_full_chain.ps1'
$source=Get-Content -LiteralPath $sourcePath -Raw
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Original chain does not parse.'}
$changes=[Collections.Generic.List[object]]::new()
foreach($definition in $ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)){
    $text=Add-ShutdownTraceToFunction $definition.Name $definition.Extent.Text
    # Literal source anchors only: no Excel access occurs in this generator.
    $created=[regex]::Matches($text,'\$(localExcel|isolationExcel|restoreExcel) = New-Object -ComObject Excel\.Application')
    foreach($match in @($created)){
        $variable='$'+$match.Groups[1].Value
        $owner=$definition.Name+'.'+$match.Groups[1].Value
        $text=$text.Replace($match.Value,($match.Value+"`r`n"+'        Save-ShutdownChainOwner '+$variable+" '"+$owner+"'"))
    }
    if($definition.Name -ceq 'Invoke-RuntimeFivePackageEvidence'){
        foreach($variable in @('$localExcel','$isolationExcel','$restoreExcel')){
            $call=$variable+'.Quit()'
            $marker=$definition.Name+'.'+$variable.Substring(1)
            $text=$text.Replace($call,("Add-ShutdownPhase 'Excel.BeforeQuit' '"+$marker+"'; "+$call+"; Add-ShutdownPhase 'Excel.AfterQuitAttempt' '"+$marker+"'"))
        }
    }
    $changes.Add([pscustomobject]@{Function=$definition.Name;CreationAnchors=$created.Count;Changed=($text -cne $definition.Extent.Text)})
    $source=$source.Replace($definition.Extent.Text,$text)
}
if(($changes|Measure-Object CreationAnchors -Sum).Sum -ne 5){throw 'Expected five chain Excel creation boundaries.'}
$setup=@'
$ReportRoot='__TRACE_ROOT__'
. (Join-Path $repo 'tests/tooling/Slice4beShutdownTrace.ps1')
Add-Type @"
using System;using System.Runtime.InteropServices;
public static class ShutdownChainIdentity {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint processId);
}
"@
function Save-ShutdownChainOwner([object]$Application,[string]$Owner){
    [uint32]$processId=0
    [void][ShutdownChainIdentity]::GetWindowThreadProcessId([IntPtr]$Application.Hwnd,[ref]$processId)
    $process=Get-Process -Id $processId -ErrorAction Stop
    [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Owner=$Owner;ProcessId=$processId;CreatedUTC=$process.StartTime.ToUniversalTime().ToString('o')}|ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $ReportRoot 'owners.jsonl')
}
'@
$setup=$setup.Replace('__TRACE_ROOT__',$root.Replace("'","''"))
$anchor='$deployPath = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path'
if([regex]::Matches($source,[regex]::Escape($anchor)).Count -ne 1){throw 'Chain setup anchor drift.'}
$source=$source.Replace($anchor,($anchor+"`r`n"+$setup))
foreach($stage in @('Invoke-RestartReconciliation -LiveResultText $liveText','Invoke-RuntimeFivePackageEvidence')){
    $pattern='(?m)^    '+[regex]::Escape($stage)+'\r?$'
    $matches=[regex]::Matches($source,$pattern)
    if($matches.Count -ne 1){throw 'Chain stage anchor drift.'}
    $name=($stage -split ' ')[0].Replace('-','.')
    $source=$source.Replace($matches[0].Value,("    Add-ShutdownPhase 'Stage.Before' '"+$name+"'`r`n    "+$stage+"`r`n    Add-ShutdownPhase 'Stage.After' '"+$name+"'"))
}
$fixtureSource=''
if($LateFixtureReceipt){
    $receipt=(Resolve-Path -LiteralPath $LateFixtureReceipt).Path
    $match=[regex]::Match((Get-Content -LiteralPath $receipt -Raw),'(?m)^- Runtime root override:\s*(.+?)\s*$')
    if(-not $match.Success){throw 'Generated fixture reference missing.'}
    $fixtureSource=(Resolve-Path -LiteralPath $match.Groups[1].Value.Trim()).Path
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
    if(-not $fixtureSource.StartsWith($temp+'\invsys-phase6-live-',[StringComparison]::OrdinalIgnoreCase)){throw 'Only a generated live-test fixture is allowed.'}
    $begin=[regex]::Matches($source,'(?m)^    Invoke-AdminEntryGate\r?$')
    $end=[regex]::Matches($source,"(?m)^    Add-ShutdownPhase 'Stage.Before' 'Invoke.RestartReconciliation'\r?$")
    if($begin.Count -ne 1 -or $end.Count -ne 1 -or $end[0].Index -le $begin[0].Index){throw 'Diagnostic cut anchors differ.'}
    $fixture=Join-Path $root 'fixture'
    $configAnchor='$configWorkbook = $localExcel.Workbooks.Open($configPath)'
    if([regex]::Matches($source,[regex]::Escape($configAnchor)).Count -ne 1){throw 'Clone configuration anchor differs.'}
    $source=$source.Replace($configAnchor,($configAnchor+"`r`n"+"        Initialize-ShutdownClonePaths `$configWorkbook '"+$fixtureSource.Replace("'","''")+"' '"+$fixture.Replace("'","''")+"'"))
    # Re-find offsets after inserting the fixture setup above the main body.
    $begin=[regex]::Matches($source,'(?m)^    Invoke-AdminEntryGate\r?$')
    $end=[regex]::Matches($source,"(?m)^    Add-ShutdownPhase 'Stage.Before' 'Invoke.RestartReconciliation'\r?$")
    $source=$source.Substring(0,$begin[0].Index)+"    `$liveText = '- Runtime root override: "+$fixture.Replace("'","''")+"'`r`n"+$source.Substring($end[0].Index)
    $resultAnchor='$resultPath = Join-Path $repo "tests/integration/slice14_results.md"'
    if([regex]::Matches($source,[regex]::Escape($resultAnchor)).Count -ne 1){throw 'Diagnostic report anchor differs.'}
    $source=$source.Replace($resultAnchor,'$resultPath = Join-Path $ReportRoot ''late-results.md''')
    $phaseAssignment=@($ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.AssignmentStatementAst] -and $_.Left.Extent.Text -ceq '$phaseOrder'})
    if($phaseAssignment.Count -ne 1){throw 'Phase list anchor differs.'}
    $source=$source.Replace($phaseAssignment[0].Extent.Text,'$phaseOrder = @(''DiagnosticRestart'',''DiagnosticRuntimeEvidence'',''StaticAudit'')')
    $source=$source.Replace('RELEASE1_FULL_CHAIN_','LATE_CHAIN_DIAGNOSTIC_')
    $source=$source.Replace('# Slice 14 Full-Chain, Restart, and Reconciliation Evidence','# Partial Restart and Runtime Evidence Diagnostic')
    $history='(?ms)^        "## D13 trace",.*?(?=^        "\| Check \| Result \| Detail \|",)'
    if([regex]::Matches($source,$history).Count -ne 1){throw 'Diagnostic provenance anchor differs.'}
    $source=[regex]::Replace($source,$history,('        "- Diagnostic only: earlier full-chain stages were not executed. No product RED or release acceptance is claimed.",'+"`r`n"+'        "",'+"`r`n"))
}
$null=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Instrumented chain does not parse; no dispatch.'}
New-Item -ItemType Directory -Path $root|Out-Null
if($fixtureSource){
    New-Item -ItemType Directory -Path $fixture|Out-Null
    $before=@(Get-ChildItem -LiteralPath $fixtureSource -Recurse -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
    foreach($item in Get-ChildItem -LiteralPath $fixtureSource){Copy-Item -LiteralPath $item.FullName -Destination $fixture -Recurse}
    [pscustomobject]@{ReceiptHash=(Get-FileHash -LiteralPath $receipt).Hash;OriginalFiles=$before;GeneratedFixtureOnly=$true;PartialDiagnostic=$true}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root 'fixture-provenance.json')
}
[IO.File]::WriteAllText((Join-Path $root 'chain.ps1'),$source,[Text.UTF8Encoding]::new($false))
[pscustomobject]@{OriginalHash=(Get-FileHash $sourcePath).Hash;InstrumentedHash=(Get-FileHash (Join-Path $root chain.ps1)).Hash;TraceHelperHash=(Get-FileHash (Join-Path $PSScriptRoot Slice4beShutdownTrace.ps1)).Hash;Functions=$changes;DiagnosticOnly=$true;PartialDiagnostic=[bool]$LateFixtureReceipt;OriginalFileChanged=$false}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root instrumentation.json)
Write-Output ('Prepared diagnostic chain: '+$root)
