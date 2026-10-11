# Pure PowerShell checks for diagnostic instrumentation; no Excel is created.
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Slice4beShutdownTrace.ps1')
$ReportRoot=Join-Path (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path ('reports/runtime/shutdown-trace-contract/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $ReportRoot|Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check($Name,$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=[bool]$Passed})}
$source=Join-Path $PSScriptRoot '../../tools/validate_release1_full_chain.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Resolve-Path $source),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Chain source does not parse.'}
$definitions=$ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)
foreach($definition in $definitions){
    $changed=Add-ShutdownTraceToFunction $definition.Name $definition.Extent.Text
    $null=[Management.Automation.Language.Parser]::ParseInput($changed,[ref]$tokens,[ref]$errors)
    Check ('Parse.'+$definition.Name) ($errors.Count -eq 0)
    if($definition.Name -ceq 'Run-WorkbookMacro'){$macroText=$changed;. ([scriptblock]::Create($changed))}
    if($definition.Name -cnotin @('Run-WorkbookMacro','Add-Result','Invoke-RestartReconciliation')){Check ('Unchanged.'+$definition.Name) ($changed -ceq $definition.Extent.Text)}
}
$payload='private-fixture-value-'+[guid]::NewGuid().ToString('N')
$fake=[pscustomobject]@{}
$fake|Add-Member ScriptMethod Run {param($macro,$value) return $value}
Check 'ReturnValuePreserved' ((Run-WorkbookMacro $fake 'invSys.Core.xlam' 'modProbe.Read' @($payload)) -ceq $payload)
$fake|Add-Member -Force ScriptMethod Run {throw [InvalidOperationException]::new('Synthetic invocation failure')}
$thrown=$false
try{Run-WorkbookMacro $fake 'invSys.Core.xlam' 'modProbe.Fail' @($payload)|Out-Null}catch{$thrown=$true}
Check 'ExceptionNotSwallowed' $thrown
$refused=$false
try{Add-ShutdownTraceToFunction 'Run-WorkbookMacro' ($macroText+$macroText)|Out-Null}catch{$refused=$true}
Check 'AmbiguousAnchorRefused' $refused
Add-ShutdownPhase 'Check' 'Sample.Expected' $true
$rows=@(Get-Content (Join-Path $ReportRoot phases.jsonl)|ForEach-Object{$_|ConvertFrom-Json})
$raw=Get-Content (Join-Path $ReportRoot phases.jsonl) -Raw
Check 'ArgumentsAndWorkbookNameAbsent' (-not $raw.Contains($payload) -and -not $raw.Contains('invSys.Core.xlam'))
Check 'ExactMetadataFields' (@($rows|Where-Object{($_.PSObject.Properties.Name -join '|') -cne 'UTC|Phase|Marker|Passed'}).Count -eq 0)
Check 'ExpectedBoundaries' (($rows.Phase -join '|') -ceq 'Macro.Before|Macro.Before|Check' -and ($rows.Marker -join '|') -ceq 'modProbe.Read|modProbe.Fail|Sample.Expected')
$refused=$false
try{Add-ShutdownPhase 'Macro.Before' "unexpected`ntext"}catch{$refused=$true}
Check 'UnsafeMarkerRefused' $refused
$old=Join-Path ([IO.Path]::GetTempPath()) 'invsys-fixture-source'
$clone=Join-Path ([IO.Path]::GetTempPath()) 'invsys-fixture-clone'
Check 'RootRebound' ((Convert-ShutdownFixturePath $old $old $clone) -ceq $clone)
Check 'BackupSubpathRebound' ((Convert-ShutdownFixturePath (Join-Path $old 'Backups') $old $clone) -ceq (Join-Path $clone 'Backups'))
Check 'BlankStationFallbackRetained' ((Convert-ShutdownFixturePath '' $old $clone) -ceq '')
Check 'ExistingClonePathRetained' ((Convert-ShutdownFixturePath $clone $old $clone) -ceq $clone)
$refused=$false
try{Convert-ShutdownFixturePath ($old+'-outside') $old $clone|Out-Null}catch{$refused=$true}
Check 'SiblingPrefixRefused' $refused
$refused=$false
try{Convert-ShutdownFixturePath (Join-Path $old '../outside') $old $clone|Out-Null}catch{$refused=$true}
Check 'TraversalOutsideSourceRefused' $refused

# Generate scripts without dispatching Excel. Compare original business calls,
# and execute only the partial report writer to protect its diagnostic scope.
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$generator=Join-Path $PSScriptRoot 'New-Slice4beChainTrace.ps1'
$sourceHash=(Get-FileHash $source).Hash
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('invsys-phase6-live-trace-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture|Out-Null
try {
    $sentinel=Join-Path $fixture 'unknown-column-sentinel.txt'
    [IO.File]::WriteAllText($sentinel,'Synthetic generator fixture; no workbook or operational values.')
    $sentinelHash=(Get-FileHash $sentinel).Hash
    $receipt=Join-Path $ReportRoot 'synthetic-receipt.md'
    Set-Content $receipt ('- Runtime root override: '+$fixture)
    $fullRoot=Join-Path $ReportRoot 'full'
    $lateRoot=Join-Path $ReportRoot 'late'
    & $generator -OutputDirectory $fullRoot|Out-Null
    & $generator -OutputDirectory $lateRoot -LateFixtureReceipt $receipt|Out-Null
    $full=[Management.Automation.Language.Parser]::ParseFile((Join-Path $fullRoot 'chain.ps1'),[ref]$tokens,[ref]$errors)
    Check 'GeneratedFullParses' ($errors.Count -eq 0)
    $late=[Management.Automation.Language.Parser]::ParseFile((Join-Path $lateRoot 'chain.ps1'),[ref]$tokens,[ref]$errors)
    Check 'GeneratedPartialParses' ($errors.Count -eq 0)
    function BusinessCalls($Tree){
        return (@($Tree.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -cin @('Run-WorkbookMacro','Add-Result','Add-ResultRow','Restore-LiveRuntimeContext','Add-ListObjectRow')},$true)|ForEach-Object{$_.Extent.Text}) -join "`n")
    }
    Check 'FullBusinessCallsUnchanged' ((BusinessCalls $ast) -ceq (BusinessCalls $full))
    $orderedWriter=@($full.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)|Where-Object Name -CEQ 'New-OrderedLiveValidator')[0]
    & {
        param($Definition,$Destination)
        $ReportRoot=$Destination;$tempRoot=Join-Path $Destination 'ordered-test'
        New-Item -ItemType Directory $tempRoot|Out-Null
        . ([scriptblock]::Create($Definition.Extent.Text))
        New-OrderedLiveValidator|Out-Null
    } $orderedWriter $fullRoot
    $originalLive=[Management.Automation.Language.Parser]::ParseFile((Join-Path $fullRoot 'ordered-live-original.ps1'),[ref]$tokens,[ref]$errors)
    $tracedLive=[Management.Automation.Language.Parser]::ParseFile((Join-Path $fullRoot 'ordered-live-traced.ps1'),[ref]$tokens,[ref]$errors)
    Check 'TracedOrderedLiveParses' ($errors.Count -eq 0)
    Check 'LiveBusinessCallsUnchanged' ((BusinessCalls $originalLive) -ceq (BusinessCalls $tracedLive))
    function ProjectionStatements($Tree){
        $body=@($Tree.EndBlock.Statements|Where-Object{$_ -is [Management.Automation.Language.TryStatementAst]})[0].Body
        $began=$false
        foreach($statement in $body.Statements){
            if($statement.Extent.Text -ceq '$currentStep = "Delete and rebuild canonical inventory projections"'){$began=$true}
            if($statement.Extent.Text -ceq '$currentStep = "Run two consecutive Production batches through form actions"'){break}
            if($began -and $statement.Extent.Text -notlike 'Add-ShutdownPhase *'){$statement.Extent.Text}
        }
    }
    $projectionStatements=@(ProjectionStatements $originalLive)
    Check 'EveryProjectionStatementUnchanged' (($projectionStatements -join "`n") -ceq ((ProjectionStatements $tracedLive) -join "`n"))
    $map=Get-Content (Join-Path $fullRoot 'projection-map.json') -Raw|ConvertFrom-Json
    Check 'ProjectionMapMetadataOnly' (@($map).Count -eq $projectionStatements.Count -and @($map|Where-Object{($_.PSObject.Properties.Name -join '|') -cne 'Marker|SourceLine|Kind'}).Count -eq 0)
    $tracedText=Get-Content (Join-Path $fullRoot 'ordered-live-traced.ps1') -Raw
    Check 'EveryProjectionBoundaryPaired' ([regex]::Matches($tracedText,"Add-ShutdownPhase 'Stage.Before' 'Projection\.[0-9]+'").Count -eq $map.Count -and [regex]::Matches($tracedText,"Add-ShutdownPhase 'Stage.After' 'Projection\.[0-9]+'").Count -eq $map.Count)
    Check 'LiveOwnerCapturedOnce' ([regex]::Matches($tracedText,"Owner='LiveRole.excel'").Count -eq 1)
    $refused=$false
    try{Add-LiveProjectionTrace $tracedText $fullRoot|Out-Null}catch{$refused=$true}
    Check 'DuplicateLiveTracingRefused' $refused
    foreach($name in @('Invoke-RestartReconciliation','Invoke-RuntimeFivePackageEvidence','Invoke-StaticRetiredPathRatchet')){
        $before=@($definitions|Where-Object Name -CEQ $name)[0]
        $after=@($late.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)|Where-Object Name -CEQ $name)[0]
        Check ('PartialBusinessCallsUnchanged.'+$name) ((BusinessCalls $before) -ceq (BusinessCalls $after))
    }
    $main=@($late.EndBlock.Statements|Where-Object{$_ -is [Management.Automation.Language.TryStatementAst]})[0].Body
    $stages=@($main.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -like 'Invoke-*'},$true)|ForEach-Object{$_.GetCommandName()})
    Check 'OnlyDeclaredPartialStages' (($stages -join '|') -ceq 'Invoke-RestartReconciliation|Invoke-RuntimeFivePackageEvidence|Invoke-StaticRetiredPathRatchet')
    $writer=@($late.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)|Where-Object Name -CEQ 'Write-Evidence')[0]
    $report=& {
        param($Definition,$Destination)
        $results=@([pscustomobject]@{Check='Synthetic.Check';Passed=$true;Detail='Synthetic diagnostic assertion.'})
        $phaseOrder=@('DiagnosticRestart','DiagnosticRuntimeEvidence','StaticAudit')
        $resultPath=$Destination
        . ([scriptblock]::Create($Definition.Extent.Text))
        Write-Evidence
        Get-Content $resultPath -Raw
    } $writer (Join-Path $lateRoot 'writer-test.md')
    Check 'PartialReportDoesNotClaimFullChain' ($report.Contains('# Partial Restart and Runtime Evidence Diagnostic') -and $report.Contains('No product RED or release acceptance is claimed.') -and -not $report.Contains('## D13 trace') -and -not $report.Contains('30/30'))
    Check 'OriginalChainPreserved' ((Get-FileHash $source).Hash -ceq $sourceHash)
    Check 'SourceAndUnknownFixtureFilePreserved' ((Get-FileHash $sentinel).Hash -ceq $sentinelHash -and (Get-FileHash (Join-Path $lateRoot 'fixture/unknown-column-sentinel.txt')).Hash -ceq $sentinelHash)
}finally{
    $resolvedFixture=(Resolve-Path -LiteralPath $fixture).Path
    $allowed=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\invsys-phase6-live-trace-'
    if(-not $resolvedFixture.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Synthetic cleanup outside its allowed root.'}
    Remove-Item -LiteralPath $resolvedFixture -Recurse -Force
}
$outcomeCommand=Get-Command Convert-ProjectionOutcomeDiagnostic -ErrorAction SilentlyContinue
Check 'ProjectionOutcomeDiagnosticAvailable' ($null -ne $outcomeCommand)
if($null -ne $outcomeCommand){
    $sample=Convert-ProjectionOutcomeDiagnostic ('Processed=0; Report=Applied=0; SkipDup=0; Poison=1; RunId='+$payload+'; LastPoison='+$payload) 'POISON' 'AUTH_DENIED' 6 6 6 6 0 $false
    Check 'ProjectionOutcomeCounts' ($sample.Processed -eq 0 -and $sample.Applied -eq 0 -and $sample.SkipDup -eq 0 -and $sample.Poison -eq 1 -and $sample.ReportRecognized)
    Check 'ProjectionOutcomeClassification' ($sample.Status -ceq 'POISON' -and $sample.ErrorClass -ceq 'AUTH_DENIED' -and $sample.LogDelta -eq 0 -and $sample.AppliedDelta -eq 0 -and $sample.TriggerMatches -eq 0 -and -not $sample.ExactIdentity)
    $unknown=Convert-ProjectionOutcomeDiagnostic $payload $payload $payload 6 7 6 7 1 $true
    Check 'ProjectionUnknownValuesRedacted' ($unknown.Status -ceq 'OTHER' -and $unknown.ErrorClass -ceq 'OTHER' -and -not $unknown.ReportRecognized -and $null -eq $unknown.Processed)
    $malformed=Convert-ProjectionOutcomeDiagnostic 'Processed=1; Report=Applied=1; SkipDup=0; Poison=0secret' '' '' 0 1 0 1 1 $true
    Check 'ProjectionMalformedCountsNotInferred' (-not $malformed.ReportRecognized)
    $serialized=@($sample,$unknown,$malformed)|ConvertTo-Json -Depth 4
    Check 'ProjectionPayloadNeverSerialized' (-not $serialized.Contains($payload) -and -not $serialized.Contains('RunId') -and -not $serialized.Contains('LastPoison'))
    Check 'ProjectionDiagnosticFixedShape' (($sample.PSObject.Properties.Name -join '|') -ceq 'ReportRecognized|Processed|Applied|SkipDup|Poison|Status|ErrorClass|LogDelta|AppliedDelta|TriggerMatches|ExactIdentity')
}
$controlTree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Test-Slice4beShutdownControl.ps1'),[ref]$tokens,[ref]$errors)
$cleanup=$controlTree.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.VariablePath.UserPath -ceq 'localReference'},$true)
if($null -eq $cleanup){throw 'Projection cleanup loop unavailable.'}
$released=@(& {
    param($text)
    $inbox='caller-owned-sentinel';$releasedReferences=[Collections.Generic.List[string]]::new()
    function Release-ComObject($value){if($null -ne $value){$releasedReferences.Add([string]$value)}}
    & {param($text) $sku='local-owned-sentinel';. ([scriptblock]::Create($text))} $text
    $releasedReferences.ToArray()
} $cleanup.Extent.Text)
Check 'ProjectionCleanupPreservesCallerReferences' ($released -cnotcontains 'caller-owned-sentinel')
Check 'ProjectionCleanupReleasesAcquiredReference' ($released -ccontains 'local-owned-sentinel')
Add-Type @'
using System.IO;using System.Threading;using System.Threading.Tasks;
public static class TraceSharingControl {
 public static Task HoldThenRelease(string path) {
  var stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.Read);
  return Task.Run(()=>{Thread.Sleep(150);stream.Dispose();});
 }
}
'@
$held=[TraceSharingControl]::HoldThenRelease((Join-Path $ReportRoot 'phases.jsonl'))
$appendSucceeded=$false
try{Add-ShutdownPhase 'Stage.Before' 'Contention.Check';$appendSucceeded=$true}catch{}finally{$held.Wait()}
Check 'TraceSurvivesBriefReaderContention' $appendSucceeded
$written=@(Get-Content (Join-Path $ReportRoot 'phases.jsonl')|ForEach-Object{$_|ConvertFrom-Json}|Where-Object Marker -CEQ 'Contention.Check')
Check 'TraceContentionAppendsExactlyOnce' ($written.Count -eq 1)
$checks|ConvertTo-Json|Set-Content (Join-Path $ReportRoot checks.json)
$failed=@($checks|Where-Object{-not $_.Passed}).Count
[pscustomobject]@{ReportRoot=$ReportRoot;Passed=$checks.Count-$failed;Failed=$failed;ExcelCreated=$false}|ConvertTo-Json
if($failed){exit 1}
