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
        return (@($Tree.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -cin @('Run-WorkbookMacro','Add-Result')},$true)|ForEach-Object{$_.Extent.Text}) -join "`n")
    }
    Check 'FullBusinessCallsUnchanged' ((BusinessCalls $ast) -ceq (BusinessCalls $full))
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
$checks|ConvertTo-Json|Set-Content (Join-Path $ReportRoot checks.json)
$failed=@($checks|Where-Object{-not $_.Passed}).Count
[pscustomobject]@{ReportRoot=$ReportRoot;Passed=$checks.Count-$failed;Failed=$failed;ExcelCreated=$false}|ConvertTo-Json
if($failed){exit 1}
