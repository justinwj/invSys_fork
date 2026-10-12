# Pure preflight tests: call the real installer, stop before any probe mutation.
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'ProjectionBoundaryTrace.ps1')
$root=Join-Path (Resolve-Path (Join-Path $PSScriptRoot '../..')) ('reports/runtime/projection-trace-anchors/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
function Fixture([string]$Mode){
    $packages=@{}
    foreach($package in (Get-ProjectionBoundaryTracePlan|Group-Object Package)){
        $components=[pscustomobject]@{Modules=@{};MutationAttempts=0}
        $components|Add-Member ScriptMethod Item {param($name) return $this.Modules[$name]}
        $components|Add-Member ScriptMethod Add {param($kind) $this.MutationAttempts++;throw 'TracePreflightPassed'}
        foreach($group in ($package.Group|Group-Object Module)){
            $module=[pscustomobject]@{Text=[Collections.Generic.List[string]]::new();Procedures=@{}}
            foreach($procedure in ($group.Group|Group-Object Procedure)){
                $start=$module.Text.Count+1
                foreach($step in $procedure.Group){
                    $line='    '+$step.Anchor
                    if($step.Stage -ceq 'modInventorySchema.EnsureTableWithHeaders.7'){
                        switch($Mode){
                            'Missing'{continue}
                            'CaseOnly'{$line='    lo.name = tablename'}
                            'Duplicate'{$module.Text.Add($line)}
                            'CaseDuplicate'{$module.Text.Add('    lo.name = tablename')}
                        }
                        if($Mode -ceq 'Missing'){continue}
                    }
                    $module.Text.Add($line)
                }
                $module.Procedures[$procedure.Name]=[pscustomobject]@{Start=$start;Count=$module.Text.Count-$start+1}
            }
            $module|Add-Member ScriptMethod ProcStartLine {param($name,$kind) return $this.Procedures[$name].Start}
            $module|Add-Member ScriptMethod ProcCountLines {param($name,$kind) return $this.Procedures[$name].Count}
            $module|Add-Member ScriptMethod Lines {param($line,$count) return $this.Text[$line-1]}
            $components.Modules[$group.Name]=[pscustomobject]@{CodeModule=$module}
        }
        $packages[$package.Name]=[pscustomobject]@{VBProject=[pscustomobject]@{VBComponents=$components}}
    }
    return $packages
}
Check 'Plan.Stable37Stages' (@(Get-ProjectionBoundaryTracePlan).Count -eq 37)
foreach($mode in @('Exact','CaseOnly','Missing','Duplicate','CaseDuplicate')){
    $packages=Fixture $mode;$reached=$false;$refused=$false
    try{Install-ProjectionBoundaryTrace -Excel $null -Packages $packages -PackageRoot 'unused'}catch{
        $reached=$_.Exception.Message.Contains('TracePreflightPassed')
        $refused=$_.Exception.Message.Contains('Trace anchor missing/ambiguous: modInventorySchema.EnsureTableWithHeaders.7;')
    }
    $mutations=0;foreach($package in $packages.Values){$mutations+=$package.VBProject.VBComponents.MutationAttempts}
    if($mode -cin @('Exact','CaseOnly')){
        Check ($mode+'.ValidPreflight') $reached
        Check ($mode+'.StopsBeforeProbeWrite') ($mutations -eq 1)
    }else{
        Check ($mode+'.RejectedAtExactStage') $refused
        Check ($mode+'.NoProbeMutation') ($mutations -eq 0)
    }
}
$checks.ToArray()|ConvertTo-Json|Set-Content (Join-Path $root 'checks.json')
$failed=@($checks|Where-Object {-not $_.Passed})
[pscustomobject]@{ReportRoot=$root;Passed=$checks.Count-$failed.Count;Failed=$failed.Count;ExcelCreated=$false;ProductRedClaimed=$false}|ConvertTo-Json
if($failed.Count){$failed|ForEach-Object {Write-Output ($_.Check+': FAIL')};exit 1}
