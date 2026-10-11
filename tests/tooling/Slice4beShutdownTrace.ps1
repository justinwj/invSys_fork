# Redact the diagnostic's explicitly read generated trigger fields before output.
# The report tail can contain identifiers and errors: never retain or serialize it.
function Convert-ProjectionOutcomeDiagnostic([string]$Report,[string]$Status,[string]$ErrorCode,[int]$BeforeLog,[int]$AfterLog,[int]$BeforeApplied,[int]$AfterApplied,[int]$TriggerMatches,[bool]$ExactIdentity) {
    $match=[regex]::Match($Report,'\AProcessed=([0-9]{1,9}); Report=Applied=([0-9]{1,9}); SkipDup=([0-9]{1,9}); Poison=([0-9]{1,9})(?:;|\z)')
    $counts=@($null,$null,$null,$null)
    if($match.Success){for($i=0;$i -lt 4;$i++){$counts[$i]=[int]$match.Groups[$i+1].Value}}
    $safeStatus='OTHER';$safeError='OTHER'
    if($Status -ceq ''){$safeStatus='EMPTY'}elseif($Status -cin @('NEW','PROCESSED','SKIP_DUP','POISON')){$safeStatus=$Status}
    if($ErrorCode -ceq ''){$safeError='NONE'}elseif($ErrorCode -cin @('AUTH_DENIED','INVALID_EVENT','INVALID_EVENT_TYPE','DESIGNS_WORKBOOK_NOT_FOUND','UNKNOWN_APPLY_STATUS')){$safeError=$ErrorCode}
    return [pscustomobject]@{ReportRecognized=$match.Success;Processed=$counts[0];Applied=$counts[1];SkipDup=$counts[2];Poison=$counts[3];Status=$safeStatus;ErrorClass=$safeError;LogDelta=$AfterLog-$BeforeLog;AppliedDelta=$AfterApplied-$BeforeApplied;TriggerMatches=$TriggerMatches;ExactIdentity=$ExactIdentity}
}

# Diagnostic metadata only. Never pass workbook values, macro arguments or errors.
function Add-ShutdownPhase([string]$Phase,[string]$Marker='',[bool]$Passed=$false) {
    if($Phase -cnotmatch '^(Restart\.(Before|After)|Stage\.(Before|After)|Macro\.Before|Check|Workbook\.(BeforeClose|AfterClose|BeforeRelease|AfterRelease)|Excel\.(BeforeQuit|AfterQuitAttempt|BeforeRelease|AfterRelease))$'){throw 'Unknown shutdown phase.'}
    if($Marker -cnotmatch '^[A-Za-z0-9_.-]{0,128}$'){throw 'Invalid shutdown marker.'}
    $line=[pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Phase=$Phase;Marker=$Marker;Passed=$Passed}|ConvertTo-Json -Compress
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes($line+[Environment]::NewLine)
    for($attempt=0;$attempt -lt 100;$attempt++){
        $stream=$null
        try{
            # Permit readers, serialize writers, and keep each JSON line whole.
            $stream=[IO.FileStream]::new((Join-Path $ReportRoot 'phases.jsonl'),[IO.FileMode]::Append,[IO.FileAccess]::Write,[IO.FileShare]::Read)
            $stream.Write($bytes,0,$bytes.Length)
            return
        }catch [IO.IOException]{
            $code=$_.Exception.HResult -band 0xffff
            if($code -notin @(32,33) -or $attempt -eq 99){throw}
            Start-Sleep -Milliseconds 10
        }finally{if($null -ne $stream){$stream.Dispose()}}
    }
}

function Add-ShutdownTraceToFunction([string]$Name,[string]$Text) {
    $edits=@()
    switch -CaseSensitive ($Name) {
        'Run-WorkbookMacro' {$edits+=,@('$macro = "', ('Add-ShutdownPhase ''Macro.Before'' $MacroName'+"`r`n"+'    $macro = "'))}
        'Add-Result' {$edits+=,@('$results.Add(', ('Add-ShutdownPhase ''Check'' $Check $Passed'+"`r`n"+'    $results.Add('))}
        'Invoke-RestartReconciliation' {
            $edits+=,@('try { $workbook.Close($false) } catch {}', ('Add-ShutdownPhase ''Workbook.BeforeClose'''+"`r`n"+'            try { $workbook.Close($false) } catch {}'+"`r`n"+'            Add-ShutdownPhase ''Workbook.AfterClose'''))
            $edits+=,@('Release-ComObject $workbook', ('Add-ShutdownPhase ''Workbook.BeforeRelease'''+"`r`n"+'            Release-ComObject $workbook'+"`r`n"+'            Add-ShutdownPhase ''Workbook.AfterRelease'''))
            $edits+=,@('try { $localExcel.Quit() } catch {}', ('Add-ShutdownPhase ''Excel.BeforeQuit'''+"`r`n"+'        try { $localExcel.Quit() } catch {}'+"`r`n"+'        Add-ShutdownPhase ''Excel.AfterQuitAttempt'''))
            $edits+=,@('Release-ComObject $localExcel', ('Add-ShutdownPhase ''Excel.BeforeRelease'''+"`r`n"+'        Release-ComObject $localExcel'+"`r`n"+'        Add-ShutdownPhase ''Excel.AfterRelease'''))
        }
    }
    foreach($edit in $edits){
        if([regex]::Matches($Text,[regex]::Escape($edit[0])).Count -ne 1){throw 'Shutdown trace anchor drift; no diagnostic dispatched.'}
        $Text=$Text.Replace($edit[0],$edit[1])
    }
    return $Text
}

# The ordered live validator owns its Excel instance. Instrument its generated
# source before dispatch; never attach a second automation client for tracing.
function Add-LiveProjectionTrace([string]$Source,[string]$TraceRoot) {
    if($Source.Contains('Add-ShutdownPhase') -or $Source.Contains("Owner='LiveRole.excel'")){throw 'Live validator is already traced.'}
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Live validator does not parse.'}
    $main=@($ast.EndBlock.Statements|Where-Object{$_ -is [Management.Automation.Language.TryStatementAst]})
    if($main.Count -ne 1){throw 'Live main boundary differs.'}
    $statements=@($main[0].Body.Statements)
    $begin=@($statements|Where-Object{$_.Extent.Text -ceq '$currentStep = "Delete and rebuild canonical inventory projections"'})
    $end=@($statements|Where-Object{$_.Extent.Text -ceq '$currentStep = "Run two consecutive Production batches through form actions"'})
    if($begin.Count -ne 1 -or $end.Count -ne 1 -or $begin[0].Extent.StartOffset -ge $end[0].Extent.StartOffset){throw 'Projection trace boundaries differ.'}
    $selected=@($statements|Where-Object{$_.Extent.StartOffset -ge $begin[0].Extent.StartOffset -and $_.Extent.StartOffset -lt $end[0].Extent.StartOffset})
    $map=[Collections.Generic.List[object]]::new()
    for($index=$selected.Count-1;$index -ge 0;$index--){
        $statement=$selected[$index];$marker='Projection.'+($index+1).ToString('00')
        $map.Add([pscustomobject]@{Marker=$marker;SourceLine=$statement.Extent.StartLineNumber;Kind=$statement.GetType().Name})
        $Source=$Source.Insert($statement.Extent.EndOffset,("`r`n    Add-ShutdownPhase 'Stage.After' '"+$marker+"'"))
        $Source=$Source.Insert($statement.Extent.StartOffset,("Add-ShutdownPhase 'Stage.Before' '"+$marker+"'`r`n    "))
    }
    $setup='$ReportRoot = '''+$TraceRoot.Replace("'","''")+''' '+"`r`n"+'. (Join-Path $repo ''tests/tooling/Slice4beShutdownTrace.ps1'')'
    $owner=@'
    [uint32]$traceProcessId=0
    [void][InvSysLiveValidationWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$traceProcessId)
    $traceProcess=Get-Process -Id $traceProcessId -ErrorAction Stop
    [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Owner='LiveRole.excel';ProcessId=$traceProcessId;CreatedUTC=$traceProcess.StartTime.ToUniversalTime().ToString('o')}|ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $ReportRoot 'owners.jsonl')
'@
    $edits=@(
        @('$repo = (Resolve-Path $RepoRoot).Path', ('$repo = (Resolve-Path $RepoRoot).Path'+"`r`n"+$setup)),
        @('$excel = New-Object -ComObject Excel.Application', ('$excel = New-Object -ComObject Excel.Application'+"`r`n"+$owner)),
        @('$fullMacro = "', ('Add-ShutdownPhase ''Macro.Before'' $MacroName'+"`r`n    "+'$fullMacro = "')),
        @('$Rows.Add(', ('Add-ShutdownPhase ''Check'' $Check $Passed'+"`r`n    "+'$Rows.Add('))
    )
    foreach($edit in $edits){
        if([regex]::Matches($Source,[regex]::Escape($edit[0])).Count -ne 1){throw 'Live trace anchor drift.'}
        $Source=$Source.Replace($edit[0],$edit[1])
    }
    $null=[Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Traced live validator does not parse.'}
    $map|Sort-Object Marker|ConvertTo-Json|Set-Content (Join-Path $TraceRoot 'projection-map.json')
    return $Source
}

function Convert-ShutdownFixturePath([string]$Value,[string]$SourceRoot,[string]$CloneRoot) {
    if([string]::IsNullOrWhiteSpace($Value)){return $Value}
    $old=[IO.Path]::GetFullPath($SourceRoot).TrimEnd('\')
    $new=[IO.Path]::GetFullPath($CloneRoot).TrimEnd('\')
    $path=[IO.Path]::GetFullPath($Value).TrimEnd('\')
    if($path.Equals($new,[StringComparison]::OrdinalIgnoreCase) -or $path.StartsWith($new+'\',[StringComparison]::OrdinalIgnoreCase)){return $path}
    if($path.Equals($old,[StringComparison]::OrdinalIgnoreCase)){return $new}
    if($path.StartsWith($old+'\',[StringComparison]::OrdinalIgnoreCase)){return $new+$path.Substring($old.Length)}
    throw 'Fixture path is outside its recorded source and clone roots.'
}

function Initialize-ShutdownClonePaths($Workbook,[string]$SourceRoot,[string]$CloneRoot) {
    if(-not ([IO.Path]::GetDirectoryName([string]$Workbook.FullName)).Equals([IO.Path]::GetFullPath($CloneRoot).TrimEnd('\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Only the cloned configuration may be initialized.'}
    $writes=[Collections.Generic.List[object]]::new()
    try {
        foreach($pair in @(@('tblWarehouseConfig','PathDataRoot',$true),@('tblWarehouseConfig','PathBackupRoot',$false),@('tblStationConfig','PathInboxRoot',$false))){
            $table=$null;$columns=$null;$column=$null;$body=$null;$cells=$null;$rows=$null
            try {
                $table=Get-ListObject -Workbook $Workbook -TableName $pair[0]
                if($null -eq $table){throw 'Expected cloned configuration table missing.'}
                $columns=$table.ListColumns
                try{$column=$columns.Item($pair[1])}catch{if($pair[2]){throw 'Required cloned path header missing.'}}
                if($null -eq $column){continue}
                $body=$column.DataBodyRange
                if($null -eq $body){throw 'Cloned path rows missing.'}
                $rows=$body.Rows;$cells=$body.Cells
                for($index=1;$index -le $rows.Count;$index++){
                    $cell=$cells.Item($index,1)
                    try{
                        $value=Convert-ShutdownFixturePath ([string]$cell.Value2) $SourceRoot $CloneRoot
                        $writes.Add([pscustomobject]@{Cell=$cell;Value=$value});$cell=$null
                    }finally{Release-ComObject $cell}
                }
            }finally{foreach($item in @($cells,$rows,$body,$column,$columns,$table)){Release-ComObject $item}}
        }
        # All paths are validated before any cell is changed; other columns stay intact.
        foreach($write in $writes){$write.Cell.Value2=$write.Value}
        $Workbook.Save()
        [pscustomobject]@{ClonedConfigOnly=$true;PathCellsInitialized=$writes.Count;OriginalRootValuesLogged=$false}|ConvertTo-Json|Set-Content (Join-Path $ReportRoot 'fixture-paths.json')
    }finally{foreach($write in $writes){Release-ComObject $write.Cell}}
}
