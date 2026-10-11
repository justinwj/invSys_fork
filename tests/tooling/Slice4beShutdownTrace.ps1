# Diagnostic metadata only. Never pass workbook values, macro arguments or errors.
function Add-ShutdownPhase([string]$Phase,[string]$Marker='',[bool]$Passed=$false) {
    if($Phase -cnotmatch '^(Restart\.(Before|After)|Stage\.(Before|After)|Macro\.Before|Check|Workbook\.(BeforeClose|AfterClose|BeforeRelease|AfterRelease)|Excel\.(BeforeQuit|AfterQuitAttempt|BeforeRelease|AfterRelease))$'){throw 'Unknown shutdown phase.'}
    if($Marker -cnotmatch '^[A-Za-z0-9_.-]{0,128}$'){throw 'Invalid shutdown marker.'}
    [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Phase=$Phase;Marker=$Marker;Passed=$Passed}|ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $ReportRoot 'phases.jsonl')
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
