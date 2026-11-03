# D8-INV-CONFIG-01 retained shared/automation contracts. This complements, rather
# than replaces, the named owning-caller matrix and packaged operator regressions.
function Install-InventoryConfigRetainedProbe {
    $project=$packages['invSys.Core.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryConfigRetained'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function Shared(ByVal warehouse As String, ByVal station As String, _
        ByVal role As String, ByVal config As String) As String
    Dim path As String, report As String
    If modConfig.EnsureStationInbox(warehouse, station, role, config, path, report) Then
        Shared = "OK|" & path
    Else
        Shared = "FAIL"
    End If
End Function
'@)
}

function Test-InventoryConfigRetainedShared($Fixture) {
    $paths=@{}
    foreach($entry in @('Shared','Automation','Packed')){
        foreach($role in @('RECEIVE','SHIP','PRODUCTION')){
            $id='InventoryConfigRetained.'+$entry+'.'+$role
            SelectTarget $Fixture
            Set-InventoryConfigPresentation $Fixture
            $before=Get-InventoryConfigDiagnosticSnapshot $Fixture
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            $arguments=@($Fixture.Warehouse,'S1',$role,$Fixture.Config)
            $macro=switch($entry){
                'Shared' {'TestInventoryConfigRetained.Shared'}
                'Automation' {'modConfig.EnsureStationInboxForAutomation'}
                'Packed' {'modConfig.EnsureStationInboxPackedForAutomation'}
            }
            if($entry -ceq 'Packed'){$arguments=@($arguments -join '|')}
            $result=[string](Run 'invSys.Core.xlam' $macro $arguments)
            Check ($id+'.Succeeded') ($result.StartsWith('OK|',[StringComparison]::Ordinal))
            if(-not $result.StartsWith('OK|',[StringComparison]::Ordinal)){continue}
            $path=$result.Substring(3)
            $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
            $owned=[IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)
            Check ($id+'.InboxWithinOwnedFixture') ($owned -and (Test-Path -LiteralPath $path -PathType Leaf))
            if(-not $owned -or -not (Test-Path -LiteralPath $path -PathType Leaf)){throw 'Retained inbox is outside the owned test fixture.'}
            if($entry -ceq 'Shared'){$paths[$role]=$path}
            Check ($id+'.SameRoleDestinationAcrossEntries') ($paths.ContainsKey($role) -and $paths[$role] -ceq $path)
            $after=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $changes=@(@($before.Cells.Keys)+@($after.Cells.Keys)|Select-Object -Unique|Where-Object {
                -not $before.Cells.ContainsKey($_) -or -not $after.Cells.ContainsKey($_) -or $before.Cells[$_] -cne $after.Cells[$_]
            })
            Check ($id+'.ConfigValuesAndShapePreserved') ($changes.Count -eq 0 -and $before.Shapes -ceq $after.Shapes)
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            Check ($id+'.RetainedConfigSave') ($saveFacts.Count -eq 3 -and [int]$saveFacts[0] -eq 1 -and [int]$saveFacts[2] -eq 0)
            $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
            $sheets=$null;$sheet=$null;$columns=$null;$column=$null;$rows=$null;$row=$null;$font=$null
            try{
                $sheets=$book.Worksheets;$sheet=$sheets.Item(1);$columns=$sheet.Columns;$column=$columns.Item(1)
                $rows=$sheet.Rows;$row=$rows.Item(1);$font=$row.Font
                Check ($id+'.RetainedAutoFit') ([double]$column.ColumnWidth -ne 43)
                Check ($id+'.RetainedHeaderBold') ($font.Bold -eq $true)
            }finally{
                foreach($ownedCom in @($font,$row,$rows,$column,$columns,$sheet,$sheets)){if($null -ne $ownedCom){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ownedCom)}}
                $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
            }
            Check ($id+'.InspectionPreservedConfigBytes') ($after.ReadOnlyProof -and $after.FileHash -ceq (Get-AdminInventoryEntryHash $Fixture.Config))
            $inboxHash=Get-AdminInventoryEntryHash $path
            $book=$excel.Workbooks.Open($path,0,$true);$table=$null
            try{
                $name=switch($role){'RECEIVE'{'tblInboxReceive'} 'SHIP'{'tblInboxShip'} 'PRODUCTION'{'tblInboxProd'}}
                $table=Table $book $name
                Check ($id+'.ExpectedInboxTable') ($null -ne $table)
            }finally{
                if($null -ne $table){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($table)}
                $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
            }
            Check ($id+'.InspectionPreservedInboxBytes') ($inboxHash -ceq (Get-AdminInventoryEntryHash $path))
        }
    }
}
