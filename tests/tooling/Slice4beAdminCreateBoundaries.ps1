# Optional observation faults exercise the actual Create handler and Core owner.
function Test-AdminCreateBoundaries($Fixture,$Other,$HostBook){
    $otherBook=$excel.Workbooks.Add();$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Create boundary canary'
    try{
        foreach($case in @('Disabled','ControlOff','UserOff','SwitchWorkbook','SignOut','SwitchTarget','AttemptWriteFailure','OutcomeWriteFailure')){
            SelectTarget $Fixture
            $policy=if($case -cin @('Disabled','ControlOff','UserOff')){$case}else{'Enabled'}
            if(-not [bool](Probe 'SetPolicy' @($policy))){throw 'Create boundary policy prerequisite unavailable.'}
            [void](Invoke 'OpenForm' @($HostBook.Name))
            $id='CB'+[guid]::NewGuid().ToString('N').Substring(0,8)
            $root=Join-Path $runRoot ('create-boundary-'+$case);$share=Join-Path $runRoot ('create-share-'+$case)
            [void](Invoke 'Prepare' @($id,$root,$share,'Training'))
            [void](Invoke 'ResetTrackingNotices')
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $fault=if($policy -ceq 'Enabled'){$case}else{''}
            [void](Probe 'ArmBoundary' @($fault,$Other.Root,$Other.Secret,$otherBook.Name))
            try{
                [void](Invoke 'Create')
                $facts=[string](Invoke 'Facts')
                $reached=if($fault -like '*WriteFailure'){[int](Probe 'WriteFailures') -eq 1}elseif($fault){[bool](Probe 'EntryObserved')}else{$true}
                if(-not $reached){throw 'Create observation boundary was not reached; not behavioral RED.'}
                $prefix='Create.Boundary.'+$case+'.';$denied=$case -cin @('SignOut','SwitchTarget')
                Check ($prefix+'OwnerTruth') ($facts -ceq $(if($denied){'False|False|Create|False|0|0'}else{'True|True|Close|True|1|1'}))
                Check ($prefix+'DestinationTruth') ((Test-Path -LiteralPath $root) -eq (-not $denied))
                $notices=if($case -cin @('Disabled','ControlOff','UserOff','SwitchWorkbook')){0}else{1}
                Check ($prefix+'OneNoticeOrNone') ([int](Invoke 'TrackingNotices') -eq $notices)
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_CREATE_SUBMIT')
                $honest=if($case -ceq 'SwitchWorkbook'){$records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'CREATED_PUBLISHED').Count -eq 1}elseif($case -ceq 'OutcomeWriteFailure'){$records.Count -eq 1 -and $records[0].OutcomeCode -ceq 'REQUESTED'}else{$records.Count -eq 0}
                foreach($record in $records){$honest=$honest -and $record.WarehouseId -ceq $Fixture.Warehouse -and @($record.SourceEventRefs).Count -eq 0}
                Check ($prefix+'HonestSourceEvidence') $honest
                Check ($prefix+'NoRedirectedEvidence') ((@(Get-Slice4beActivityFiles $Other)-join '|') -ceq ($otherBefore -join '|'))
                Check ($prefix+'OtherWorkbookPreserved') ([string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Create boundary canary')
            }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');SelectTarget $Fixture}
        }
        if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Create revocation policy unavailable.'}
        $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb');$original=[IO.File]::ReadAllBytes($auth)
        try{
            [void](Invoke 'OpenForm' @($HostBook.Name))
            $root=Join-Path $runRoot 'recording-revoked';[void](Invoke 'Prepare' @('REVOKED',$root,'','Training'))
            $book=$excel.Workbooks.Open($auth,0,$false);$caps=Table $book 'tblCapabilities';$changed=0
            foreach($row in $caps.ListRows){
                if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){
                    $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='INACTIVE';$changed++
                }
            }
            if(-not $changed){throw 'Create saved-revocation fixture unavailable.'}
            $book.Save();$book.Close($false)
            Observe 'SUBMIT' 'DENIED' {[void](Invoke 'Create')} 'SavedRevocation'
            Check 'Create.Recording.SavedRevocationNoDestination' (-not (Test-Path -LiteralPath $root))
        }finally{[void](Invoke 'CloseForm');[IO.File]::WriteAllBytes($auth,$original);SelectTarget $Fixture}
    }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');$otherBook.Close($false);SelectTarget $Fixture}
}
