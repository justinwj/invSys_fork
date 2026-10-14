# The shared observer must not block, repeat or redirect the real Tester owner.
function Test-AdminTesterBoundaries($Fixture,$Other,$HostBook){
    $otherBook=$excel.Workbooks.Add();$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Tester boundary canary'
    try{
        foreach($case in @('Disabled','ControlOff','UserOff','SwitchWorkbook','SignOut','SwitchTarget','AttemptWriteFailure','OutcomeWriteFailure')){
            SelectTarget $Fixture
            $policy=if($case -cin @('Disabled','ControlOff','UserOff')){$case}else{'Enabled'}
            if(-not [bool](Probe 'SetPolicy' @($policy))){throw 'Tester boundary policy unavailable.'}
            [void](Invoke 'OpenForm' @($HostBook.Name))
            $id='TEST'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
            $root=Join-Path $runRoot ('tester-boundary-'+$case)
            [void](Invoke 'PrepareSetup' @($id,$root,''));[void](Invoke 'ResetTrackingNotices')
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $fault=if($policy -ceq 'Enabled'){$case}else{''}
            [void](Probe 'ArmBoundary' @($fault,$Other.Root,$Other.Secret,$otherBook.Name))
            try{
                [void](Probe 'ResetOwnerCalls')
                $success=[bool](Invoke 'Setup')
                $reached=if($fault -like '*WriteFailure'){[int](Probe 'WriteFailures') -eq 1}elseif($fault){[bool](Probe 'EntryObserved')}else{$true}
                if(-not $reached){throw 'Tester optional boundary not reached; not behavioral RED.'}
                $prefix='Tester.Boundary.'+$case+'.';$denied=$case -cin @('SignOut','SwitchTarget')
                Check ($prefix+'OwnerTruth') ($success -eq (-not $denied))
                Check ($prefix+'OwnerInvocationCount') ([int](Probe 'OwnerCalls' @('SETUP')) -eq $(if($denied){0}else{1}))
                Check ($prefix+'DestinationTruth') ((Test-Path -LiteralPath (Join-Path $root 'invSys.TesterOwnership.json')) -eq (-not $denied))
                $notices=if($case -cin @('Disabled','ControlOff','UserOff','SwitchWorkbook')){0}else{1}
                Check ($prefix+'OneNoticeOrNone') ([int](Invoke 'TrackingNotices') -eq $notices)
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_TESTER_SETUP')
                $honest=if($case -ceq 'SwitchWorkbook'){$records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'COMPLETED').Count -eq 1}elseif($case -ceq 'OutcomeWriteFailure'){$records.Count -eq 1 -and $records[0].OutcomeCode -ceq 'REQUESTED'}else{$records.Count -eq 0}
                foreach($record in $records){$honest=$honest -and $record.WarehouseId -ceq $Fixture.Warehouse -and @($record.SourceEventRefs).Count -eq 0}
                Check ($prefix+'HonestSourceEvidence') $honest
                Check ($prefix+'NoRedirectedEvidence') ((@(Get-Slice4beActivityFiles $Other)-join '|') -ceq ($otherBefore -join '|'))
                Check ($prefix+'OtherWorkbookPreserved') ([string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Tester boundary canary')
            }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');SelectTarget $Fixture}
        }
    }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');$otherBook.Close($false);SelectTarget $Fixture}
}
