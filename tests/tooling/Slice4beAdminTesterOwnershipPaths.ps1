# Read-only facts and injected session/scope changes around real packaged commands.
function Install-TesterOwnershipProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestTesterFaults'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mFault As String, mOtherRoot As String, mInjected As Boolean
Private mTracking As Boolean, mUser As String, mWarehouse As String, mRoot As String
Private mLeak As Boolean, mPrints As Long
Private mPicker As Boolean, mPickerValue As String
Public Sub Arm(ByVal phase As String, Optional ByVal otherRoot As String = "")
    mFault = phase: mOtherRoot = otherRoot: mInjected = False
End Sub
Public Sub AtPhase(ByVal phase As String)
    If mTracking And phase = "SEED" Then
        modAuth.OwnerDecisionForTest mUser, mWarehouse, mRoot
        modProcessor.OwnerLockDiagnosticForTest mWarehouse
    End If
    If mFault <> phase Or mInjected Then Exit Sub
    mInjected = True
    If mOtherRoot = "" Then modAuth.SignOut Else modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
End Sub
Public Function Injected() As Boolean
    Injected = mInjected
End Function
Public Function Generation() As Long
    Generation = modAuthSession.Version()
End Function
Public Sub ArmPicker(ByVal result As String)
    mPickerValue = result: mPicker = True: mInjected = False
End Sub
Public Function PickerPending() As Boolean
    PickerPending = mPicker
End Function
Public Function PickerResult() As String
    PickerResult = mPickerValue: mPicker = False: mInjected = True
    modAuth.SignOut
End Function
Public Sub StartLogCapture(ByVal user As String, ByVal warehouse As String, ByVal root As String)
    mUser = user: mWarehouse = warehouse: mRoot = root
    mLeak = False: mPrints = 0: mTracking = True
End Sub
Public Sub CapturePrint(ParamArray values() As Variant)
    Dim value As Variant, text As String
    If Not mTracking Then Exit Sub
    mPrints = mPrints + 1
    For Each value In values
        text = CStr(value)
        If InStr(1, text, mUser, vbTextCompare) Or InStr(1, text, mWarehouse, vbTextCompare) Or InStr(1, text, mRoot, vbTextCompare) Then mLeak = True
    Next value
End Sub
Public Function FinishLogCapture() As String
    FinishLogCapture = CStr(mLeak) & "|" & CStr(mPrints)
    mTracking = False: mUser = "": mWarehouse = "": mRoot = ""
End Function
'@)
    foreach($name in @('modAuth','modProcessor','modDiagnostics')){
        $module=$core.VBComponents.Item($name).CodeModule
        $body=$module.Lines(1,$module.CountOfLines)
        $body=$body.Replace('Debug.Print ', 'TestTesterFaults.CapturePrint ')
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($body)
    }
    $core.VBComponents.Item('modAuth').CodeModule.AddFromString(@'
Public Sub OwnerDecisionForTest(ByVal user As String, ByVal warehouse As String, ByVal detail As String)
    LogDecision "", user, "ADMIN_MAINT", warehouse, "S1", "DENY", "TESTER_TEST", detail
End Sub
'@)
    $core.VBComponents.Item('modProcessor').CodeModule.AddFromString(@'
Public Sub OwnerLockDiagnosticForTest(ByVal warehouse As String)
    LogInventoryWorkbookLockDiagnosticsProcessor warehouse
End Sub
'@)
    $module=$core.VBComponents.Item('modTesterProvisioning').CodeModule
    $body=$module.Lines(1,$module.CountOfLines)
    foreach($entry in @(@('CreateTesterRuntimeArtifactsTesterSetup','RUNTIME'),@('SeedTesterScenarioInventory','SEED'),@('ProvisionTesterAuth','AUTH'),@('CreateOrVerifyReceivingWorkbook','WORKBOOK'),@('UpdateConfigSharePointRoot','CONFIG'),@('EnsureTesterSharePointPackage','PUBLISH'))){
        $pattern='(?m)^(    If Not modTesterCommands.CheckPhase\(\) Then GoTo FailSoft\r?\n)(    If Not '+$entry[0]+'\()'
        if($entry[1] -ceq 'RUNTIME'){$pattern='(?m)^(    If Not modTesterCommands.CheckPhase\(\) Then GoTo FailSoft\r?\n)(    If Not runtimeArtifactsExist Then)'}
        if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Tester phase boundary changed.'}
        $body=[regex]::Replace($body,$pattern,('    TestTesterFaults.AtPhase "'+$entry[1]+'"'+"`r`n"+'$1$2'))
    }
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($body)
    $module=$core.VBComponents.Item('modTesterCommands').CodeModule
    $start=$module.ProcStartLine('DeleteGenerated',0);$count=$module.ProcCountLines('DeleteGenerated',0)
    $body=$module.Lines($start,$count)
    $needle="    ' Verification has its own authority boundary and performs no repair or save."
    if(-not $body.Contains($needle)){throw 'Tester verification boundary changed.'}
    $body=$body.Replace($needle,'    TestTesterFaults.AtPhase "VERIFY"'+"`r`n"+$needle)
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    $admin=$packages['invSys.Admin.xlam'].VBProject
    foreach($entry in @(@($core,'modDeploymentPaths','BrowseForFolderPath'),@($admin,'modTesterSetup','BrowseForSharePointRoot'))){
        $module=$entry[0].VBComponents.Item($entry[1]).CodeModule
        $start=$module.ProcStartLine($entry[2],0);$count=$module.ProcCountLines($entry[2],0)
        $body=$module.Lines($start,$count)
        $body=[regex]::Replace($body,'(?m)^(    On Error GoTo )',('    If TestTesterFaults.PickerPending() Then '+$entry[2]+' = TestTesterFaults.PickerResult(): Exit Function'+"`r`n"+'$1'))
        $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    }
    $module=$core.VBComponents.Item('modTesterProvisioning').CodeModule
    $start=$module.ProcStartLine('DetectSharePointRoot',0);$count=$module.ProcCountLines('DetectSharePointRoot',0)
    $body=$module.Lines($start,$count).Replace('    On Error GoTo FailDetect',"    If TestTesterFaults.PickerPending() Then Exit Function`r`n    On Error GoTo FailDetect")
    $module.DeleteLines($start,$count);$module.InsertLines($start,$body)
    $admin.VBComponents.Item('TestTesterOwner').CodeModule.AddFromString(@'
Public Function DirectSetup() As Boolean
    DirectSetup = frmSetupTesterStation.OwnerDirectSetupForTest()
End Function
Public Function DirectDelete() As Boolean
    DirectDelete = frmSetupTesterStation.OwnerDirectDeleteForTest()
End Function
Public Function Layout(ByVal enlarged As Boolean) As String
    Layout = frmSetupTesterStation.OwnerLayoutForTest(enlarged)
End Function
Public Sub Picker(ByVal kind As String)
    frmSetupTesterStation.OwnerPickerForTest kind
End Sub
Public Function SameInputs(ByVal root As String, ByVal publish As String) As Boolean
    SameInputs = frmSetupTesterStation.OwnerSameInputsForTest(root, publish)
End Function
'@)
    $admin.VBComponents.Item('frmSetupTesterStation').CodeModule.AddFromString(@'
Public Function OwnerDirectSetupForTest() As Boolean
    Dim spec As modTesterSetup.TesterSetupSpec, secret As String, notice As String, path As String
    If Not BuildSpecFromForm(spec, secret) Then Err.Raise 5
    OwnerDirectSetupForTest = modTesterCommands.Setup(mTesterContext, spec.UserId, spec.PinHash, spec.WarehouseId, _
        spec.StationId, spec.PathLocal, spec.PathSharePointRoot, notice, path)
    secret = "": spec.PinHash = ""
End Function
Public Function OwnerDirectDeleteForTest() As Boolean
    Dim spec As modTesterSetup.TesterSetupSpec, notice As String
    If Not BuildCleanupSpecFromForm(spec) Then Err.Raise 5
    OwnerDirectDeleteForTest = modTesterCommands.DeleteGenerated(mTesterContext, spec.WarehouseId, spec.StationId, _
        spec.PathLocal, spec.PathSharePointRoot, True, notice)
End Function
Public Function OwnerLayoutForTest(ByVal enlarged As Boolean) As String
    Dim control As Object, clipped As Long, staleHeading As Boolean
    If enlarged Then Me.Width = 760: Me.Height = 600 Else Me.Width = 600: Me.Height = 500
    For Each control In Me.Controls
        If control.Visible Then
            If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > Me.InsideWidth Or control.Top + control.Height > Me.InsideHeight Then clipped = clipped + 1
            If TypeName(control) = "Label" Or TypeName(control) = "Frame" Then
                If Trim$(CStr(control.Caption)) = "Create Warehouse" Then staleHeading = True
            End If
        End If
    Next control
    Me.Repaint
    OwnerLayoutForTest = CStr(clipped) & "|" & CStr(Me.btnOK.Enabled And Me.btnCancel.Enabled) & "|" & CStr(staleHeading)
End Function
Public Sub OwnerPickerForTest(ByVal kind As String)
    If kind = "HUB" Then mBtnHubPathHelper_Click Else mBtnSharePointHelper_Click
End Sub
Public Function OwnerSameInputsForTest(ByVal root As String, ByVal publish As String) As Boolean
    OwnerSameInputsForTest = (CStr(Me.txtPathLocal.Value) = root And CStr(Me.txtPathSharePoint.Value) = publish)
End Function
'@)
}

function Test-TesterOwnershipPaths($Fixture,$Other) {
    function T([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestTesterOwner.'+$Name) $Values}
    function F([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestTesterFaults.'+$Name) $Values}
    $hostBook=$excel.Workbooks.Add();$openBook=$null
    $root=Join-Path $runRoot 'tester-owned-reopen'
    $id='TEST'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
    $sourceHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    try {
        SelectTarget $Fixture
        $generation=[int](F 'Generation')
        Check 'Tester.Layout.DesignerHeadingMatchesPurpose' ($packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmSetupTesterStation').Designer.Caption -ceq 'Test Environment Setup')
        if(-not [bool](T 'OpenForm' @($hostBook.Name))){throw 'Tester form did not open.'}
        [void](T 'PrepareSetup' @($id,$root,''))
        [void](F 'StartLogCapture' @('fixture-user',$id,$root))
        $perf=Join-Path ([IO.Path]::GetTempPath()) 'invSys.Inventory.Sync.log'
        $before=if(Test-Path -LiteralPath $perf){(Get-Item -LiteralPath $perf).Length}else{0}
        $complete=[bool](T 'Setup')
        $facts=([string](F 'FinishLogCapture')).Split('|')
        Check 'Tester.Ownership.SetupCompleted' $complete
        if(-not $complete){throw 'Ownership tests require an actually created environment.'}
        Check 'Tester.Logging.PrintSinkExercised' ([int]$facts[1] -gt 0)
        Check 'Tester.Logging.EnteredValuesExcludedFromPrint' ($facts[0] -ceq 'False')
        $stream=[IO.File]::Open($perf,'Open','Read','ReadWrite');[void]$stream.Seek($before,'Begin');$reader=[IO.StreamReader]::new($stream)
        try{$newLog=$reader.ReadToEnd()}finally{$reader.Dispose()}
        Check 'Tester.Logging.EnteredValuesExcludedFromFile' (-not($newLog.Contains($id) -or $newLog.Contains($root) -or $newLog.Contains('fixture-user')))
        $receipt=Join-Path $root 'invSys.TesterOwnership.json'
        $record=Get-Content -LiteralPath $receipt -Raw|ConvertFrom-Json
        $receiptHash=(Get-FileHash -LiteralPath $receipt).Hash
        Check 'Tester.Receipt.CreatorSessionBound' ($record.CreatorUser -ceq 'config-admin' -and $record.CreatorGeneration -eq $generation -and $record.SourceWarehouse -ceq $Fixture.Warehouse)
        Check 'Tester.Receipt.ExactScopeBound' ($record.Root -ceq $root -and $record.Warehouse -ceq $id -and $record.Station -ceq 'S1')
        $unclaimed=@(Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object {$_.Name -cne 'invSys.TesterOwnership.json' -and $_.FullName -notin $record.Artifacts})
        Check 'Tester.Receipt.AllGeneratedFilesAllowlisted' ($unclaimed.Count -eq 0)
        # Only generated artifact names: strip the entered root and synthetic warehouse ID.
        foreach($item in $unclaimed){Write-Output ('Tester.UnclaimedArtifact='+$item.FullName.Substring($root.Length).Replace($id,'WAREHOUSE'))}
        foreach($enlarged in @($false,$true)){
            $layout=([string](T 'Layout' @($enlarged))).Split('|')
            $label=if($enlarged){'Enlarged'}else{'Default'}
            Check ('Tester.Layout.'+$label+'.WithinBounds') ($layout[0] -ceq '0')
            Check ('Tester.Layout.'+$label+'.CloseControlsEnabled') ($layout[1] -ceq 'True')
            Check ('Tester.Layout.'+$label+'.CorrectHeading') ($layout[2] -ceq 'False')
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Test Environment Setup' ('tester-success-'+$label.ToLowerInvariant()+'.png')}
        }
        if($AdminTesterSetupEvidenceOnly){return}
        [void](T 'CloseForm')
        SelectTarget $Fixture
        if(-not [bool](T 'OpenForm' @($hostBook.Name))){throw 'Reopened Tester form did not open.'}
        [void](T 'PrepareSetup' @($id,$root,''))
        Check 'Tester.Setup.ReopenedOwnedResume' ([bool](T 'Setup'))
        Check 'Tester.Receipt.CreatorProvenanceImmutable' ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
        [void](T 'CloseForm')
        $config=Join-Path $root ($id+'.invSys.Config.xlsb')
        # A different Admin is authorized by the receipt's source warehouse, not by possession.
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        try{
            $caps=$null
            foreach($sheet in $auth.Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -ceq 'tblCapabilities'){$caps=$table}}}
            if($null -eq $caps){throw 'Generated fixture capability table missing.'}
            $row=$caps.ListRows.Add()
            foreach($pair in @{UserId='config-producer';Capability='ADMIN_MAINT';WarehouseId=$Fixture.Warehouse;StationId='S1';Status='Active'}.GetEnumerator()){$row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value}
            $auth.Save()
        }finally{$auth.Close($false)}
        SelectTarget $Fixture 'config-producer'
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
        Check 'Tester.Receipt.DifferentSourceAdminMayResume' ([bool](T 'Setup'))
        Check 'Tester.Receipt.DifferentAdminCannotRewriteCreator' ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
        [void](T 'CloseForm')
        # Saved revocation must override a cached capability without changing the session.
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        try{
            foreach($sheet in $auth.Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -ceq 'tblCapabilities'){
                foreach($entry in $table.ListRows){if([string]$entry.Range.Cells.Item(1,$table.ListColumns.Item('UserId').Index).Value2 -ceq 'config-producer' -and [string]$entry.Range.Cells.Item(1,$table.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){$entry.Range.Cells.Item(1,$table.ListColumns.Item('Status').Index).Value2='Inactive'}}
            }}}
            $auth.Save()
        }finally{$auth.Close($false)}
        Check 'Tester.Authority.SavedRevocation.SetupDenied' (-not [bool](T 'DirectSetup'))
        Check 'Tester.Authority.SavedRevocation.DeleteDenied' (-not [bool](T 'DirectDelete'))
        Check 'Tester.Authority.SavedRevocation.ReceiptPreserved' ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
        [void](T 'CloseForm');SelectTarget $Fixture
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        try{
            $auth.Worksheets.Item(1).Range('Z99').Value2='Unsaved source fixture sentinel'
            Check 'Tester.Authority.DirtySourceAuth.SetupDenied' (-not [bool](T 'DirectSetup'))
            Check 'Tester.Authority.DirtySourceAuth.DeleteDenied' (-not [bool](T 'DirectDelete'))
            Check 'Tester.Authority.DirtySourceAuth.RemainsOpenUnsaved' (-not [bool]$auth.Saved)
        }finally{$auth.Close($false);[void](T 'CloseForm')}
        SelectTarget $Other
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
        Check 'Tester.Receipt.FreshOtherWarehouseAdminCannotAdopt' (-not [bool](T 'Setup'))
        Check 'Tester.Receipt.FreshOtherWarehouseAdminCannotDelete' (-not [bool](T 'DirectDelete'))
        [void](T 'CloseForm')
        $originalReceipt=[IO.File]::ReadAllBytes($receipt)
        $configHash=(Get-FileHash -LiteralPath $config).Hash
        foreach($fault in @('SourceWarehouse','Root','Artifacts','CreatorGeneration','Version')){
            $invalid=([Text.Encoding]::UTF8.GetString($originalReceipt)|ConvertFrom-Json)
            switch($fault){
                'SourceWarehouse'{$invalid.SourceWarehouse=$Other.Warehouse}
                'Root'{$invalid.Root=$Other.Root}
                'Artifacts'{$invalid.Artifacts[0]=Join-Path $runRoot 'foreign-canary.xlsb'}
                'CreatorGeneration'{$invalid.PSObject.Properties.Remove('CreatorGeneration')}
                'Version'{$invalid.Version=999}
            }
            [IO.File]::WriteAllText($receipt,($invalid|ConvertTo-Json -Depth 4))
            SelectTarget $Fixture
            [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
            Check ('Tester.Receipt.Invalid'+$fault+'.SetupRefused') (-not [bool](T 'Setup'))
            [void](T 'Prepare' @($id,$root,'Admin'));[void](T 'DeleteGenerated')
            Check ('Tester.Receipt.Invalid'+$fault+'.ConfigPreserved') ((Get-FileHash -LiteralPath $config).Hash -ceq $configHash)
            [void](T 'CloseForm');[IO.File]::WriteAllBytes($receipt,$originalReceipt)
        }
        foreach($dirty in @($false,$true)){
            SelectTarget $Fixture
            $configHash=(Get-FileHash -LiteralPath $config).Hash
            $openBook=$excel.Workbooks.Open($config,0,$false)
            if($dirty){$openBook.Worksheets.Item(1).Range('Z99').Value2='Unsaved owned test sentinel'}
            if(-not [bool](T 'OpenForm' @($hostBook.Name))){throw 'Borrowed-workbook form did not open.'}
            [void](T 'PrepareSetup' @($id,$root,''))
            $label=if($dirty){'Dirty'}else{'Clean'}
            Check ('Tester.Setup.Open'+$label+'.Refused') (-not [bool](T 'Setup'))
            [void](T 'Prepare' @($id,$root,'Admin'))
            [void](T 'DeleteGenerated')
            Check ('Tester.Open'+$label+'.WorkbookStillOpen') ($openBook.Name -ceq ($id+'.invSys.Config.xlsb'))
            Check ('Tester.Open'+$label+'.SavedStatePreserved') ([bool]$openBook.Saved -eq (-not $dirty))
            Check ('Tester.Open'+$label+'.ReceiptPreserved') ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
            [void](T 'CloseForm');$openBook.Close($false);$openBook=$null
            Check ('Tester.Delete.Open'+$label+'.FilePreserved') ((Get-FileHash -LiteralPath $config).Hash -ceq $configHash)
        }
        foreach($mode in @('SignedOut','Reader','OtherWarehouseAdmin','SameAdminNewSession')){
            SelectTarget $Fixture
            [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
            switch($mode){'SignedOut'{[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')};'Reader'{SelectTarget $Fixture 'config-reader'};'OtherWarehouseAdmin'{SelectTarget $Other};'SameAdminNewSession'{SelectTarget $Fixture}}
            Check ('Tester.DirectCore.'+$mode+'.SetupDenied') (-not [bool](T 'DirectSetup'))
            Check ('Tester.DirectCore.'+$mode+'.DeleteDenied') (-not [bool](T 'DirectDelete'))
            Check ('Tester.DirectCore.'+$mode+'.ReceiptPreserved') ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
            [void](T 'CloseForm')
        }
        SelectTarget $Fixture
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,''))
        $hostBook.Close($false);$hostBook=$excel.Workbooks.Add()
        Check 'Tester.SourceWorkbook.ClosedSourceCannotRebind' (-not [bool](T 'Setup'))
        Check 'Tester.SourceWorkbook.DirectCoreDeleteDenied' (-not [bool](T 'DirectDelete'))
        Check 'Tester.SourceWorkbook.ReceiptPreserved' ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
        [void](T 'CloseForm')
        foreach($picker in @('HubChanged','HubCancelled','ShareUnchanged')){
            SelectTarget $Fixture
            $publish=Join-Path $runRoot 'tester-picker-publication'
            [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$root,$publish))
            $result=switch($picker){'HubChanged'{Join-Path $runRoot 'changed-picker-root'};'HubCancelled'{''};'ShareUnchanged'{$publish}}
            [void](F 'ArmPicker' @($result))
            $kind=if($picker -ceq 'ShareUnchanged'){'SHARE'}else{'HUB'}
            [void](T 'Picker' @($kind))
            Check ('Tester.Picker.'+$picker+'.YieldExercised') ([bool](F 'Injected'))
            Check ('Tester.Picker.'+$picker+'.AuthorityRechecked') (([string](T 'Summary')).StartsWith('The invSys session, source warehouse or Admin rights changed.'))
            Check ('Tester.Picker.'+$picker+'.InputsPreserved') ([bool](T 'SameInputs' @($root,$publish)))
            [void](T 'CloseForm')
        }
        # Verification is a separate read-only authority boundary. Resume partial cleanup.
        SelectTarget $Fixture
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'Prepare' @($id,$root,'Admin'))
        [void](F 'Arm' @('VERIFY',''))
        [void](T 'DeleteGenerated')
        Check 'Tester.Cleanup.VerificationAuthorityChangeInjected' ([bool](F 'Injected'))
        Check 'Tester.Cleanup.VerificationRefusesSuccess' (-not ([string](T 'Summary')).StartsWith('Test Environment generated files deleted.'))
        Check 'Tester.Cleanup.PartialReceiptRetained' ((Get-FileHash -LiteralPath $receipt).Hash -ceq $receiptHash)
        [void](T 'CloseForm');[void](F 'Arm' @('',''))
        SelectTarget $Fixture
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'Prepare' @($id,$root,'Admin'))
        [void](T 'DeleteGenerated')
        Check 'Tester.Cleanup.ReopenedPartialCompletes' (([string](T 'Summary')).StartsWith('Test Environment generated files deleted.') -and -not(Test-Path -LiteralPath $receipt))
        Check 'Tester.Cleanup.AllGeneratedFilesRemoved' (@(Get-ChildItem -LiteralPath $root -Recurse -File).Count -eq 0)
        [void](T 'CloseForm')
        foreach($phase in @('RUNTIME','SEED','AUTH','WORKBOOK','CONFIG','PUBLISH')){
            SelectTarget $Fixture
            $phaseRoot=Join-Path $runRoot ('tester-phase-'+$phase)
            [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$phaseRoot,''))
            [void](F 'Arm' @($phase,''))
            Check ('Tester.Setup.Phase.'+$phase+'.Refused') (-not [bool](T 'Setup'))
            Check ('Tester.Setup.Phase.'+$phase+'.InjectionReached') ([bool](F 'Injected'))
            Check ('Tester.Setup.Phase.'+$phase+'.ReceiptRetained') (Test-Path -LiteralPath (Join-Path $phaseRoot 'invSys.TesterOwnership.json'))
            [void](T 'CloseForm');[void](F 'Arm' @('',''))
        }
        SelectTarget $Fixture
        $phaseRoot=Join-Path $runRoot 'tester-scope-change'
        $redirectRoot=Join-Path $runRoot 'tester-scope-foreign'
        [void](T 'OpenForm' @($hostBook.Name));[void](T 'PrepareSetup' @($id,$phaseRoot,''))
        [void](F 'Arm' @('SEED',$redirectRoot))
        Check 'Tester.Setup.ChangedWriteRoot.Refused' (-not [bool](T 'Setup'))
        Check 'Tester.Setup.ChangedWriteRoot.Injected' ([bool](F 'Injected'))
        Check 'Tester.Setup.ChangedWriteRoot.NoForeignWrites' (-not(Test-Path -LiteralPath $redirectRoot))
        [void](T 'CloseForm');[void](F 'Arm' @('',''))
        Check 'Tester.Ownership.SourceConfigPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $sourceHash)
    }finally{
        [void](F 'Arm' @('',''));[void](F 'FinishLogCapture');[void](T 'CloseForm')
        if($null -ne $openBook){$openBook.Close($false)}
        $hostBook.Close($false);SelectTarget $Fixture
    }
}
