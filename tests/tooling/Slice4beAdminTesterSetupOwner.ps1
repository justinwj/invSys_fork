# Real packaged Ribbon and Delete Generated handler; only owned disposable canaries.
# Adapters supply confirmation, expose fixed facts and suppress raw diagnostic payloads.
function Install-AdminTesterSetupOwnerProbe {
    $commands=$null
    try{$commands=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modTesterCommands').CodeModule}catch{}
    if($null -ne $commands){
        $commands.InsertLines($commands.CountOfDeclarationLines+1,'Private mOwnerGuardSiteForTest As String')
        $start=$commands.ProcStartLine('CheckPhase',0);$count=$commands.ProcCountLines('CheckPhase',0)
        $body=$commands.Lines($start,$count)
        foreach($entry in @(@('If Not modTesterAuthority.Check(mContext)','SOURCE'),@('If Not modTesterOwnership.Unchanged','RECEIPT'),@('If Not FilesAvailable','OPEN_FILES'))){
            $body=$body.Replace('    '+$entry[0],('    mOwnerGuardSiteForTest = "'+$entry[1]+'"'+"`r`n"+'    '+$entry[0]))
        }
        $commands.DeleteLines($start,$count);$commands.InsertLines($start,$body)
        $commands.AddFromString("Public Function OwnerGuardSiteForTest() As String`r`nOwnerGuardSiteForTest = mOwnerGuardSiteForTest`r`nEnd Function")
    }
    $project=$packages['invSys.Admin.xlam'].VBProject
    $control=$project.VBComponents.Add(2);$control.Name='TestTesterRibbon'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminSetupTesterStation"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestTesterOwner'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function OpenForm(ByVal hostName As String) As Boolean
    Dim control As New TestTesterRibbon, item As Object
    Application.Workbooks(hostName).Activate
    RibbonOnActionAdmin control
    DoEvents
    For Each item In VBA.UserForms
        If TypeName(item) = "frmSetupTesterStation" Then OpenForm = item.Visible
    Next item
End Function
Public Sub Prepare(ByVal warehouse As String, ByVal root As String, ByVal mode As String)
    frmSetupTesterStation.OwnerPrepareForTest warehouse, root, mode
End Sub
Public Sub DeleteGenerated()
    frmSetupTesterStation.OwnerDeleteForTest
End Sub
Public Sub PrepareSetup(ByVal warehouse As String, ByVal root As String, ByVal share As String)
    frmSetupTesterStation.OwnerPrepareSetupForTest warehouse, root, share
End Sub
Public Function Setup() As Boolean
    Setup = frmSetupTesterStation.OwnerSetupForTest()
End Function
Public Sub CloseForm()
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmSetupTesterStation" Then item.OwnerCloseForTest: Exit Sub
    Next item
End Sub
Public Function Facts() As String
    Facts = frmSetupTesterStation.OwnerFactsForTest
End Function
Public Function Summary() As String
    Summary = frmSetupTesterStation.lblSummary.Caption
End Function
'@)
    $form=$project.VBComponents.Item('frmSetupTesterStation').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,'Private mTestMode As String, mTestConfirmations As Long')
    $form.AddFromString(@'
Public Sub OwnerPrepareForTest(ByVal warehouse As String, ByVal root As String, ByVal mode As String)
    mFormBusy = True
    Me.txtWarehouseId.Value = warehouse
    Me.txtStationId.Value = "S1"
    Me.txtPathLocal.Value = root
    Me.txtPathSharePoint.Value = ""
    Me.txtAdminUser.Value = "fixture-user"
    mTestMode = mode: mTestConfirmations = 0
    mFormBusy = False
End Sub
Public Sub OwnerDeleteForTest()
    mBtnDeleteGenerated_Click
End Sub
Public Sub OwnerPrepareSetupForTest(ByVal warehouse As String, ByVal root As String, ByVal share As String)
    Dim secret As String
    OwnerPrepareForTest warehouse, root, "Setup"
    Randomize
    secret = CStr(Timer) & CStr(Rnd())
    Me.txtWarehouseName.Value = secret
    mTxtConfirmPin.Value = secret
    Me.txtPathSharePoint.Value = share
    secret = ""
End Sub
Public Function OwnerSetupForTest() As Boolean
    btnOK_Click
    OwnerSetupForTest = mSetupSucceeded
End Function
Public Sub OwnerCloseForTest()
    btnCancel_Click
End Sub
Public Function OwnerFactsForTest() As String
    OwnerFactsForTest = CStr(mTestConfirmations) & "|" & CStr(Left$(Me.lblSummary.Caption, 41) = "Test Environment generated files deleted.")
End Function
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    If Title <> "Delete Test Environment Files" Then Err.Raise 5, , "Unexpected diagnostic dialog."
    mTestConfirmations = mTestConfirmations + 1
    If mTestMode = "ConfirmationSignOut" Then modAuth.SignOut
    If mTestMode = "Cancel" Then MsgBox = vbNo Else MsgBox = vbYes
End Function
'@)
    $entry=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$entry.ProcStartLine('Admin_SetupTesterStation_Click',0);$count=$entry.ProcCountLines('Admin_SetupTesterStation_Click',0)
    $source=$entry.Lines($start,$count)
    if($source.Contains('modTesterPresentation.ShowTesterSetup')){
        $entry=$project.VBComponents.Item('modTesterPresentation').CodeModule
        $start=$entry.ProcStartLine('ShowTesterSetup',0);$count=$entry.ProcCountLines('ShowTesterSetup',0)
        $source=$entry.Lines($start,$count)
    }
    if([regex]::Matches($source,'frmSetupTesterStation.Show\b').Count -ne 1){throw 'Tester callback presentation boundary changed.'}
    $source=$source.Replace('frmSetupTesterStation.Show','frmSetupTesterStation.Show vbModeless')
    $entry.DeleteLines($start,$count);$entry.InsertLines($start,$source)
    $project.VBComponents.Item('modTesterSetup').CodeModule.AddFromString(@'
Private Sub LogDiagnosticEvent(ByVal category As String, ByVal detail As String)
    ' Test-only sink: never persist raw entered paths or report payloads.
End Sub
'@)
}

function Test-AdminTesterSetupPaths($Fixture,$Other) {
    function Invoke-TesterSetup([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestTesterOwner.'+$Name) $Values}
    $hostBook=$excel.Workbooks.Add()
    $sourceHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherHash=(Get-FileHash -LiteralPath $Other.Config).Hash
    try {
        foreach($mode in @('Admin','SignedOut','Reader','OtherWarehouseAdmin','Unmarked','Published','Nested')){
            SelectTarget $Fixture
            $context=[string](Run 'invSys.Core.xlam' 'modAuthUserCommands.CaptureContext')
            if(-not [bool](Invoke-TesterSetup 'OpenForm' @($hostBook.Name))){throw 'Real Tester callback did not open; no valid Setup test.'}
            $id='TEST'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
            $root=Join-Path $runRoot ('tester-setup-'+$mode)
            $share=Join-Path $runRoot ('tester-publish-'+$mode)
            New-Item -ItemType Directory -Path $share|Out-Null
            $foreign=$null
            if($mode -ceq 'Unmarked'){
                New-Item -ItemType Directory -Path $root|Out-Null
                $foreign=Join-Path $root 'foreign.txt'
            }
            if($mode -ceq 'Published'){$foreign=Join-Path $share 'foreign.txt'}
            if($foreign){[IO.File]::WriteAllText($foreign,'Must not adopt this destination');$foreignHash=(Get-FileHash -LiteralPath $foreign).Hash}
            if($mode -ceq 'Nested'){$share=Join-Path $root 'nested-publication'}
            [void](Invoke-TesterSetup 'PrepareSetup' @($id,$root,$share))
            switch($mode){
                'SignedOut'{[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                'Reader'{SelectTarget $Fixture 'config-reader'}
                'OtherWarehouseAdmin'{SelectTarget $Other}
            }
            $completed=[bool](Invoke-TesterSetup 'Setup')
            $summary=[string](Invoke-TesterSetup 'Summary')
            Check ('Tester.Setup.'+$mode+'.FixedFeedback') (-not($summary.Contains($root) -or $summary.Contains($id) -or $summary.Contains('fixture-user')))
            if($mode -ceq 'Admin'){
                Check 'Tester.Setup.Admin.Completed' $completed
                Check 'Tester.Setup.Admin.ConfigCreated' (Test-Path -LiteralPath (Join-Path $root ($id+'.invSys.Config.xlsb')))
                Check 'Tester.Setup.Admin.ReceivingCreated' (Test-Path -LiteralPath (Join-Path $root ($id+'.Receiving.Operator.xlsm')))
                Check 'Tester.Setup.Admin.SourceSessionPreserved' ([string](Run 'invSys.Core.xlam' 'modAuthUserCommands.CaptureContext') -ceq $context)
                # Approved receipt boundary; a bare runtime cannot prove disposable ownership.
                Check 'Tester.Setup.Admin.OwnershipReceiptCreated' (Test-Path -LiteralPath (Join-Path $root 'invSys.TesterOwnership.json'))
            }else{
                Check ('Tester.Setup.'+$mode+'.NoCompletion') (-not $completed)
                $expectedLocal=if($mode -ceq 'Unmarked'){1}else{0}
                $expectedShare=if($mode -ceq 'Published'){1}else{0}
                $localFiles=if(Test-Path -LiteralPath $root){@(Get-ChildItem -LiteralPath $root -Recurse -File -Force).Count}else{0}
                $shareFiles=if(Test-Path -LiteralPath $share){@(Get-ChildItem -LiteralPath $share -Recurse -File -Force).Count}else{0}
                Check ('Tester.Setup.'+$mode+'.NoDestinationWrites') ($localFiles -eq $expectedLocal)
                Check ('Tester.Setup.'+$mode+'.NoPublicationWrites') ($shareFiles -eq $expectedShare)
                if($foreign){Check ('Tester.Setup.'+$mode+'.ForeignPreserved') ((Get-FileHash -LiteralPath $foreign).Hash -ceq $foreignHash)}
            }
            [void](Invoke-TesterSetup 'CloseForm')
        }
        Check 'Tester.Setup.SourceConfigPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $sourceHash)
        Check 'Tester.Setup.OtherConfigPreserved' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherHash)
    } finally {
        [void](Invoke-TesterSetup 'CloseForm')
        $hostBook.Close($false)
        SelectTarget $Fixture
    }
}

function Test-AdminTesterSetupOwner($Fixture,$Other) {
    function Invoke-Tester([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestTesterOwner.'+$Name) $Values}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $hostBook=$excel.Workbooks.Add()
    $sourceHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherHash=(Get-FileHash -LiteralPath $Other.Config).Hash
    try {
        foreach($mode in @('Admin','Cancel','SignedOut','Reader','OtherWarehouseAdmin','ConfirmationSignOut','Unmarked')){
            SelectTarget $Fixture
            if(-not [bool](Invoke-Tester 'OpenForm' @($hostBook.Name))){throw 'Authorized real Tester callback did not open; not behavioral RED.'}
            $root=[IO.Path]::GetFullPath((Join-Path $runRoot ('tester-cleanup-'+$mode)))
            if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Cleanup target escaped owned fixture root.'}
            New-Item -ItemType Directory -Path $root|Out-Null
            $id='TEST'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
            $canary=Join-Path $root ($id+'.Receiving.Operator.xlsm')
            $unrelated=Join-Path $root 'unrelated-canary.txt'
            if($mode -ceq 'Unmarked'){
                [IO.File]::WriteAllText($canary,'Disposable deletion diagnostic fixture')
            }else{
                [void](Invoke-Tester 'PrepareSetup' @($id,$root,''))
                if(-not [bool](Invoke-Tester 'Setup')){
                    $notice=[string](Invoke-Tester 'Summary')
                    $known=@('The invSys session, source warehouse or Admin rights changed. Reopen Test Environment Setup.',
                        'The destination is unavailable, conflicting or not verifiably owned. Setup did not complete.',
                        'Setup could not finish. Ownership is retained for retry. No rollback is claimed.',
                        'Authority or destination scope changed. Work stopped; ownership is retained for retry.',
                        'Setup could not finish. No rollback is claimed.')
                    [pscustomobject]@{SetupRefusalCode=[array]::IndexOf($known,$notice);ReceiptCreated=(Test-Path -LiteralPath (Join-Path $root 'invSys.TesterOwnership.json'));Files=@(Get-ChildItem -LiteralPath $root -Recurse -File).Count}|ConvertTo-Json -Compress
                    $site=[string](Run 'invSys.Core.xlam' 'modTesterCommands.OwnerGuardSiteForTest')
                    if($site -cin @('SOURCE','RECEIPT','OPEN_FILES')){Write-Output ('Tester guard refusal site: '+$site)}
                    throw 'Authorized setup failed; cannot test owned cleanup.'
                }
                [void](Invoke-Tester 'CloseForm')
                SelectTarget $Fixture
                if(-not [bool](Invoke-Tester 'OpenForm' @($hostBook.Name))){throw 'Reopened Tester callback did not open.'}
            }
            [IO.File]::WriteAllText($unrelated,'Must be preserved')
            $keepHash=(Get-FileHash -LiteralPath $unrelated).Hash
            [void](Invoke-Tester 'Prepare' @($id,$root,$mode))
            switch($mode){
                'SignedOut'{[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                'Reader'{SelectTarget $Fixture 'config-reader'}
                'OtherWarehouseAdmin'{SelectTarget $Other}
            }
            [void](Invoke-Tester 'DeleteGenerated')
            $facts=([string](Invoke-Tester 'Facts')).Split('|')
            Check ('Tester.Delete.'+$mode+'.ConfirmationOnce') ($facts[0] -ceq '1')
            if($mode -cin @('SignedOut','ConfirmationSignOut')){Check ('Tester.Delete.'+$mode+'.SignedOutVerified') (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))}
            if($mode -ceq 'Reader'){Check 'Tester.Delete.Reader.DeniedByCore' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-reader',$Fixture.Warehouse,'S1')))}
            if($mode -ceq 'OtherWarehouseAdmin'){Check 'Tester.Delete.OtherWarehouseAdmin.TargetChanged' ([string](Run 'invSys.Core.xlam' 'modNasConnection.GetCurrentTargetWarehouseId') -ceq $Other.Warehouse)}
            if($mode -ceq 'Admin'){
                Check 'Tester.Delete.Admin.ActualFileRemoved' (-not (Test-Path -LiteralPath $canary))
                Check 'Tester.Delete.Admin.ReceiptRemoved' (-not (Test-Path -LiteralPath (Join-Path $root 'invSys.TesterOwnership.json')))
            }else{
                Check ('Tester.Delete.'+$mode+'.FilePreserved') (Test-Path -LiteralPath $canary)
            }
            Check ('Tester.Delete.'+$mode+'.UnrelatedFilePreserved') ((Get-FileHash -LiteralPath $unrelated).Hash -ceq $keepHash)
            [void](Invoke-Tester 'CloseForm')
        }
        Check 'Tester.Delete.SourceConfigPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $sourceHash)
        Check 'Tester.Delete.OtherConfigPreserved' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherHash)
    } finally {
        [void](Invoke-Tester 'CloseForm')
        $hostBook.Close($false)
        SelectTarget $Fixture
    }
}
