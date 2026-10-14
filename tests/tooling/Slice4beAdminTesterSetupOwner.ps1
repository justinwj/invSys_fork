# Real packaged Ribbon and Delete Generated handler; only owned disposable canaries.
# Adapters supply confirmation, expose fixed facts and suppress raw diagnostic payloads.
function Install-AdminTesterSetupOwnerProbe {
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
Public Sub CloseForm()
    Dim item As Object
    For Each item In VBA.UserForms
        If TypeName(item) = "frmSetupTesterStation" Then item.OwnerCloseForTest: Exit Sub
    Next item
End Sub
Public Function Facts() As String
    Facts = frmSetupTesterStation.OwnerFactsForTest
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
    if([regex]::Matches($source,'frmSetupTesterStation.Show\b').Count -ne 1){throw 'Tester callback presentation boundary changed.'}
    $source=$source.Replace('frmSetupTesterStation.Show','frmSetupTesterStation.Show vbModeless')
    $entry.DeleteLines($start,$count);$entry.InsertLines($start,$source)
    $project.VBComponents.Item('modTesterSetup').CodeModule.AddFromString(@'
Private Sub LogDiagnosticEvent(ByVal category As String, ByVal detail As String)
    ' Test-only sink: never persist raw entered paths or report payloads.
End Sub
'@)
}

function Test-AdminTesterSetupOwner($Fixture,$Other) {
    function Invoke-Tester([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestTesterOwner.'+$Name) $Values}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $hostBook=$excel.Workbooks.Add()
    $sourceHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherHash=(Get-FileHash -LiteralPath $Other.Config).Hash
    try {
        foreach($mode in @('Admin','Cancel','SignedOut','Reader','OtherWarehouseAdmin','ConfirmationSignOut')){
            SelectTarget $Fixture
            if(-not [bool](Invoke-Tester 'OpenForm' @($hostBook.Name))){throw 'Authorized real Tester callback did not open; not behavioral RED.'}
            $root=[IO.Path]::GetFullPath((Join-Path $runRoot ('tester-cleanup-'+$mode)))
            if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Cleanup target escaped owned fixture root.'}
            New-Item -ItemType Directory -Path $root|Out-Null
            $id='TEST'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
            $canary=Join-Path $root ($id+'.Receiving.Operator.xlsm')
            $unrelated=Join-Path $root 'unrelated-canary.txt'
            [IO.File]::WriteAllText($canary,'Disposable deletion diagnostic fixture')
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
