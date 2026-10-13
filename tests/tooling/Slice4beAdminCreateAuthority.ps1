# D8-CREATE-01: preserve existing owner checks; use existing sign-out/target wiring.
. (Join-Path $PSScriptRoot 'Slice4beAdminCreateEdges.ps1')
function Install-CreateAuthorityProbe($Project,$Form,$Probe) {
    $exists=$Form.Lines(1,$Form.CountOfLines) -match '(?m)^Private Sub mBtnPublishExisting_Click\('
    $body=if($exists){'    Call mBtnPublishExisting_Click: OwnerPublishExistingForTest = (Me.Tag = "COMPLETE")'}else{'    OwnerPublishExistingForTest = False'}
    $Form.AddFromString("Public Function OwnerPublishExistingForTest() As Boolean`r`n$body`r`nEnd Function")
    $Form.AddFromString(@'
Public Function OwnerPublishControlForTest() As Boolean
    Dim item As Object
    For Each item In Me.Controls
        If item.Name = "btnPublishExisting" Then OwnerPublishControlForTest = item.Visible And item.Enabled
    Next item
End Function
Public Function OwnerLayoutForTest() As Boolean
    Dim a As Object, b As Object
    OwnerLayoutForTest = True
    For Each a In Me.Controls
        If a.Visible Then
            If a.Left < 0 Or a.Top < 0 Or a.Left + a.Width > Me.InsideWidth + 1 Or a.Top + a.Height > Me.InsideHeight + 1 Then OwnerLayoutForTest = False
            If TypeName(a) = "CommandButton" Then
                For Each b In Me.Controls
                    If b.Visible And TypeName(b) = "CommandButton" And a.Name <> b.Name Then
                        If a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And a.Top < b.Top + b.Height And b.Top < a.Top + a.Height Then OwnerLayoutForTest = False
                    End If
                Next b
            End If
        End If
    Next a
End Function
'@)
    $Probe.AddFromString(@'
Public Function PublishExisting() As Boolean
    PublishExisting = frmCreateWarehouse.OwnerPublishExistingForTest()
End Function
Public Function PublishControl() As Boolean
    PublishControl = frmCreateWarehouse.OwnerPublishControlForTest()
End Function
Public Function Layout() As Boolean
    Layout = frmCreateWarehouse.OwnerLayoutForTest()
End Function
Public Sub Expand()
    frmCreateWarehouse.Width = 800: frmCreateWarehouse.Height = 700
End Sub
'@)
    $core=$packages['invSys.Core.xlam'].VBProject
    $calls=$core.VBComponents.Add(1);$calls.Name='TestCreateCalls'
    $calls.CodeModule.AddFromString(@'
Option Explicit
Private mBootstrap As Long, mPublish As Long
Public Sub Reset()
    mBootstrap = 0: mPublish = 0
End Sub
Public Sub ObserveBootstrap()
    mBootstrap = mBootstrap + 1
End Sub
Public Sub ObservePublish()
    mPublish = mPublish + 1
End Sub
Public Function Bootstraps() As Long
    Bootstraps = mBootstrap
End Function
Public Function Publications() As Long
    Publications = mPublish
End Function
'@)
    $bootstrap=$core.VBComponents.Item('modWarehouseBootstrap').CodeModule
    foreach($entry in @(@('BootstrapWarehouseLocal','ObserveBootstrap'),@('PublishInitialArtifacts','ObservePublish'))){
        $declaration=@(for($i=1;$i -le $bootstrap.CountOfLines;$i++){if($bootstrap.Lines($i,1) -match ('^Public Function '+$entry[0]+'\(')){$i}})
        if($declaration.Count -ne 1){throw 'Actual Core bootstrap/publication boundary unavailable.'}
        $bootstrap.InsertLines($declaration[0]+1,('    TestCreateCalls.'+$entry[1]))
    }
    $observer=$core.VBComponents.Add(1);$observer.Name='TestCreateLogs'
    $observer.CodeModule.AddFromString(@'
Option Explicit
Private mRoot As String, mId As String, mLeaks As Long, mSeen As Long
Public Sub Arm(ByVal root As String, ByVal warehouseId As String)
    mRoot = root: mId = warehouseId: mLeaks = 0: mSeen = 0
End Sub
Public Sub Observe(ByVal text As String)
    If mRoot = "" Then Exit Sub
    mSeen = mSeen + 1
    If InStr(1, text, mRoot, vbTextCompare) > 0 Or InStr(1, text, mId, vbBinaryCompare) > 0 Or InStr(1, text, "Generated Create coverage", vbBinaryCompare) > 0 Then mLeaks = mLeaks + 1
End Sub
Public Function Leaks() As Long
    Leaks = mLeaks
End Function
Public Function Seen() As Long
    Seen = mSeen
End Function
Public Sub Disarm()
    mRoot = "": mId = ""
End Sub
'@)
    $log=$core.VBComponents.Item('modPerfLog').CodeModule
    $line=@(for($i=1;$i -le $log.CountOfLines;$i++){if($log.Lines($i,1) -match '^    Print #fileNum, lineText$'){$i}})
    if($line.Count -ne 1){throw 'Actual diagnostic file-write boundary unavailable.'}
    foreach($i in ($line|Sort-Object -Descending)){$log.InsertLines($i,'    TestCreateLogs.Observe lineText')}
    Install-CreateEdgeProbe $core $calls.CodeModule
}

function Test-CreateAuthority($Fixture,$Other,$HostBook) {
    if($CaptureEvidence){$excel.Visible=$true}
    function Invoke-Owner([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestCreateOwner.'+$Name) $Values}
    function Open-Owner($Case){
        SelectTarget $Fixture
        if(-not [bool](Invoke-Owner 'OpenForm' @($HostBook.Name))){throw 'Authorized Create callback unavailable; not authorization RED.'}
        [void](Invoke-Owner 'Prepare' @($Case.Id,$Case.Root,$Case.Share,'Training'))
    }
    function Case([string]$Label){
        $id='CA'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
        $root=Join-Path $runRoot ('authority-'+$Label)
        [pscustomobject]@{Id=$id;Root=$root;Share='';Config=(Join-Path $root ($id+'.invSys.Config.xlsb'))}
    }
    function Pins($Case){
        $pins=@{}
        foreach($suffix in @('.invSys.Config.xlsb','.invSys.Auth.xlsb','.invSys.Data.Inventory.xlsb')){
            $path=Join-Path $Case.Root ($Case.Id+$suffix)
            $pins[$path]=(Get-FileHash -LiteralPath $path).Hash
        }
        return $pins
    }
    function Same($Pins){foreach($path in $Pins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $Pins[$path]){return $false}};return $true}
    function State([string]$Mode){
        switch($Mode){
            'SignedOut' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            'Reader' {SelectTarget $Fixture 'config-reader'}
            'OtherWarehouseAdmin' {SelectTarget $Other}
            'NewSession' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut');SelectTarget $Fixture}
        }
        $signed=[bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn')
        $actor=[string](Run 'invSys.Core.xlam' 'modAuth.GetCurrentUserId')
        $warehouse=[string](Run 'invSys.Core.xlam' 'modNasConnection.GetCurrentTargetWarehouseId')
        $cap=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT',$actor,$warehouse,'S1'))
        $valid=switch($Mode){
            'SignedOut' {-not $signed -and -not $cap}
            'Reader' {$signed -and -not $cap -and $warehouse -ceq $Fixture.Warehouse}
            'OtherWarehouseAdmin' {$signed -and $cap -and $warehouse -ceq $Other.Warehouse}
            'NewSession' {$signed -and $cap -and $warehouse -ceq $Fixture.Warehouse}
        }
        Check ('Create.Authority.'+$Mode+'.Precondition') $valid
        if(-not $valid){throw 'Authorization state was not established.'}
    }
    foreach($mode in @('Reader','OtherWarehouseAdmin','NewSession')){
        $case=Case ('create-'+$mode);Open-Owner $case;State $mode
        [void](Invoke-Owner 'Create')
        Check ('Create.Authority.'+$mode+'.CreateDenied') (-not (Test-Path -LiteralPath $case.Config))
        Check ('Create.Authority.'+$mode+'.NoCompletionClaim') (-not ([string](Invoke-Owner 'Facts')).StartsWith('True|True|'))
        [void](Invoke-Owner 'CloseForm')
    }
    foreach($mode in @('SignedOut','Reader','OtherWarehouseAdmin','NewSession')){
        $case=Case ('retry-'+$mode)
        $case.Share=Join-Path $runRoot ('obstruct-'+$mode)
        New-Item -ItemType Directory -Path $case.Share|Out-Null
        [IO.File]::WriteAllText((Join-Path $case.Share $case.Id),'Generated publish obstruction')
        Open-Owner $case
        [void](Run 'invSys.Core.xlam' 'TestCreateLogs.Arm' @($runRoot,$case.Id))
        [void](Invoke-Owner 'Create')
        $facts=[string](Invoke-Owner 'Facts')
        Check ('Create.Authority.'+$mode+'.RealPartialSuccess') ($facts -ceq 'True|False|Retry Publish|True|1|1')
        if($facts -cne 'True|False|Retry Publish|True|1|1'){
            if($script:CreateCoreAvailable){Write-Output ('Existing-runtime reader last source line: '+[long](Run 'invSys.Core.xlam' 'TestCreateCalls.LastReadStep'))}
            if($script:CreateCoreAvailable){Write-Output ([string](Run 'invSys.Core.xlam' 'TestCreateCalls.ContextFacts'))}
            throw 'Partial completion fixture did not reach actual publication failure.'
        }
        Check ('Create.Redaction.'+$mode+'.DiagnosticOutputObserved') ([long](Run 'invSys.Core.xlam' 'TestCreateLogs.Seen') -gt 0)
        Check ('Create.Redaction.'+$mode+'.NoEnteredValuesInDiagnosticOutput') ([long](Run 'invSys.Core.xlam' 'TestCreateLogs.Leaks') -eq 0)
        [void](Run 'invSys.Core.xlam' 'TestCreateLogs.Disarm')
        $before=Pins $case
        $case.Share=Join-Path $runRoot ('retry-authorized-destination-'+$mode)
        New-Item -ItemType Directory -Path $case.Share|Out-Null
        [void](Invoke-Owner 'SetShare' @($case.Share))
        State $mode
        [void](Invoke-Owner 'Create')
        Check ('Create.Authority.'+$mode+'.RetryDenied') (-not (Test-Path -LiteralPath (Join-Path $case.Share ($case.Id+'.config.json'))))
        Check ('Create.Authority.'+$mode+'.RetryNoCompletionClaim') (-not ([string](Invoke-Owner 'Facts')).StartsWith('True|True|'))
        Check ('Create.Authority.'+$mode+'.EarlierLocalArtifactsPreserved') (Same $before)
        if($CaptureEvidence -and $mode -ceq 'SignedOut'){
            Initialize-SettingsCapture
            CaptureOwnedFormByCaptionEvidence 'Create Warehouse' 'create-denied-retry.png'
        }
        [void](Invoke-Owner 'CloseForm')
        if($mode -ceq 'SignedOut'){$resumeCase=$case;$resumePins=$before}
    }
    # A NEW form and live session must recover publication from saved identity.
    $resumeCase.Share=Join-Path $runRoot 'reopened-publication'
    New-Item -ItemType Directory -Path $resumeCase.Share|Out-Null
    Open-Owner $resumeCase
    $present=[bool](Invoke-Owner 'PublishControl')
    Check 'Create.Reopened.ExplicitPublishExistingControl' $present
    # Deliberately unrelated new-warehouse fields must not change saved identity.
    [void](Invoke-Owner 'Prepare' @('UNRELATED_ENTRY',$resumeCase.Root,$resumeCase.Share,'Operational'))
    Check 'Create.Reopened.PublishExistingCompletes' ([bool](Invoke-Owner 'PublishExisting'))
    Check 'Create.Reopened.PublishesSavedIdentity' (Test-Path -LiteralPath (Join-Path $resumeCase.Share ($resumeCase.Id+'.config.json')))
    Check 'Create.Reopened.NoNewEnteredIdentity' (-not (Test-Path -LiteralPath (Join-Path $resumeCase.Share 'UNRELATED_ENTRY.config.json')))
    Check 'Create.Reopened.LocalAuthorityBytesPreserved' (Same $resumePins)
    Check 'Create.Reopened.ZeroBootstrapCalls' (([string](Invoke-Owner 'Facts')).Split('|')[4] -ceq '0')
    Check 'Create.Reopened.MinimumLayout' ([bool](Invoke-Owner 'Layout'))
    if($CaptureEvidence){
        Initialize-SettingsCapture
        CaptureOwnedFormByCaptionEvidence 'Create Warehouse' 'create-reopened-published.png'
    }
    [void](Invoke-Owner 'Expand')
    Check 'Create.Reopened.ExpandedLayout' ([bool](Invoke-Owner 'Layout'))
    if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Create Warehouse' 'create-reopened-expanded.png'}
    [void](Invoke-Owner 'CloseForm')
    Test-CreateEdges $Fixture $Other $HostBook $resumeCase
}
