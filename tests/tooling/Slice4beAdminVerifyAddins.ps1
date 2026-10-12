# D18 actual public callback; generated non-executable files never replace XLAMs.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminVerifyBoundaries.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminVerifyGuide.ps1')
function Install-AdminVerifyAddinsProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestAdminVerify'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mCalls As Long, mResult As String, mNotices As Long
Private mExpectedRoot As String, mRootMatches As Boolean
Public Sub ObserveOwnerRoot(ByVal root As String)
    mRootMatches = (StrComp(root, mExpectedRoot & "\Addins\", vbTextCompare) = 0)
End Sub
Public Sub OwnerEntered()
    mCalls = mCalls + 1
End Sub
Public Function ObserveMessage(ByVal text As String) As VbMsgBoxResult
    If Left$(text, 21) = "Tracking unavailable:" Then
        mNotices = mNotices + 1
    ElseIf Left$(text, 37) = "All required add-ins are published." & vbCrLf Then
        mResult = "VERIFIED"
    ElseIf InStr(1, text, "Add-ins publish verification failed.", vbBinaryCompare) = 1 Then
        mResult = "FAILED"
    Else
        mResult = "OTHER"
    End If
    ObserveMessage = vbOK
End Function
Public Function Invoke(ByVal root As String, ByVal hostName As String, ByVal otherName As String) As String
    Dim target As Workbook, bound As Boolean, otherBefore As Long
    mCalls = 0: mResult = "": mNotices = 0
    mExpectedRoot = root: mRootMatches = False
    modAddinsPublish.SetAddinsPublishSharePointRootOverride root
    Application.Workbooks(hostName).Activate
    Set target = modAdmin.ResolveInteractiveAdminWorkbook()
    bound = (target Is Application.Workbooks(hostName))
    If Not bound Then Err.Raise 5, , "Generated callback host activation prerequisite unavailable"
    otherBefore = Application.Workbooks(otherName).Worksheets.Count
    modAdmin.Verify_AddinsPublished
    modAddinsPublish.ClearAddinsPublishSharePointRootOverride
    Invoke = CStr(mCalls) & "|" & mResult & "|" & CStr(mNotices) & "|" & CStr(bound) & "|" & CStr(otherBefore) & "|" & CStr(Application.Workbooks(otherName).Worksheets.Count) & "|" & CStr(mRootMatches)
End Function
Public Function PrepareHost(ByVal name As String) As Boolean
    Dim wb As Workbook, lo As ListObject, column As ListColumn, report As String
    Set wb = Application.Workbooks(name)
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wb, report) Then Exit Function
    Set lo = wb.Worksheets("UserCredentials").ListObjects("UserCredentials")
    Set column = lo.ListColumns.Add: column.Name = "Operator_Custom_Canary"
    If lo.ListRows.Count = 0 Then lo.ListRows.Add
    column.DataBodyRange.Cells(1, 1).Value2 = "Preserve this custom value"
    PrepareHost = True
End Function
Public Function HostPreserved(ByVal name As String) As Boolean
    Dim lo As ListObject
    Set lo = Application.Workbooks(name).Worksheets("UserCredentials").ListObjects("UserCredentials")
    HostPreserved = (lo.ListColumns("Operator_Custom_Canary").DataBodyRange.Cells(1, 1).Value2 = "Preserve this custom value")
End Function
'@)
    $owner=$admin.VBComponents.Item('modAddinsPublish').CodeModule
    $owner.InsertLines($owner.ProcBodyLine('VerifyAddinsPublished',0)+1,'    TestAdminVerify.OwnerEntered')
    $start=$owner.ProcStartLine('VerifyAddinsPublished',0);$end=$start+$owner.ProcCountLines('VerifyAddinsPublished',0)
    $rootLines=@(for($line=$start;$line -lt $end;$line++){if($owner.Lines($line,1).Trim() -ceq 'If addinsRoot = "" Then'){$line}})
    if($rootLines.Count -ne 1){throw 'Owner root observation boundary changed; not product RED.'}
    $owner.InsertLines($rootLines[0],'    TestAdminVerify.ObserveOwnerRoot addinsRoot')
    $presentation=@($admin.VBComponents|Where-Object Name -CEQ 'modAdminVerifyAddins')
    $messages=if($presentation.Count){$presentation[0].CodeModule}else{$admin.VBComponents.Item('modAdmin').CodeModule}
    $messages.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestAdminVerify.ObserveMessage(Prompt)
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1);$core.Name='TestVerifyActivity'
    $core.CodeModule.AddFromString(@'
Option Explicit
Public Function SetPolicy(ByVal mode As String) As Boolean
    Dim context As String, version As Long, request As String, notice As String
    Dim model As Object, row As Variant, user As Object, users As Collection
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "Capture")
    For Each row In model("Controls")
        If mode = "ControlOff" And row("ControlId") = "ADMIN_VERIFY_ADDINS" Then row("Collect") = False
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    request = modTrainingJson.EncodeObject(model)
    SetPolicy = modTrackingPolicySettings.SavePolicy(context, version, request, notice)
End Function
Public Function Definition(ByVal version As Long) As String
    Dim record As Object
    Set record = modActivityCatalog.Control("ADMIN_VERIFY_ADDINS", version)
    If Not record Is Nothing Then Definition = modTrainingJson.EncodeObject(record)
End Function
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(34)
        Set before = modActivityCatalog.Control(CStr(id), 34)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
    Install-AdminVerifyBoundaries $core.CodeModule
    if($AdminVerifyGuidePaths){Install-AdminVerifyGuideProbe}
}

function Test-AdminVerifyAddins($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestVerifyActivity.'+$Method) $Values}
    $catalogVersion=[int](Probe 'CurrentCatalogVersion')
    function FileHash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Pins([string]$Root){$pins=@{};Get-ChildItem -LiteralPath $Root -Recurse -File|ForEach-Object {$pins[$_.FullName]=FileHash $_.FullName};return $pins}
    function SamePins($Before,$After){if($Before.Count -ne $After.Count){return $false};foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}};return $true}
    SelectTarget $Fixture
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add();$otherSheetCount=$otherBook.Worksheets.Count
    $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Other workbook canary'
    if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.PrepareHost' @($hostBook.Name))){throw 'Generated operator host unavailable; not product RED.'}
    $hostReady=[bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.HostPreserved' @($hostBook.Name))
    if(-not $hostReady){throw 'Generated custom-column canary unavailable before callback; not product RED.'}
    Check 'AdminVerify.Fixture.CustomColumnPresentBeforeCallback' $hostReady
    Write-Output ('Verify fixture other sheet count before/after host preparation: '+$otherSheetCount+'/'+$otherBook.Worksheets.Count)
    $otherPins=Pins $Other.Root
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=FileHash $inventory
    $id='ADMIN_VERIFY_ADDINS'
    try{
        if($AdminVerifyExpectationChoicesOnly){Test-AdminVerifyGuide $Fixture $hostBook $otherBook;return}
        $wire=[string](Probe 'Definition' @(35));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
        Check 'AdminVerify.Catalog.RegisteredOwnerAndCapability' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_ADDINS_VERIFY' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'ADMIN_MAINT')
        Check 'AdminVerify.Catalog.AbsentFrom34' ([string](Probe 'Definition' @(34)) -ceq '')
        Check 'AdminVerify.Catalog.PriorDefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
        foreach($case in @('AllPresent','Missing','ZeroByte','Disabled','ControlOff','UserOff')){
            $root=Join-Path $runRoot ('verify-'+$case);$addins=Join-Path $root 'Addins';New-Item -ItemType Directory -Path $addins -Force|Out-Null
            foreach($name in @('Core','Inventory.Domain','Designs.Domain','Operations','Admin')){
                if($case -ceq 'Missing' -and $name -ceq 'Admin'){continue}
                $body=if($case -ceq 'ZeroByte' -and $name -ceq 'Core'){''}else{'Non-executable file-presence fixture'}
                [IO.File]::WriteAllText((Join-Path $addins ('invSys.'+$name+'.xlam')),$body)
            }
            $mode=if($case -in @('Disabled','ControlOff','UserOff')){$case}else{'Enabled'}
            if(-not [bool](Probe 'SetPolicy' @($mode))){throw 'Explicit policy unavailable; not product RED.'}
            $policyHash=FileHash $Fixture.Config;$files=Pins $root;$before=@(Get-Slice4beActivityFiles $Fixture)
            $observation=([string](Run 'invSys.Admin.xlam' 'TestAdminVerify.Invoke' @($root,$hostBook.Name,$otherBook.Name))) -split '\|'
            if($observation.Count -ne 7){throw 'Owner observation unavailable; not product RED.'}
            $expected=if($case -in @('Missing','ZeroByte')){'FAILED'}else{'VERIFIED'}
            $prefix='AdminVerify.'+$case+'.'
            Check ($prefix+'ExpectedWorkbookResolvedAtEntry') ($observation[3] -ceq 'True')
            Write-Output ('Verify other sheet count at entry/exit: '+$observation[4]+'/'+$observation[5])
            Check ($prefix+'OwnerInvokedExactlyOnce') ($observation[0] -ceq '1')
            Check ($prefix+'ActualCallbackDisplaysOwnerResult') ($observation[1] -ceq $expected)
            Check ($prefix+'NoSpuriousTrackingNotice') ($observation[2] -ceq '0')
            Check ($prefix+'FilesNotPublishedOrChanged') (SamePins $files (Pins $root))
            Check ($prefix+'CapturedHostAndOtherWorkbookPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.HostPreserved' @($hostBook.Name)) -and [string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary' -and $otherBook.Worksheets.Count -eq $otherSheetCount)
            Check ($prefix+'HostCustomValuePreserved') ([bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.HostPreserved' @($hostBook.Name)))
            Check ($prefix+'OtherWorkbookValuePreserved') ([string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
            Check ($prefix+'OtherWorkbookSheetCountPreserved') ($otherBook.Worksheets.Count -eq $otherSheetCount)
            Check ($prefix+'InventoryAndPolicyPreserved') ((FileHash $inventory) -ceq $inventoryHash -and (FileHash $Fixture.Config) -ceq $policyHash)
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ $id)
            if($mode -ne 'Enabled'){Check ($prefix+'RecordingSuppressed') ($records.Count -eq 0);continue}
            $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($records|Where-Object OutcomeCode -CEQ $expected)
            $pair=$records.Count -eq 2 -and $attempts.Count -eq 1 -and $outcomes.Count -eq 1
            Check ($prefix+'ExactRecordedPair') ($pair -and $attempts[0].ActivityId -ceq $outcomes[0].ActivityId -and $attempts[0].RecordId -cne $outcomes[0].RecordId)
            $safe=$pair;$bound=$pair;$conclusion=$pair
            foreach($record in $records){
                $bound=$bound -and $record.OwnerId -ceq 'ADMIN_ADDINS_VERIFY' -and $record.CatalogVersion -eq $catalogVersion -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown'
                $json=$record|ConvertTo-Json -Depth 20 -Compress
                $safe=$safe -and @($record.SourceEventRefs).Count -eq 0
                foreach($value in @($root,$Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'Operator_Custom_Canary')){$safe=$safe -and -not $json.Contains($value)}
                $conclusion=$conclusion -and ([bool](Probe 'Concludes' @($json)) -eq ($record.OutcomeCode -ceq 'VERIFIED'))
            }
            Check ($prefix+'CapturedContextAndHonestEffect') $bound
            Check ($prefix+'NoPathsValuesOrDomainReferences') $safe
            Check ($prefix+'OnlyVerifiedConcludesCommand') $conclusion
        }
        Check 'AdminVerify.OtherWarehousePreserved' (SamePins $otherPins (Pins $Other.Root))
        Test-AdminVerifyBoundaries $Fixture $Other $hostBook $otherBook
        if($AdminVerifyGuidePaths){Test-AdminVerifyGuide $Fixture $hostBook $otherBook}
    }finally{$hostBook.Close($false);$otherBook.Close($false)}
}
