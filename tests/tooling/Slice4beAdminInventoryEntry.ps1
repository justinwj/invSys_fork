# Disposable-fixture entry audit. Invoke the packaged Ribbon and modal OK handler.
# Adapters supply fixture fields, expose fixed facts and suppress raw owner notices.
function Install-AdminInventoryEntryProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $control=$project.VBComponents.Add(2);$control.Name='TestInventoryEntryRibbon'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminAddInventoryItem"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryEntry'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mRoot As String, mSecret As String, mOtherBook As String
Private mHostName As String, mCapturedHost As Boolean
Private mShows As Long, mAccepts As Long, mNotices As Long
Private mAccepted As Boolean, mYieldSucceeded As Boolean
Private mPayloadSeen As Boolean, mLegacyPayload As Boolean, mLegacyDisplay As Boolean
Public Sub Execute(ByVal hostName As String, ByVal mode As String, _
                   ByVal targetRoot As String, ByVal secret As String, ByVal otherBook As String)
    Dim control As New TestInventoryEntryRibbon
    mMode = mode: mRoot = targetRoot: mSecret = secret: mOtherBook = otherBook
    mHostName = hostName: mCapturedHost = False
    mShows = 0: mAccepts = 0: mNotices = 0: mAccepted = False: mYieldSucceeded = False
    mPayloadSeen = False: mLegacyPayload = False: mLegacyDisplay = False
    Application.Workbooks(hostName).Activate
    RibbonOnActionAdmin control
    mSecret = ""
End Sub
Public Sub Drive(ByVal form As frmAddInventoryItem)
    Dim selected As String, signed As String
    mShows = mShows + 1
    If mShows > 1 Then form.InventoryEntryCancelForTest: Exit Sub
    mLegacyDisplay = form.InventoryEntryLegacyDisplayForTest()
    mCapturedHost = form.InventoryEntryOwnerBoundForTest(mHostName)
    Select Case mMode
        Case "ActiveWorkbook"
            Application.Workbooks(mOtherBook).Activate
            mYieldSucceeded = True
        Case "SignedOut"
            modAuth.SignOut
            mYieldSucceeded = Not modAuth.IsSignedIn()
        Case "OtherWarehouseAdmin"
            modRuntimeWorkbooks.SetCoreDataRootOverride mRoot
            selected = modNasConnection.SelectWarehouseTargetForAutomation(mRoot, mRoot, "S1", False)
            modNasConnection.SetCurrentTargetPathsForTest "\\fixture-host\config-command", mRoot
            signed = modAuth.SignInCurrentTargetForAutomation("config-admin", mSecret, "")
            mYieldSucceeded = (Left$(selected, 3) = "OK|" And Left$(signed, 3) = "OK|")
        Case "ReplacementSession"
            modAuth.SignOut
            signed = modAuth.SignInCurrentTargetForAutomation("config-admin", mSecret, "")
            mYieldSucceeded = (Left$(signed, 3) = "OK|")
        Case "Baseline"
            mYieldSucceeded = True
    End Select
    If Not mYieldSucceeded Then form.InventoryEntryCancelForTest: Exit Sub
    mAccepts = mAccepts + 1
    mAccepted = form.InventoryEntryAcceptForTest()
    If Not mAccepted Then form.InventoryEntryCancelForTest
End Sub
Public Sub Notice()
    mNotices = mNotices + 1
End Sub
Public Sub RecordPayload(ByVal payload As String)
    mPayloadSeen = True
    mLegacyPayload = (InStr(1, payload, Chr$(34) & "ROW" & Chr$(34), vbTextCompare) > 0)
End Sub
Public Function Facts() As String
    Facts = CStr(mShows) & "|" & CStr(mAccepts) & "|" & CStr(mAccepted) & "|" & _
            CStr(mYieldSucceeded) & "|" & CStr(mNotices) & "|" & CStr(mPayloadSeen) & "|" & _
            CStr(mLegacyPayload) & "|" & CStr(mLegacyDisplay) & "|" & CStr(mCapturedHost)
End Function
'@)
    $form=$project.VBComponents.Item('frmAddInventoryItem').CodeModule
    $start=$form.ProcStartLine('UserForm_Activate',0);$count=$form.ProcCountLines('UserForm_Activate',0)
    $body=$form.Lines($start,$count)
    $boundary='    If Not mAnchors Is Nothing Then mAnchors.ResizeControls'
    if([regex]::Matches($body,[regex]::Escape($boundary)).Count -ne 1){throw 'Inventory activation observation boundary changed; not product RED.'}
    $body=$body.Replace($boundary,($boundary+"`r`n"+'    TestInventoryEntry.Drive Me'))
    $form.DeleteLines($start,$count);$form.InsertLines($start,$body)
    $form.AddFromString(@'
Public Function InventoryEntryAcceptForTest() As Boolean
    mTxtItemName.Value = "Inventory entry fixture"
    mCmbUom.Value = "LB"
    mCmbLocation.Value = "CLEARVIEW"
    mTxtQty.Value = "1"
    mBtnOK_Click
    InventoryEntryAcceptForTest = mAccepted
End Function
Public Sub InventoryEntryCancelForTest()
    mBtnCancel_Click
End Sub
Public Function InventoryEntryLegacyDisplayForTest() As Boolean
    InventoryEntryLegacyDisplayForTest = (InStr(1, mLblGenerated.Caption, "ROW:", vbTextCompare) > 0)
End Function
Public Function InventoryEntryOwnerBoundForTest(ByVal hostName As String) As Boolean
    Dim expected As Workbook
    Set expected = Application.Workbooks(hostName)
    InventoryEntryOwnerBoundForTest = (mOperatorWorkbook Is expected)
End Function
'@)
    $owner=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('AddInventoryItemForWarehouse',0);$count=$owner.ProcCountLines('AddInventoryItemForWarehouse',0)
    $body=$owner.Lines($start,$count)
    $boundary='    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)'
    if([regex]::Matches($body,[regex]::Escape($boundary)).Count -ne 1){throw 'Inventory payload observation boundary changed; not product RED.'}
    $body=$body.Replace($boundary,($boundary+"`r`n"+'    TestInventoryEntry.RecordPayload payloadJson'))
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $owner.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal Title As String = "") As VbMsgBoxResult
    TestInventoryEntry.Notice
    MsgBox = vbOK
End Function
'@)
}

function Get-AdminInventoryEntryFacts($Fixture) {
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$null;$opened=$false
    foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$book=$candidate;break}}
    if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$true);$opened=$true}
    try {
        $log=Table $book 'tblInventoryLog';$applied=Table $book 'tblAppliedEvents'
        $legacy=@($book.Worksheets|ForEach-Object {$_.ListObjects}|ForEach-Object {$_.ListColumns}|Where-Object {$_.Name.Trim().ToUpperInvariant() -ceq 'ROW'}).Count
        $keys=@();foreach($row in $log.ListRows){$key=[string]$row.Range.Cells.Item(1,$log.ListColumns.Item('System_Key').Index).Value2;if($key){$keys+=$key}}
        $lastDelta=if($log.ListRows.Count){[double]$log.ListRows.Item($log.ListRows.Count).Range.Cells.Item(1,$log.ListColumns.Item('QtyDelta').Index).Value2}else{0}
        [pscustomobject]@{LogRows=$log.ListRows.Count;AppliedRows=$applied.ListRows.Count;LegacyHeaders=$legacy;DistinctKeys=@($keys|Select-Object -Unique).Count;LastDelta=$lastDelta;Saved=$book.Saved;FileHash=(Get-FileHash -LiteralPath $path).Hash}
    } finally {if($opened){$book.Close($false)}}
}

function Test-AdminInventoryEntry($Fixture,$Other) {
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    Check 'InventoryEntry.GeneratedFixtureSeedApplied' $seed.StartsWith('OK|')
    if(-not $seed.StartsWith('OK|')){throw 'Generated inventory fixture unavailable; not product RED.'}
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add()
    $hostBook.SaveAs((Join-Path $runRoot 'inventory-entry-host.xlsx'),51)
    $otherBook.SaveAs((Join-Path $runRoot 'inventory-entry-unrelated.xlsx'),51)
    $otherSheetCount=$otherBook.Worksheets.Count
    $authHash=(Get-FileHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))).Hash
    $configHash=(Get-FileHash $Fixture.Config).Hash
    $otherInventoryHash=(Get-FileHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Data.Inventory.xlsb'))).Hash
    $otherAuthHash=(Get-FileHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb'))).Hash
    $otherConfigHash=(Get-FileHash $Other.Config).Hash
    try {
        foreach($mode in @('Baseline','ActiveWorkbook','SignedOut','OtherWarehouseAdmin','ReplacementSession')){
            SelectTarget $Fixture
            $before=Get-AdminInventoryEntryFacts $Fixture
            $target=if($mode -ceq 'OtherWarehouseAdmin'){$Other}else{$Fixture}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryEntry.Execute' @($hostBook.Name,$mode,$target.Root,$target.Secret,$otherBook.Name))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryEntry.Facts')).Split('|')
            $after=Get-AdminInventoryEntryFacts $Fixture
            $actionReached=$facts.Count -eq 9 -and [int]$facts[1] -eq 1 -and $facts[2] -ceq 'True' -and $facts[3] -ceq 'True'
            Check ('InventoryEntry.'+$mode+'.ActualModalAcceptReached') $actionReached
            if(-not $actionReached){throw 'Inventory modal action fixture did not reach OK; not product RED.'}
            Check ('InventoryEntry.'+$mode+'.OriginalHostCapturedBeforeYield') ($facts[8] -ceq 'True')
            $wrote=$after.LogRows -gt $before.LogRows -or $after.AppliedRows -gt $before.AppliedRows
            if($mode -cin @('Baseline','ActiveWorkbook')){
                Check ('InventoryEntry.'+$mode+'.OwningCreateApplied') ($after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
                Check ('InventoryEntry.'+$mode+'.NewDistinctEntity') ($after.DistinctKeys -eq $before.DistinctKeys+1)
                Check ('InventoryEntry.'+$mode+'.ExactStartingQuantitySaved') ($after.LastDelta -eq 1 -and $after.Saved -eq $true -and $after.FileHash -cne $before.FileHash)
                Check ('InventoryEntry.'+$mode+'.CreatePayloadContainsNoROW') ($facts[5] -ceq 'True' -and $facts[6] -ceq 'False')
            }elseif($mode -cne 'ReplacementSession'){
                Check ('InventoryEntry.'+$mode+'.NoInventoryMutation') (-not $wrote)
            }
            Check ('InventoryEntry.'+$mode+'.ManagedHeadersContainNoROW') ($after.LegacyHeaders -eq 0)
            Check ('InventoryEntry.'+$mode+'.DisplayContainsNoROWIdentity') ($facts[7] -ceq 'False')
            Check ('InventoryEntry.'+$mode+'.UnrelatedWorkbookStructurePreserved') ($otherBook.Worksheets.Count -eq $otherSheetCount)
            [pscustomobject]@{Case=$mode;ModalAcceptReached=$actionReached;InventoryWriteObserved=$wrote;SavedInventoryBytesChanged=($after.Saved -eq $true -and $after.FileHash -cne $before.FileHash);ReplacementSessionRefusalContractNotAssumed=($mode -ceq 'ReplacementSession')}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'inventory-entry-observations.jsonl')
        }
        Check 'InventoryEntry.SourceAuthAndConfigPreserved' ($authHash -ceq (Get-FileHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))).Hash -and $configHash -ceq (Get-FileHash $Fixture.Config).Hash)
        Check 'InventoryEntry.OtherWarehouseInventoryPreserved' ($otherInventoryHash -ceq (Get-FileHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Data.Inventory.xlsb'))).Hash)
        Check 'InventoryEntry.OtherWarehouseAuthorityPreserved' ($otherAuthHash -ceq (Get-FileHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb'))).Hash -and $otherConfigHash -ceq (Get-FileHash $Other.Config).Hash)
    } finally {$hostBook.Close($false);$otherBook.Close($false)}
}
