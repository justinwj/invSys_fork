# Actual packaged actions on Admin-generated disposable fixtures. Values stay in memory.
function Install-InventoryCommandBoundaryProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryCommands'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mDriving As Boolean, mAction As String, mCase As String, mPhase As String
Private mRoot As String, mWh As String, mSecret As String, mOtherRoot As String
Private mHost As String, mSku As String, mBefore As String
Private mShows As Long, mHandler As Long, mYield As Boolean, mPrepared As Boolean
Private mBound As Boolean, mSignatureReady As Boolean
Public Function IsDriving() As Boolean
    IsDriving = mDriving
End Function
Public Function PrepareBoundaryTable(ByVal host As String) As Boolean
    Dim name As String, report As String, wb As Workbook, lo As ListObject, ws As Worksheet
    Set wb = Application.Workbooks(host)
    If Not modAdminInventoryWorksheet.CreateInventoryWorksheetTable(wb, name, report) Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = name Then
                lo.ListColumns("Action").DataBodyRange.Cells(1, 1).Value2 = "ADD"
                lo.ListColumns("Item Name").DataBodyRange.Cells(1, 1).Value2 = "Zero worksheet fixture"
                lo.ListColumns("UOM").DataBodyRange.Cells(1, 1).Value2 = "LB"
                lo.ListColumns("Qty Mode").DataBodyRange.Cells(1, 1).Value2 = "COUNTED"
                lo.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 0
                lo.ListColumns("Default Location").DataBodyRange.Cells(1, 1).Value2 = "CLEARVIEW"
                wb.Save
                lo.DataBodyRange.Cells(1, 1).Select
                PrepareBoundaryTable = True
                Exit Function
            End If
        Next lo
    Next ws
End Function
Public Sub SetSelection(ByVal otherRoot As String, ByVal sku As String)
    mOtherRoot = otherRoot: mSku = sku
End Sub
Public Sub Execute(ByVal host As String, ByVal action As String, ByVal caseName As String, _
                   ByVal phase As String, ByVal root As String, ByVal wh As String, _
                   ByVal secret As String)
    Dim control As New TestInventoryEntryRibbon
    mDriving = True: mHost = host: mAction = action: mCase = caseName: mPhase = phase
    mRoot = root: mWh = wh: mSecret = secret
    mShows = 0: mHandler = 0: mYield = False: mPrepared = False: mBound = False
    mBefore = "": mSignatureReady = False
    Application.Workbooks(host).Activate
    On Error GoTo Failed
    RibbonOnActionAdmin control
    mDriving = False: mSecret = ""
    Exit Sub
Failed:
    mDriving = False: mSecret = ""
    Err.Raise vbObjectError + 2101, , "Inventory action adapter failed."
End Sub
Public Sub Drive(ByVal form As frmAddInventoryItem)
    mShows = mShows + 1
    If mShows > 1 Then form.InventoryEntryCancelForTest: Exit Sub
    mBound = form.InventoryEntryOwnerBoundForTest(mHost)
    mPrepared = form.InventoryCommandPrepareForTest(mAction, mSku)
    If Not mPrepared Then form.InventoryEntryCancelForTest: Exit Sub
    If mPhase = "ModalReturn" Then YieldAt mPhase
    mHandler = mHandler + 1
    form.InventoryCommandInvokeForTest mAction
    If form.Visible Then form.InventoryEntryCancelForTest
End Sub
Public Sub YieldAt(ByVal phase As String)
    If Not mDriving Or mYield Or phase <> mPhase Then Exit Sub
    Dim signed As String, selected As String, fields As Object, report As String
    Select Case mCase
        Case "Authorized", "Zero", "Utility"
            mYield = True
        Case "SignedOut"
            modAuth.SignOut
            mYield = Not modAuth.IsSignedIn()
        Case "ReplacementSession", "Reader", "OtherWarehouseAdmin"
            modAuth.SignOut
            If mCase = "OtherWarehouseAdmin" Then
                modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
                selected = modNasConnection.SelectWarehouseTargetForAutomation(mOtherRoot, mOtherRoot, "S1", False)
                modNasConnection.SetCurrentTargetPathsForTest "\\fixture-host\config-command", mOtherRoot
            End If
            signed = modAuth.SignInCurrentTargetForAutomation(IIf(mCase = "Reader", "config-reader", "config-admin"), mSecret, "")
            mYield = (Left$(signed, 3) = "OK|")
        Case "CatalogChanged"
            Set fields = CreateObject("Scripting.Dictionary")
            mYield = modAdmin.UpdateInventoryItemCatalogForWarehouse(mWh, mSku, _
                "Externally changed fixture catalog", "LB", "CLEARVIEW", , , , , , , fields, report)
    End Select
    If mYield Then
        mBefore = InventorySignature()
        mSignatureReady = True
    End If
End Sub
Public Function DialogAnswer(ByVal prompt As String) As String
    If InStr(1, prompt, "Why the edit", vbTextCompare) > 0 Then
        YieldAt "EditReason"
    ElseIf InStr(1, prompt, "deleted from managed", vbTextCompare) > 0 Then
        YieldAt "DeleteReason"
    End If
    DialogAnswer = "Disposable fixture reason"
End Function
Public Function DialogChoice(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle) As VbMsgBoxResult
    If (buttons And vbYesNo) = vbYesNo Then
        YieldAt "DeleteConfirmation"
        DialogChoice = vbYes
    Else
        DialogChoice = vbOK
    End If
End Function
Private Function InventorySignature() As String
    Dim wb As Workbook, candidate As Workbook, prior As Workbook, opened As Boolean
    Dim ws As Worksheet, lo As ListObject, cell As Range, result As String, path As String
    Set prior = Application.ActiveWorkbook
    path = mRoot & "\" & mWh & ".invSys.Data.Inventory.xlsb"
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
        opened = True
    End If
    On Error GoTo Failed
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            result = result & lo.Name & "|" & CStr(lo.ListRows.Count)
            For Each cell In lo.Range.Cells
                If IsError(cell.Value2) Then
                    result = result & "|Error"
                Else
                    result = result & "|" & CStr(Len(CStr(cell.Value2))) & ":" & CStr(cell.Value2)
                End If
            Next cell
        Next lo
    Next ws
    InventorySignature = result
CleanExit:
    If opened Then wb.Close SaveChanges:=False
    If Not prior Is Nothing Then prior.Activate
    Exit Function
Failed:
    If opened Then wb.Close SaveChanges:=False
    If Not prior Is Nothing Then prior.Activate
    Err.Raise vbObjectError + 2102, , "Inventory observation unavailable."
End Function
Public Function Facts() As String
    Dim unchanged As Boolean
    If mSignatureReady Then unchanged = (mBefore = InventorySignature())
    Facts = CStr(mPrepared) & "|" & CStr(mYield) & "|" & CStr(mHandler) & "|" & _
        CStr(mBound) & "|" & CStr(mSignatureReady And unchanged)
End Function
'@)
    $form=$project.VBComponents.Item('frmAddInventoryItem').CodeModule
    $form.AddFromString(@'
Public Function InventoryCommandPrepareForTest(ByVal action As String, ByVal sku As String) As Boolean
    Dim i As Long
    If action = "Edit" Or action = "Delete" Then
        mBtnEditMode_Click
        For i = 0 To mCmbEditItem.ListCount - 1
            If CStr(mCmbEditItem.List(i, 1)) = sku Then
                mLoading = True: mCmbEditItem.ListIndex = i: mLoading = False
                mCmbEditItem_Change
                InventoryCommandPrepareForTest = (mSelectedEditSku = sku)
                Exit Function
            End If
        Next i
        Exit Function
    End If
    If action = "Add" Then
        mTxtItemName.Value = "Boundary fixture item"
        mCmbUom.Value = "LB": mCmbLocation.Value = "CLEARVIEW"
        mTxtQty.Value = "0"
    End If
    InventoryCommandPrepareForTest = True
End Function
Public Sub InventoryCommandInvokeForTest(ByVal action As String)
    Select Case action
        Case "Add", "Edit": mBtnOK_Click
        Case "Delete": mBtnDeleteItem_Click
        Case "Create": mBtnCreateInventoryTable_Click
        Case "Upload": mBtnUploadInventoryTable_Click
    End Select
End Sub
Private Function InputBox(ByVal Prompt As String, Optional ByVal Title As String = "", _
                          Optional ByVal Default As String = "") As String
    InputBox = TestInventoryCommands.DialogAnswer(Prompt)
End Function
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestInventoryCommands.DialogChoice(Prompt, Buttons)
End Function
'@)
}

function Get-InventoryBoundarySku($Fixture) {
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$null;$owned=$false
    foreach($b in $excel.Workbooks){if($b.FullName -ieq $path){$book=$b;break}}
    if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$true);$owned=$true}
    try {
        $catalog=Table $book 'tblSkuCatalog'
        foreach($row in $catalog.ListRows){
            $state=[string]$row.Range.Cells.Item(1,$catalog.ListColumns.Item('CATALOG_STATE').Index).Value2
            $sku=[string]$row.Range.Cells.Item(1,$catalog.ListColumns.Item('ITEM_CODE').Index).Value2
            if($sku -and $state -cne 'RETIRED'){return $sku}
        }
        throw 'No generated active catalog item is available; not product RED.'
    }finally{if($owned){$book.Close($false)}}
}

function Test-InventoryCommandBoundaries($Fixture,$Other) {
    $cases=@('Authorized','SignedOut','ReplacementSession','Reader','OtherWarehouseAdmin','CatalogChanged')
    foreach($action in @('Add','Edit','Create','Upload','Delete')){
        foreach($case in $cases){
            SelectTarget $Fixture
            $hostBook=$excel.Workbooks.Add()
            $hostBook.SaveAs((Join-Path $runRoot ('inventory-boundary-'+$action+'-'+$case+'.xlsx')),51)
            try {
                $sku=Get-InventoryBoundarySku $Fixture
                if($action -eq 'Upload'){
                    # Fixture preparation uses the public staging helper; measured action uses the real form handler.
                    $ready=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
                    if(-not $ready){throw 'Upload staging preparation unavailable; not product RED.'}
                }
                $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $hostHash=Get-AdminInventoryEntryHash $hostBook.FullName
                $phase=if($action -eq 'Edit'){'EditReason'}elseif($action -eq 'Delete'){'DeleteConfirmation'}else{'ModalReturn'}
                # The other warehouse uses its own ephemeral fixture credential.
                $secret=if($case -ceq 'OtherWarehouseAdmin'){$Other.Secret}else{$Fixture.Secret}
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,$action,$case,$phase,$Fixture.Root,$Fixture.Warehouse,$secret))
                $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
                $id='InventoryCommand.'+$action+'.'+$phase+'.'+$case
                $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
                Check ($id+'.ActualHandlerAndYieldReached') $ready
                if(-not $ready){throw 'Actual Inventory action/yield fixture unavailable; not product RED.'}
                Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
                $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
                if($case -ceq 'Authorized'){
                    if($action -in @('Add','Upload')){
                        Check ($id+'.ExactZeroEntityPersisted') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.AppliedRows -eq $before.AppliedRows+1)
                    }elseif($action -eq 'Create'){
                        Check ($id+'.OriginalOperatorStagingCreated') ([long](Run 'invSys.Admin.xlam' 'modAdminInventoryWorksheet.CountInventoryWorksheetTables' @($hostBook)) -eq 1)
                    }else{
                        Check ($id+'.OwningCommandChangesObserved') ($facts[4] -ceq 'False')
                    }
                }else{
                    Check ($id+'.NoInventoryMutationAfterYield') ($facts[4] -ceq 'True')
                    Check ($id+'.NoNewPersistedEvents') ($after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
                    Check ($id+'.NoOperatorSaveAfterYield') ($hostHash -ceq (Get-AdminInventoryEntryHash $hostBook.FullName))
                }
            }finally{$hostBook.Close($false)}
        }
    }
}
