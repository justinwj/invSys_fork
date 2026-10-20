# Owned, service-created test warehouses only. Business values and exact command /
# event identities remain in memory or ignored disposable workbook copies. Reports
# contain fixed check labels and booleans; no row values or raw errors are emitted.
function Install-InventoryEditReceiptProbe {
    $project=$packages['invSys.Core.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryEditReceipt'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mSource As String, mHostPath As String, mCopy As String
Private mReached As Boolean, mCopied As Boolean, mPlanned As String, mPreserved As Boolean
Public Sub Arm(ByVal mode As String, ByVal source As String, ByVal hostPath As String, ByVal copyPrefix As String)
    mMode = mode: mSource = source: mHostPath = hostPath: mCopy = copyPrefix
    mReached = False: mCopied = False: mPlanned = "": mPreserved = False
End Sub
Public Function BeforeQueue(ByVal plannedEvent As String) As Boolean
    Dim fso As Object, sourceHash As String, hostHash As String
    If mMode = "" Then Exit Function
    mReached = True: mPlanned = plannedEvent
    BeforeQueue = (mMode = "QueueFailure")
    On Error GoTo CopyFailed
    sourceHash = modTrainingWire.FileSha256(mSource)
    hostHash = modTrainingWire.FileSha256(mHostPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile mSource, mCopy & ".xlsb", False
    fso.CopyFile mHostPath, mCopy & ".xlsx", False
    mCopied = True
    mPreserved = (sourceHash <> "" And hostHash <> "" And _
        sourceHash = modTrainingWire.FileSha256(mSource) And _
        hostHash = modTrainingWire.FileSha256(mHostPath))
CopyFailed:
End Function
Public Function Facts() As String
    Facts = CStr(mReached) & "|" & CStr(mCopied) & "|" & CStr(mPreserved)
End Function
Public Function PlannedEventForOwnedTest() As String
    PlannedEventForOwnedTest = mPlanned
End Function
'@)
    $owner=$project.VBComponents.Item('modRoleEventWriter').CodeModule
    $start=$owner.ProcStartLine('QueueAdminInventoryAdjustEvent',0);$count=$owner.ProcCountLines('QueueAdminInventoryAdjustEvent',0)
    $body=$owner.Lines($start,$count)
    $boundary='(?im)^[\t ]*QueueAdminInventoryAdjustEvent = QueueEventCore\('
    if([regex]::Matches($body,$boundary).Count -ne 1){throw 'Owned queue observation boundary unavailable; not RED.'}
    $hook="    If TestInventoryEditReceipt.BeforeQueue(eventIdOut) Then`r`n        errorMessage = ""Owned queue failure.""`r`n        Exit Function`r`n    End If`r`n"
    $body=[regex]::Replace($body,$boundary,[Text.RegularExpressions.MatchEvaluator]{param($m) $hook+$m.Value})
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}
function Get-InventoryEditCell($Row,$Table,[string]$Header) {
    foreach($column in $Table.ListColumns){
        if(([string]$column.Name).Trim() -ieq $Header){return [string]$Row.Range.Cells.Item(1,$column.Index).Value2}
    }
    return ''
}
function Read-InventoryEditReceiptCopy([string]$Path,[string]$Sku,[string]$Command='') {
    $hash=Get-AdminInventoryEntryHash $Path
    $book=$excel.Workbooks.Open($Path,0,$true)
    $result=[pscustomobject]@{CatalogMatches=0;Item='';Reason='';History='';Custom='';Receipts=@();Applied=@();Log=@();Preserved=$false}
    try{
        foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){
            if($table.Name -ieq 'tblSkuCatalog'){
                foreach($row in $table.ListRows){if((Get-InventoryEditCell $row $table 'ITEM_CODE') -ceq $Sku){
                    $result.CatalogMatches++
                    $result.Item=Get-InventoryEditCell $row $table 'ITEM'
                    $result.Reason=Get-InventoryEditCell $row $table 'LAST_EDIT_REASON'
                    $result.History=Get-InventoryEditCell $row $table 'EDIT_HISTORY'
                    $result.Custom=Get-InventoryEditCell $row $table 'USER NOTE'
                }}
            }elseif($table.Name -ieq 'tblAppliedEvents'){
                foreach($row in $table.ListRows){$result.Applied+=,(Get-InventoryEditCell $row $table 'EventID')}
            }elseif($table.Name -ieq 'tblInventoryLog'){
                foreach($row in $table.ListRows){
                    if((Get-InventoryEditCell $row $table 'SKU') -ceq $Sku){
                        $result.Log+=,[pscustomobject]@{Event=(Get-InventoryEditCell $row $table 'EventID');Key=(Get-InventoryEditCell $row $table 'System_Key');Delta=(Get-InventoryEditCell $row $table 'QtyDelta')}
                    }
                }
            }elseif($table.Name -ieq 'tblInventoryEditReceipts'){
                foreach($row in $table.ListRows){
                    $rowSku=Get-InventoryEditCell $row $table 'SKU'
                    $rowCommand=Get-InventoryEditCell $row $table 'CommandId'
                    if($rowSku -ceq $Sku -or ($Command -ne '' -and $rowCommand -ceq $Command)){
                        $result.Receipts+=,[pscustomobject]@{
                            Command=$rowCommand;Sku=$rowSku
                            Warehouse=(Get-InventoryEditCell $row $table 'WarehouseId')
                            Station=(Get-InventoryEditCell $row $table 'StationId')
                            User=(Get-InventoryEditCell $row $table 'UserId')
                            Event=(Get-InventoryEditCell $row $table 'PlannedEventId')
                        }
                    }
                }
            }
        }}
    }finally{$book.Close($false)}
    $result.Preserved=((Get-AdminInventoryEntryHash $Path) -ceq $hash)
    return $result
}
function Get-InventoryEditSavedResult([string]$Path) {
    $hash=Get-AdminInventoryEntryHash $Path;$book=$excel.Workbooks.Open($Path,0,$true)
    try{
        $table=@($book.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
        $text=[string]$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2
        $match=[regex]::Match($text,'(?im)^Command:\s*([^\r\n]+)')
        $command='';if($match.Success){$command=$match.Groups[1].Value.Trim()}
        $status=[string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2
    }finally{$book.Close($false)}
    return [pscustomobject]@{Command=$command;Status=$status;Preserved=((Get-AdminInventoryEntryHash $Path) -ceq $hash)}
}
function Test-InventoryEditBinding($Snapshot,$Saved,$Fixture,[string]$Sku) {
    if($Saved.Command -eq '' -or $Snapshot.Receipts.Count -ne 1){return $false}
    $receipt=$Snapshot.Receipts[0]
    return ($receipt.Command -ceq $Saved.Command -and $receipt.Sku -ceq $Sku -and $receipt.Warehouse -ceq $Fixture.Warehouse -and $receipt.Station -ceq 'S1' -and $receipt.User -ceq 'config-admin')
}
function Test-InventoryEditReceiptPhases($Fixture,$Other) {
    foreach($case in @('CatalogOnly','NoopQuantity','QuantityChange','QueueFailure')){
        SelectTarget $Fixture
        $id='InventoryEditReceipt.'+$case
        $hostBook=$excel.Workbooks.Add();$hostPath=Join-Path $runRoot ('edit-receipt-'+$case+'.xlsx');$hostBook.SaveAs($hostPath,51)
        $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        try{
            $prepared=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
            if(-not $prepared){throw 'Owned operator table unavailable; not RED.'}
            $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
            $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Retained receipt extension'
            $hostBook.Save();$initial=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ReceiptSetup',$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $created=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($created.EntityQuantities.Keys|Where-Object {-not $initial.EntityQuantities.ContainsKey($_)})
            $ready=$newKeys.Count -eq 1 -and $created.EntityQuantities[$newKeys[0]] -eq 0 -and $created.LogRows -eq $initial.LogRows+1 -and $created.AppliedRows -eq $initial.AppliedRows+1
            Check ($id+'.OwningAddFixtureSaved') $ready
            if(-not $ready){throw 'Owning ADD fixture unavailable; not RED.'}
            $key=[string]$newKeys[0];$sku=[string]$table.ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2
            $item='Edited receipt '+$case;$reason='Owned receipt reason '+$case
            $table.ListColumns.Item('Action').DataBodyRange.Cells.Item(1,1).Value2='EDIT'
            $table.ListColumns.Item('Item Name').DataBodyRange.Cells.Item(1,1).Value2=$item
            $table.ListColumns.Item('Edit Reason').DataBodyRange.Cells.Item(1,1).Value2=$reason
            $quantity='';if($case -eq 'NoopQuantity'){$quantity='0'}elseif($case -in @('QuantityChange','QueueFailure')){$quantity='5'}
            $table.ListColumns.Item('Quantity').DataBodyRange.Cells.Item(1,1).Value2=$quantity
            $table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2=''
            $table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=''
            $hostBook.Save();$table.DataBodyRange.Cells.Item(1,1).Select()
            $prefix=Join-Path $runRoot ('before-queue-'+$case)
            [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @($case,$sourcePath,$hostPath,$prefix))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ReceiptEdit',$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $reached=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
            Check ($id+'.ActualCapturedUploadReached') $reached
            if(-not $reached){throw 'Actual EDIT action unavailable; not RED.'}
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $savedCopy=Join-Path $runRoot ('saved-edit-'+$case+'.xlsx');Copy-Item -LiteralPath $hostPath -Destination $savedCopy
            $saved=Get-InventoryEditSavedResult $savedCopy
            $snapshot=Read-InventoryEditReceiptCopy $sourcePath $sku $saved.Command
            Check ($id+'.SavedCatalogAndReason') ($snapshot.CatalogMatches -eq 1 -and $snapshot.Item -ceq $item -and $snapshot.Reason -ceq $reason)
            Check ($id+'.SingleCatalogApplication') ([regex]::Matches($snapshot.History,[regex]::Escape($reason)).Count -eq 1)
            Check ($id+'.ReciprocalSavedCommandAndSource') (Test-InventoryEditBinding $snapshot $saved $Fixture $sku)
            Check ($id+'.CustomValuesPreserved') ($snapshot.Custom -ceq 'Retained receipt extension' -and [string]$extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Retained receipt extension')
            Check ($id+'.ReadOnlyInspectionPreservedBytes') ($snapshot.Preserved -and $saved.Preserved)
            if($case -ne 'QueueFailure'){Check ($id+'.CompletionSavedInOperator') ($saved.Status -ceq 'IMPORTED')}
            $delta=0;if($case -eq 'QuantityChange'){$delta=1}
            $expectedQty=0;if($delta){$expectedQty=5}
            Check ($id+'.OnlyIntendedQuantityEffect') ($after.LogRows -eq $created.LogRows+$delta -and $after.AppliedRows -eq $created.AppliedRows+$delta -and $after.EntityQuantities.Count -eq $created.EntityQuantities.Count -and $after.EntityQuantities[$key] -eq $expectedQty)
            $queue=([string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Facts')).Split('|')
            $planned=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PlannedEventForOwnedTest')
            if($case -in @('QuantityChange','QueueFailure')){
                $observed=$queue.Count -eq 3 -and $queue[0] -ceq 'True' -and $queue[1] -ceq 'True' -and $queue[2] -ceq 'True'
                Check ($id+'.QueueBoundaryObservedWithoutSourceMutation') $observed
                if(-not $observed){throw 'Queue boundary snapshot unavailable; not RED.'}
                $beforeQueue=Read-InventoryEditReceiptCopy ($prefix+'.xlsb') $sku
                $beforeResult=Get-InventoryEditSavedResult ($prefix+'.xlsx')
                $joint=Test-InventoryEditBinding $beforeQueue $beforeResult $Fixture $sku
                Check ($id+'.CatalogReceiptAndPlannedEventSavedBeforeQueue') ($joint -and $planned -ne '' -and $beforeQueue.Receipts[0].Event -ceq $planned -and $beforeQueue.Item -ceq $item -and $beforeQueue.Reason -ceq $reason -and $beforeQueue.Preserved -and $beforeResult.Preserved)
                if($case -eq 'QuantityChange'){
                    $exact=@($snapshot.Log|Where-Object {$_.Event -ceq $planned -and $_.Key -ceq $key -and [double]$_.Delta -eq 5})
                    Check ($id+'.PlannedEventHasExactAppliedAndLogProof') ($planned -ne '' -and @($snapshot.Applied|Where-Object {$_ -ceq $planned}).Count -eq 1 -and $exact.Count -eq 1)
                }
            }else{
                Check ($id+'.NoInventedQuantityEvent') ($queue[0] -ceq 'False' -and $snapshot.Receipts.Count -eq 1 -and $snapshot.Receipts[0].Event -eq '')
            }
            if($case -eq 'QueueFailure'){
                Check ($id+'.IncompleteQuantityNotMarkedImported') ($saved.Status -cne 'IMPORTED')
                $hostBook.Close($false);$hostBook=$excel.Workbooks.Open($hostPath,0,$false)
                $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                SelectTarget $Fixture;$hostBook.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('Retry',$sourcePath,$hostPath,($prefix+'-retry')))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ReceiptRetry',$hostBook.Name))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
                $retry=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $retrySnapshot=Read-InventoryEditReceiptCopy $sourcePath $sku
                $retryEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PlannedEventForOwnedTest')
                Check ($id+'.FreshRetryKeepsPlannedEvent') ($planned -ne '' -and $retryEvent -ceq $planned)
                Check ($id+'.FreshRetryDoesNotRepeatCatalog') ($retrySnapshot.History -ceq $snapshot.History -and $retrySnapshot.Item -ceq $item -and $retrySnapshot.Preserved)
                Check ($id+'.FreshRetryAppliesQuantityOnce') ($retry.EntityQuantities.Count -eq $after.EntityQuantities.Count -and $retry.EntityQuantities[$key] -eq 5 -and $retry.LogRows -eq $after.LogRows+1 -and $retry.AppliedRows -eq $after.AppliedRows+1)
                $exact=@($retrySnapshot.Log|Where-Object {$_.Event -ceq $planned -and $_.Key -ceq $key -and [double]$_.Delta -eq 5})
                Check ($id+'.RetryPlannedEventHasExactAppliedAndLogProof') ($planned -ne '' -and @($retrySnapshot.Applied|Where-Object {$_ -ceq $planned}).Count -eq 1 -and $exact.Count -eq 1)
            }
        }finally{$hostBook.Close($false)}
    }
}
