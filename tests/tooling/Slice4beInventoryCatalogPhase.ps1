# Component gate: execute exact source catalog procedures in an unsaved probe
# module, with deterministic Core guard outcomes. No XLAM is built or saved.
# This supplements, and never replaces, packaged callback/real-authority gates.
function Install-InventoryCatalogPhaseProbe {
    $source=Get-Content (Join-Path $repo 'src/Admin/Modules/modAdmin.bas') -Raw
    $names=@('UpdateInventoryItemCatalogForWarehouse','FindOpenWorkbookByFullNameAdmin',
        'FindListObjectByNameAdminLocal','ColumnIndexAdminLocal','FindCatalogRowBySkuAdmin',
        'CatalogCellAdmin','SetCatalogValueAdmin','EnsureCatalogColumnAdmin',
        'NormalizeCatalogColumnNameAdmin','SetSheetProtectionAdminLocal',
        'SetPictureReferencesInCatalogAdmin','ParsePictureReferencesAdmin','AppendCatalogEditHistoryAdmin',
        'PlanInventoryCatalogFieldsAdmin','PlanInventoryCatalogValueAdmin')
    $parts=@('Option Explicit')
    foreach($name in $names){
        $matches=[regex]::Matches($source,('(?ms)^(?:Public|Private) (?:Function|Sub) '+$name+'\b.*?^End (?:Function|Sub)\b'))
        if($matches.Count -ne 1){
            if($name -like 'PlanInventoryCatalog*' -and $matches.Count -eq 0){continue}
            throw ('Catalog source procedure unavailable: '+$name)
        }
        $parts+=$matches[0].Value.Replace('modInventoryCommands.','TestCatalogPhase.')
    }
    $hasContext=$parts[1] -match 'Optional ByVal interactiveContext'
    $tail=if($hasContext){', "component-context"'}else{''}
    $parts+=@'
Public Function ExecuteCatalogPhase(ByVal warehouse As String, ByVal sku As String, ByVal mode As String) As String
    Dim fields As Object, report As String, succeeded As Boolean
    Set fields = CreateObject("Scripting.Dictionary")
    fields.Add "PHASE CUSTOM", "planned custom"
    fields.Add "EDIT_HISTORY_APPEND", "Catalog phase test"
    TestCatalogPhase.Reset mode
    succeeded = UpdateInventoryItemCatalogForWarehouse(warehouse, sku, "Catalog phase updated", _
        "LB", "CLEARVIEW", "planned description", "planned vendor", "planned code", _
        "planned category", "planned external", "one.png;two.png", fields, report__TAIL__)
    ExecuteCatalogPhase = CStr(succeeded) & "|" & CStr(InStr(1, report, "partial", vbTextCompare) > 0) & _
        "|" & TestCatalogPhase.Facts()
End Function
'@.Replace('__TAIL__',$tail)
    $project=$packages['invSys.Admin.xlam'].VBProject
    $owner=$project.VBComponents.Add(1);$owner.Name='TestCatalogOwner'
    $owner.CodeModule.AddFromString(($parts -join "`r`n"))
    $guard=$project.VBComponents.Add(1);$guard.Name='TestCatalogPhase'
    $guard.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mBook As Workbook, mPlans As Object
Private mBefore As Long, mAfter As Long, mAtCheck As Long, mConfirms As Long
Private mPlanningDirty As Boolean, mWritten As Boolean
Public Sub Reset(ByVal mode As String)
    mMode = mode: Set mBook = Nothing
    Set mPlans = CreateObject("Scripting.Dictionary"): mPlans.CompareMode = vbTextCompare
    mBefore = 0: mAfter = 0: mAtCheck = 0: mConfirms = 0
    mPlanningDirty = False: mWritten = False
End Sub
Public Function BeginCatalog(ByVal key As String, ByVal sku As String, ByRef report As String) As Boolean
    BeginCatalog = True
End Function
Public Function BindCatalogWriter(ByVal key As String, ByVal name As String, ByRef report As String) As Boolean
    Set mBook = Application.Workbooks(name)
    BindCatalogWriter = True
End Function
Public Function PlanCatalogField(ByVal key As String, ByVal sku As String, ByVal header As String, _
        ByVal value As String, ByRef report As String) As Boolean
    If Not mBook.Saved Then mPlanningDirty = True
    If mMode = "PlanningDenied" And UCase$(header) = "PHASE CUSTOM" Then Exit Function
    mPlans(header) = value: PlanCatalogField = True
End Function
Public Function CheckForm(ByVal key As String, ByRef report As String) As Boolean
    mBefore = mBefore + 1: mAtCheck = mPlans.Count
    CheckForm = (mMode <> "BeforeWriteDenied")
End Function
Public Function BeginCatalogSave(ByVal key As String, ByRef report As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, column As ListColumn, header As Variant, found As Boolean
    mAfter = mAfter + 1: mWritten = True
    For Each ws In mBook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblSkuCatalog" Then
                For Each header In mPlans.Keys
                    found = False
                    For Each column In lo.ListColumns
                        If StrComp(Trim$(column.Name), header, vbTextCompare) = 0 Then
                            found = True
                            If CStr(column.DataBodyRange.Cells(1, 1).Value) <> CStr(mPlans(header)) Then mWritten = False
                        End If
                    Next column
                    If Not found Then mWritten = False
                Next header
            End If
        Next lo
    Next ws
    BeginCatalogSave = (mMode <> "BeforeSaveDenied")
End Function
Public Function ConfirmCatalog(ByVal key As String, ByRef report As String) As Boolean
    mConfirms = mConfirms + 1
    ConfirmCatalog = (mMode <> "SavedUnconfirmed" And mBook.Saved)
End Function
Public Function Facts() As String
    Facts = CStr(mBefore) & "|" & CStr(mAfter) & "|" & CStr(mAtCheck) & "|" & _
        CStr(mPlans.Count) & "|" & CStr(mConfirms) & "|" & CStr(mPlanningDirty) & "|" & CStr(mWritten)
End Function
'@)
    [pscustomobject]@{SourceComponentOnly=$true;PackagedAction=$false;DeterministicGuardSeam=$true;
        ProductSource=(Get-FileHash (Join-Path $repo 'src/Admin/Modules/modAdmin.bas')).Hash;
        XlamBuild=$false}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'catalog-component-scope.json')
}

function Get-CatalogPhaseMemory($Table) {
    # Generated fixture values are compared only in memory, never emitted.
    $items=[Collections.Generic.List[string]]::new()
    foreach($cell in $Table.Range.Cells){$items.Add([string]$cell.Formula)}
    $items -join "`n"
}

function Test-InventoryCatalogPhase($Fixture) {
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    Check 'CatalogPhase.GeneratedFixtureSeedApplied' $seed.StartsWith('OK|')
    if(-not $seed.StartsWith('OK|')){throw 'Generated catalog fixture unavailable; not product RED.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    foreach($book in @($excel.Workbooks)){if($book.FullName -ieq $path){$book.Close($false)}}
    $baseline=Join-Path $runRoot 'catalog-phase-generated-baseline.xlsb'
    Copy-Item -LiteralPath $path -Destination $baseline
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authHash=(Get-FileHash $auth).Hash;$configHash=(Get-FileHash $Fixture.Config).Hash
    foreach($mode in @('Authorized','DirtyBorrowed','PlanningDenied','BeforeWriteDenied','BeforeSaveDenied','SavedUnconfirmed')){
        Copy-Item -LiteralPath $baseline -Destination $path -Force
        $book=$excel.Workbooks.Open($path,0,$false)
        $componentStep='prepare generated catalog custom column'
        try{
            $table=Table $book 'tblSkuCatalog'
            $sku=[string]$table.ListColumns.Item('SKU').DataBodyRange.Cells.Item(1,1).Value2
            $table.Parent.Unprotect()
            $extra=$table.ListColumns.Add();$extra.Name='USER KEEP';$extra.DataBodyRange.Cells.Item(1,1).Value2='retained'
            $table.Parent.Protect()
            $book.Save()
            if($mode -eq 'DirtyBorrowed'){
                $table.Parent.Unprotect();$extra.DataBodyRange.Cells.Item(1,1).Value2='unsaved user value'
            }
            $memory=Get-CatalogPhaseMemory $table;$hash=(Get-FileHash $path).Hash;$saved=$book.Saved
            $componentStep='invoke exact source catalog helper'
            $facts=([string](Run 'invSys.Admin.xlam' 'TestCatalogOwner.ExecuteCatalogPhase' @($Fixture.Warehouse,$sku,$mode))).Split('|')
            if($facts.Count -ne 9){throw 'Catalog component facts unavailable; not product RED.'}
            $ok=$facts[0] -eq 'True';$partial=$facts[1] -eq 'True';$changed=(Get-FileHash $path).Hash -cne $hash
            $id='CatalogPhase.'+$mode
            if($mode -eq 'Authorized'){
                Check ($id+'.SucceededAndSaved') ($ok -and $changed -and $book.Saved)
                Check ($id+'.AllFieldsPlannedBeforeSingleCheck') ([int]$facts[2] -eq 1 -and [int]$facts[4] -eq [int]$facts[5] -and [int]$facts[5] -ge 13 -and $facts[7] -eq 'False')
                Check ($id+'.RecheckAfterCompleteWriteBatch') ([int]$facts[3] -eq 1 -and $facts[8] -eq 'True')
                Check ($id+'.ConfirmedAfterSave') ([int]$facts[6] -eq 1 -and -not $partial)
                Check ($id+'.UnknownColumnPreserved') ($extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'retained')
            }elseif($mode -eq 'SavedUnconfirmed'){
                Check ($id+'.NotSuccessful') (-not $ok)
                Check ($id+'.SavedOutputPreserved') $changed
                Check ($id+'.HonestPartialReport') $partial
            }else{
                Check ($id+'.Refused') (-not $ok)
                Check ($id+'.NoPersistedWrite') (-not $changed)
                if($mode -ne 'BeforeSaveDenied'){
                    Check ($id+'.MemoryAndSavedStatePreserved') ((Get-CatalogPhaseMemory $table) -ceq $memory -and $book.Saved -eq $saved)
                }else{
                    Check ($id+'.NoConfirmation') ([int]$facts[6] -eq 0)
                }
            }
            if($mode -in @('Authorized','SavedUnconfirmed')){
                $componentStep='independent persisted field read'
                $book.Close($false);$book=$null
                $diskHash=(Get-FileHash $path).Hash
                $cold=$excel.Workbooks.Open($path,0,$true)
                try{
                    $stored=Table $cold 'tblSkuCatalog'
                    $exact=$stored.ListColumns.Item('ITEM').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Catalog phase updated'
                    $exact=$exact -and $stored.ListColumns.Item('PHASE CUSTOM').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'planned custom'
                    $exact=$exact -and $stored.ListColumns.Item('IMAGE_PATH_2').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'two.png'
                    Check ($id+'.IndependentPersistedFields') $exact
                }finally{$cold.Close($false)}
                Check ($id+'.ColdReadPreservedBytes') ((Get-FileHash $path).Hash -ceq $diskHash)
            }
        }catch{throw ('Catalog component harness stage: '+$componentStep+'; status='+$_.Exception.HResult)}
        finally{if($null -ne $book){$book.Close($false)}}
    }
    Check 'CatalogPhase.AuthAndConfigPreserved' ((Get-FileHash $auth).Hash -ceq $authHash -and (Get-FileHash $Fixture.Config).Hash -ceq $configHash)
}
