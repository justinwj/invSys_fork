# D14 prerequisite: actual View Warehouses callback, generated read scope only.
. (Join-Path $PSScriptRoot 'Slice4beAdminDirectoryRecording.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminWarehouseRootWrite.ps1')
function Install-AdminDirectoryColumnsProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $ribbon=$project.VBComponents.Add(2);$ribbon.Name='TestDirectoryRibbon'
    $ribbon.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminWarehouses"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestDirectoryColumns'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mHost As Workbook, mRoot As String, mCalls As Long, mOpened As Boolean
Private mSuccess As Long, mOther As Long, mPrompts As Long, mManaged As String
Private mTextAddress As String, mFormulaAddress As String, mTailAddress As String
Private mTextIndex As Long, mFormulaIndex As Long, mTailIndex As Long
Public Function Prepare(ByVal name As String, ByVal root As String) As Boolean
    Dim report As String, ws As Worksheet
    Set mHost = Application.Workbooks(name): mRoot = root
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(mHost, report) Then Exit Function
    If Not modAdminConsole.EnsureAdminSchema(mHost, report) Then Exit Function
    Set ws = mHost.Worksheets.Add: ws.Name = "Personal_notes"
    ws.Range("A1").Value2 = "Preserve unrelated notes"
    modAdminConsole.RememberWarehouseScanRoot root
    Prepare = True
End Function
Public Function ReadRoots() As Collection
    Dim roots As New Collection
    roots.Add mRoot
    Set ReadRoots = roots
End Function
Public Sub SetRoot(ByVal root As String)
    mRoot = root
End Sub
Public Function NoResults() As Boolean
    Dim lo As ListObject
    Set lo = Directory()
    NoResults = (ManagedSignature() = "" And InStr(1, CStr(lo.ListColumns("Status").DataBodyRange.Cells(1, 1).Value2), "No warehouse config workbooks found.", vbBinaryCompare) = 1)
End Function
Public Sub OwnerEntered()
    mCalls = mCalls + 1
End Sub
Public Sub OwnerOpened()
    mOpened = (Application.ActiveWorkbook Is mHost)
    If mOpened Then mOpened = (Application.ActiveSheet.Name = "WarehouseDirectory")
End Sub
Public Function ObserveMessage(ByVal text As String) As VbMsgBoxResult
    If text = "Warehouse directory refreshed." Then mSuccess = mSuccess + 1 Else mOther = mOther + 1
    ObserveMessage = vbOK
End Function
Public Function ObservePrompt() As String
    mPrompts = mPrompts + 1
End Function
Public Function Invoke() As String
    Dim control As New TestDirectoryRibbon, selected As Boolean
    mHost.Activate: mHost.Worksheets("Personal_notes").Activate
    If Not (modAdmin.ResolveInteractiveAdminWorkbook() Is mHost) Then Err.Raise 5, , "Directory host not bound"
    mCalls = 0: mOpened = False: mSuccess = 0: mOther = 0: mPrompts = 0
    RibbonOnActionAdmin control
    selected = (Application.ActiveWorkbook Is mHost)
    If selected Then selected = (Application.ActiveSheet.Name = "WarehouseDirectory")
    Invoke = CStr(mCalls) & "|" & CStr(mOpened) & "|" & CStr(selected) & "|" & CStr(mSuccess) & "|" & CStr(mOther) & "|" & CStr(mPrompts)
End Function
Private Function Directory() As ListObject
    Set Directory = mHost.Worksheets("WarehouseDirectory").ListObjects("tblWarehouseDirectory")
End Function
Private Function ManagedSignature() As String
    Dim lo As ListObject, row As ListRow, field As Variant, value As String
    Set lo = Directory()
    For Each row In lo.ListRows
        If CStr(row.Range.Cells(1, lo.ListColumns("WarehouseId").Index).Value2) <> "" Then
            For Each field In Array("WarehouseId", "StationId", "WarehouseName", "RuntimeRoot", "ConfigWorkbook", "AuthWorkbook", "InventoryWorkbook", "ReceiveInbox", "ShipInbox", "ProdInbox", "ProcessorServiceUserId", "PathSharePointRoot", "RoleDefault", "Status")
                value = CStr(row.Range.Cells(1, lo.ListColumns(CStr(field)).Index).Value2)
                ManagedSignature = ManagedSignature & CStr(Len(value)) & ":" & value
            Next field
        End If
    Next row
End Function
Public Function AddCustomColumns() As Boolean
    Dim lo As ListObject, column As ListColumn, row As ListRow
    Set lo = Directory()
    mManaged = ManagedSignature()
    If mManaged = "" Then Exit Function
    Set column = lo.ListColumns.Add(2): column.Name = "Operator_note": mTextIndex = column.Index
    Set column = lo.ListColumns.Add(5): column.Name = "Operator_formula": mFormulaIndex = column.Index
    Set column = lo.ListColumns.Add: column.Name = "Operator_tail": mTailIndex = column.Index
    mTextAddress = lo.ListColumns("Operator_note").DataBodyRange.Cells(1, 1).Address
    mFormulaAddress = lo.ListColumns("Operator_formula").DataBodyRange.Cells(1, 1).Address
    lo.Parent.Range(mTextAddress).Value2 = "Keep this local note"
    lo.Parent.Range(mFormulaAddress).Formula = "=""Keep this formula"""
    Set row = lo.ListRows.Add
    mTailAddress = row.Range.Cells(1, mTailIndex).Address
    lo.Parent.Range(mTailAddress).Value2 = "Keep this local tail"
    AddCustomColumns = True
End Function
Public Function Preserved(ByVal part As String) As Boolean
    Dim lo As ListObject
    On Error GoTo Unavailable
    Set lo = Directory()
    Select Case part
        Case "Managed": Preserved = (ManagedSignature() = mManaged)
        Case "Headers": Preserved = (lo.ListColumns("Operator_note").Index = mTextIndex And lo.ListColumns("Operator_formula").Index = mFormulaIndex And lo.ListColumns("Operator_tail").Index = mTailIndex)
        Case "Text": Preserved = (lo.Parent.Range(mTextAddress).Value2 = "Keep this local note")
        Case "Formula": Preserved = (lo.Parent.Range(mFormulaAddress).Formula = "=""Keep this formula""")
        Case "Tail": Preserved = (lo.Parent.Range(mTailAddress).Value2 = "Keep this local tail")
        Case "Notes": Preserved = (mHost.Worksheets("Personal_notes").Range("A1").Value2 = "Preserve unrelated notes")
    End Select
Unavailable:
End Function
'@)
    $admin=$project.VBComponents.Item('modAdmin').CodeModule
    $admin.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestDirectoryColumns.ObserveMessage(Prompt)
End Function
Private Function InputBox(ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "") As String
    InputBox = TestDirectoryColumns.ObservePrompt()
End Function
'@)
    $newOwner=@($project.VBComponents|Where-Object Name -CEQ 'modAdminDirectoryActivity')
    $owner=if($newOwner.Count){$newOwner[0].CodeModule}else{$project.VBComponents.Item('modAdminConsole').CodeModule}
    $method=if($newOwner.Count){'OpenView'}else{'OpenWarehouseDirectory'}
    $resolve=if($newOwner.Count){'    Set wb = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(adminWb, ThisWorkbook, True)'}else{'    Set wb = ResolveAdminWorkbook(adminWb)'}
    if($newOwner.Count){$owner.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestDirectoryColumns.ObserveMessage(Prompt)
End Function
'@)}
    $start=$owner.ProcStartLine($method,0);$count=$owner.ProcCountLines($method,0)
    $source=$owner.Lines($start,$count)
    $changed=$source.Replace($resolve,"    TestDirectoryColumns.OwnerEntered`r`n"+$resolve).Replace(('    '+$method+' = True'),('    '+$method+" = True`r`n    TestDirectoryColumns.OwnerOpened"))
    if($source -ceq $changed){throw 'Directory owner observation seam changed.'}
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$changed)
    # Restrict discovery to generated files; retain the real directory reader/rebuild.
    $owner=$project.VBComponents.Item('modAdminConsole').CodeModule
    $start=$owner.ProcStartLine('GetWarehouseDirectoryScanRoots',0);$count=$owner.ProcCountLines('GetWarehouseDirectoryScanRoots',0)
    $owner.DeleteLines($start,$count)
    $owner.InsertLines($start,@'
Public Function GetWarehouseDirectoryScanRoots() As Collection
    Set GetWarehouseDirectoryScanRoots = TestDirectoryColumns.ReadRoots()
End Function
'@)
    if($AdminDirectoryRecording){Install-AdminDirectoryRecordingProbe}
    if($AdminWarehouseRootWriteDiagnostic){
        Install-AdminWarehouseRootWriteProbe
    }
}

function Test-AdminDirectoryColumns($Fixture,$Other) {
    function Invoke([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestDirectoryColumns.'+$Name) $Values}
    function Pins([string]$Root){$p=@{};Get-ChildItem $Root -Recurse -File|ForEach-Object {$s=[IO.File]::Open($_.FullName,'Open','Read','ReadWrite');try{$p[$_.FullName]=(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}};return $p}
    function Same($Before,$After){if($Before.Count -ne $After.Count){return $false};foreach($p in $Before.Keys){if(-not $After.ContainsKey($p) -or $Before[$p] -cne $After[$p]){return $false}};return $true}
    SelectTarget $Fixture
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add();$otherCount=$otherBook.Worksheets.Count
    $otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Other workbook canary'
    try{
        if(-not [bool](Invoke 'Prepare' @($hostBook.Name,$Fixture.Root))){throw 'Directory fixture preparation failed; not product RED.'}
        $before=Pins $Fixture.Root;$otherBefore=Pins $Other.Root
        $initial=[string](Invoke 'Invoke')
        Check 'Directory.Initial.ActualRibbonOwnerAndPresentation' ($initial -ceq '1|True|True|1|0|0')
        if($initial -cne '1|True|True|1|0|0'){throw 'Initial real directory result unavailable; not preservation RED.'}
        if(-not [bool](Invoke 'AddCustomColumns')){throw 'No generated directory rows; not preservation RED.'}
        foreach($part in @('Managed','Headers','Text','Formula','Tail','Notes')){Check ('Directory.Before.'+$part) ([bool](Invoke 'Preserved' @($part)))}
        Check 'Directory.Refresh.ActualRibbonOwnerAndPresentation' ([string](Invoke 'Invoke') -ceq '1|True|True|1|0|0')
        foreach($part in @('Managed','Headers','Text','Formula','Tail','Notes')){Check ('Directory.After.'+$part) ([bool](Invoke 'Preserved' @($part)))}
        $emptyRoot=Join-Path $runRoot 'empty-directory-scan';New-Item -ItemType Directory -Path $emptyRoot|Out-Null
        [void](Invoke 'SetRoot' @($emptyRoot))
        Check 'Directory.Empty.ActualRibbonOwnerAndPresentation' ([string](Invoke 'Invoke') -ceq '1|True|True|1|0|0')
        Check 'Directory.Empty.NoStaleManagedResultsAndExistingMessage' ([bool](Invoke 'NoResults'))
        foreach($part in @('Headers','Text','Formula','Tail','Notes')){Check ('Directory.Empty.'+$part) ([bool](Invoke 'Preserved' @($part)))}
        [void](Invoke 'SetRoot' @($Fixture.Root))
        Check 'Directory.Restored.ActualRibbonOwnerAndPresentation' ([string](Invoke 'Invoke') -ceq '1|True|True|1|0|0')
        Check 'Directory.Restored.Managed' ([bool](Invoke 'Preserved' @('Managed')))
        Check 'Directory.OtherWorkbookPreserved' ($otherBook.Worksheets.Count -eq $otherCount -and [string]$otherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
        Check 'Directory.SavedAuthorityPreserved' ((Same $before (Pins $Fixture.Root)) -and (Same $otherBefore (Pins $Other.Root)))
        if($Phase -ceq 'GREEN'){
            # Show only generated canaries; managed runtime paths remain hidden.
            $sheet=$hostBook.Worksheets.Item('WarehouseDirectory');$table=$sheet.ListObjects.Item('tblWarehouseDirectory')
            foreach($column in $table.ListColumns){if([string]$column.Name -notlike 'Operator_*'){$column.Range.EntireColumn.Hidden=$true}}
            $hostBook.Activate();$sheet.Activate();$excel.Visible=$true;$excel.WindowState=-4137
            $excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1;$excel.ActiveWindow.Zoom=100
            $table.ListColumns.Item('Operator_note').DataBodyRange.Cells.Item(1,1).Select()
            Initialize-SettingsCapture
            [void][InvSysSettingsCapture]::SetForegroundWindow([IntPtr]$excel.Hwnd)
            # False means the owned window was already foreground; input failures throw.
            [void][InvSysSettingsCapture]::ActivateByCaptionClick([IntPtr]$excel.Hwnd,[IntPtr]$excel.Hwnd)
            $captured=$false
            try{CaptureFormEvidence 'Directory local columns' 'directory-local-columns.png' ([long]$excel.Hwnd);$captured=$true}
            catch{
                if($_.Exception.GetBaseException().Message -cnotin @('Requested form content is obscured.','Requested form is not in the foreground.')){throw}
                [InvSysSettingsCapture]::CaptionFacts()|Set-Content (Join-Path $reportRoot 'directory-capture-window-facts.txt')
            }
            Check 'Directory.Visible.GeneratedLocalColumns' ($captured -and (Test-Path (Join-Path $reportRoot 'directory-local-columns.png')))
        }
        if($AdminWarehouseRootWriteDiagnostic){Test-AdminWarehouseRootWrite $Fixture $Other $hostBook $otherBook}
    }finally{$hostBook.Close($false);$otherBook.Close($false)}
    if($AdminDirectoryRecording){Test-AdminDirectoryRecording $Fixture $Other}
}
