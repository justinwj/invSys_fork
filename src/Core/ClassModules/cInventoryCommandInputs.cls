VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cInventoryCommandInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Transient Core-owned input binding. This object never crosses an XLAM boundary.
Private mWarehouse As String, mPath As String, mIdentity As String, mHash As String
Private mCatalog As cInventoryTableInput, mEntities As cInventoryTableInput
Private mStaging As cInventoryTableInput, mStage As ListObject
Private mSelectedSku As String, mPhase As String, mSaving As Boolean
Private mInitiallyBorrowed As Boolean

Public Property Get SourcePath() As String
    SourcePath = mPath
End Property

Public Function CaptureSelection(ByVal operatorBook As Workbook) As Boolean
    Dim selected As Range, candidate As ListObject
    On Error Resume Next
    Set selected = Application.Selection
    On Error GoTo Invalid
    If Not selected Is Nothing Then
        If selected.Parent.Parent Is operatorBook Then
            For Each candidate In selected.Parent.ListObjects
                If Not Application.Intersect(selected, candidate.Range) Is Nothing Then
                    If Left$(candidate.Name, 17) = "invSys_Inventory_" Then
                        Set mStage = candidate
                        Set mStaging = New cInventoryTableInput
                        If Not mStaging.Capture(candidate, "", True) Then Exit Function
                    End If
                    Exit For
                End If
            Next candidate
        End If
    End If
    CaptureSelection = True
Invalid:
End Function

Public Function CaptureSource(ByVal warehouse As String) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    mWarehouse = warehouse
    mPath = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouse)
    If mPath = "" Then Exit Function
    mIdentity = modAggregationFiles.Identity(mPath): mHash = modTrainingWire.FileSha256(mPath)
    If mIdentity = "" Or mHash = "" Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If borrowed And (book.ReadOnly Or Not book.Saved) Then Exit Function
    mInitiallyBorrowed = borrowed
    Set mCatalog = New cInventoryTableInput: Set mEntities = New cInventoryTableInput
    If Not mCatalog.Capture(modExecutionTarget.Table(book, "tblSkuCatalog"), "SKU", borrowed) Then Exit Function
    If Not mEntities.Capture(modExecutionTarget.Table(book, "tblInventoryEntities"), "System_Key", borrowed) Then Exit Function
    Set book = Nothing: Set owner = Nothing
    CaptureSource = Unchanged()
Invalid:
End Function

Public Function Unchanged() As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    If mCatalog Is Nothing Or mEntities Is Nothing Then Exit Function
    If Not mSaving Then
        If modAggregationFiles.Identity(mPath) <> mIdentity Or modTrainingWire.FileSha256(mPath) <> mHash Then Exit Function
    End If
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If borrowed And book.ReadOnly Then Exit Function
    If mPhase = "" And Not book.Saved Then Exit Function
    If Not mCatalog.Matches(modExecutionTarget.Table(book, "tblSkuCatalog")) Then Exit Function
    If Not mEntities.Matches(modExecutionTarget.Table(book, "tblInventoryEntities"), _
                            mSelectedSku, mSelectedSku = "") Then Exit Function
    If Not mStaging Is Nothing Then
        If Not mStaging.Matches(mStage) Then Exit Function
    End If
    Unchanged = True
Invalid:
End Function

Public Function BindSku(ByVal sku As String) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    If Not Unchanged() Or mPhase <> "" Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If sku <> "" Then
        If Not mEntities.Matches(modExecutionTarget.Table(book, "tblInventoryEntities"), sku) Then Exit Function
    End If
    mSelectedSku = sku: BindSku = True
Invalid:
End Function

Public Property Get StagingName() As String
    On Error GoTo Invalid
    If Not mStage Is Nothing Then StagingName = mStage.Name
Invalid:
End Property

Public Function PlanStageCell(ByVal row As Long, ByVal header As String, ByVal value As String) As Boolean
    If Not Unchanged() Or mStaging Is Nothing Then Exit Function
    Select Case UCase$(Trim$(header))
        Case "ITEM CODE", "UPLOAD STATUS", "UPLOAD RESULT"
            mStaging.PlanValue CStr(row), header, value
            PlanStageCell = True
    End Select
End Function

Public Function ConfirmStageCell() As Boolean
    If mStaging Is Nothing Then Exit Function
    If Not Unchanged() Then Exit Function
    ConfirmStageCell = mStaging.Confirm(mStage)
End Function

' The coordinator declares exact intended fields before the corresponding write.
' Partial values may match their old or declared value; unrelated changes refuse.
Public Function BeginCatalog(ByVal sku As String) As Boolean
    If Not Unchanged() Or mPhase <> "" Then Exit Function
    If sku = "" Then Exit Function
    If Not BindSku(sku) Then Exit Function
    mPhase = "Catalog": BeginCatalog = True
End Function

Public Sub PlanCatalogField(ByVal sku As String, ByVal header As String, ByVal value As Variant)
    If mPhase <> "Catalog" Or sku <> mSelectedSku Then Err.Raise 5, , "No matching catalog command."
    mCatalog.PlanValue sku, header, value
End Sub

Public Function BindCatalogWriter(ByVal book As Workbook) As Boolean
    On Error GoTo Invalid
    If mPhase <> "Catalog" Or Not Unchanged() Then Exit Function
    If Not modAggregationFiles.SameFile(book.FullName, mPath) Or book.ReadOnly Then Exit Function
    If Not book.Saved Then Exit Function
    If Not mCatalog.BindObject(modExecutionTarget.Table(book, "tblSkuCatalog")) Then Exit Function
    If Not mEntities.BindObject(modExecutionTarget.Table(book, "tblInventoryEntities")) Then Exit Function
    BindCatalogWriter = True
Invalid:
End Function

Public Function BeginSave() As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    If mPhase = "" Or Not Unchanged() Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If Not mCatalog.Matches(modExecutionTarget.Table(book, "tblSkuCatalog"), "", False, True) Then Exit Function
    mSaving = True: BeginSave = True
End Function

Public Function ConfirmCatalog() As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    If mPhase <> "Catalog" Or Not mSaving Then Exit Function
    If Not Unchanged() Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If Not book.Saved Then Exit Function
    If Not mEntities.Matches(modExecutionTarget.Table(book, "tblInventoryEntities"), mSelectedSku) Then Exit Function
    If Not mCatalog.Confirm(modExecutionTarget.Table(book, "tblSkuCatalog")) Then Exit Function
    If Not mInitiallyBorrowed Then
        mCatalog.ReleaseObjectReference
        mEntities.ReleaseObjectReference
    End If
    mIdentity = modAggregationFiles.Identity(mPath): mHash = modTrainingWire.FileSha256(mPath)
    If mIdentity = "" Or mHash = "" Then Exit Function
    mPhase = "": mSaving = False
    ConfirmCatalog = Unchanged()
Invalid:
End Function

Private Function ReadSource(ByRef owner As cSnapshotInventorySource, ByRef borrowed As Boolean) As Workbook
    Dim candidate As Workbook, book As Workbook
    For Each candidate In Application.Workbooks
        If modAggregationFiles.SameFile(candidate.FullName, mPath) Then borrowed = True: Exit For
    Next candidate
    Set owner = New cSnapshotInventorySource
    Set book = owner.Read(mWarehouse, Nothing)
    If book Is Nothing Then Exit Function
    If Not modAggregationFiles.SameFile(book.FullName, mPath) Then Exit Function
    Set ReadSource = book
End Function
