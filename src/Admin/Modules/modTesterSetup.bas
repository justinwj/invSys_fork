Attribute VB_Name = "modTesterSetup"
Option Explicit

Public Type TesterSetupSpec
    UserId As String
    PinHash As String
    WarehouseId As String
    StationId As String
    PathLocal As String
    PathSharePointRoot As String
End Type

Public Function BrowseForSharePointRoot(Optional ByVal initialPath As String = "") As String
    Const FILE_DIALOG_FOLDER_PICKER As Long = 4
    Dim picker As Object

    On Error GoTo FailBrowse
    Set picker = Application.FileDialog(FILE_DIALOG_FOLDER_PICKER)
    If picker Is Nothing Then Exit Function

    picker.Title = "Select SharePoint Sync Root"
    picker.AllowMultiSelect = False
    If Trim$(initialPath) <> "" And modDeploymentPaths.FolderExistsManaged(initialPath) Then
        picker.InitialFileName = modConfig.NormalizeFolderPathForRuntime(initialPath, True)
    End If
    If picker.Show <> -1 Then Exit Function
    If picker.SelectedItems.Count = 0 Then Exit Function

    BrowseForSharePointRoot = modConfig.NormalizeFolderPathForRuntime(CStr(picker.SelectedItems(1)), False)
    Exit Function

FailBrowse:
    BrowseForSharePointRoot = vbNullString
End Function

Public Function OpenTesterReceivingWorkbook(Optional ByVal workbookPath As String = "") As Boolean
    Dim targetPath As String
    Dim wb As Workbook

    On Error GoTo FailOpen

    targetPath = Trim$(workbookPath)
    If targetPath = "" Then targetPath = modTesterProvisioning.GetLastTesterOperatorWorkbookPath()
    If targetPath = "" Then targetPath = modConfig.NormalizeFolderPathForRuntime( _
        modDeploymentPaths.DefaultWarehouseRuntimeRootPath("WH1", False), False) & "\WH1.Receiving.Operator.xlsm"
    If targetPath = "" Or Not modDeploymentPaths.FileExistsManaged(targetPath) Then Exit Function

    Dim candidate As Workbook
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, targetPath, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then Set wb = Application.Workbooks.Open(Filename:=targetPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    If wb Is Nothing Then Exit Function
    wb.Activate
    OpenTesterReceivingWorkbook = True
    Exit Function

FailOpen:
    LogDiagnosticEvent "TESTER-SETUP", "Operator workbook could not be opened."
End Function

Public Function SetupTesterStation(ByRef spec As TesterSetupSpec) As Boolean

    SetupTesterStation = modTesterProvisioning.SetupValues(spec.UserId, spec.PinHash, spec.WarehouseId, _
        spec.StationId, spec.PathLocal, spec.PathSharePointRoot)
End Function

Public Function VerifyReceivingWorkbook(ByVal pathIn As String, Optional ByRef report As String = "") As Boolean
    VerifyReceivingWorkbook = modTesterProvisioning.VerifyReceivingWorkbook(pathIn, report)
End Function

Public Function DetectSharePointRoot(Optional ByVal warehouseId As String = "WH1") As String
    DetectSharePointRoot = modTesterProvisioning.DetectSharePointRoot(warehouseId)
End Function

Public Sub SetTesterSharePointRootOverride(ByVal rootPath As String)
    modTesterProvisioning.SetTesterSharePointRootOverride rootPath
End Sub

Public Sub ClearTesterSharePointRootOverride()
    modTesterProvisioning.ClearTesterSharePointRootOverride
End Sub

Public Function GetLastTesterSetupReport() As String
    GetLastTesterSetupReport = modTesterProvisioning.GetLastTesterSetupReport()
End Function
