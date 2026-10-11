Attribute VB_Name = "modProductionRunScaleActions"
Option Explicit
Option Private Module

' Handle the released-run path; the worksheet branch retains its existing owner.
Public Function Execute(ByVal owner As frmProduction, ByVal context As String, _
                        ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean, _
                        ByVal scaleText As String, ByRef report As String) As Boolean
    Dim action As cProductionWorksheetAction, priorLoading As Boolean, loaded As Boolean
    Dim scalePercent As Double, number As Long, source As String, description As String
    Execute = True
    If loading Or busy Then Exit Function
    loaded = modProductionReusableRun.ReusableRunIsLoaded()
    If Not loaded Then
        If Not (modConfig.GetBool("DesignsEnabled", False) Or modConfig.GetBool("FF_DesignsEnabled", False)) Then
            Execute = False
            Exit Function
        End If
    End If
    On Error GoTo Failed
    priorLoading = loading: busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_SCALE", context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    action.OutcomeCode = "REJECTED"
    If Not loaded Then
        report = "Load a released recipe before applying scale."
        GoTo Done
    End If
    If Not owner.TryParseBatchScalePercent(scaleText, scalePercent, report) Then GoTo Done
    If Not modProductionReusableRun.ApplyReusableRunScale(scalePercent, report) Then GoTo Done
    action.OutcomeCode = "FAILED"
    If Not action.CanContinue(report) Then GoTo Done
    owner.RefreshReusableRunControls False, action
    If Not action.CanContinue(report) Then GoTo Done
    action.OutcomeCode = "STAGED"
Done:
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = False
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description
    End If
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    Resume Done
End Function
