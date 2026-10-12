Attribute VB_Name = "modProductionRunScaleActions"
Option Explicit
Option Private Module

' Keep captured-context failure authoritative and never scale after failed loading.
Public Function TryLoadWorksheetRecipe(ByVal recipeId As String, ByVal action As cProductionWorksheetAction, _
                                       ByRef report As String) As Boolean
    Dim loaded As Boolean
    loaded = mProduction.LoadRecipeChooser(recipeId)
    If Not action.CanContinue(report) Then Exit Function
    If Not loaded Then
        report = "Recipe could not be loaded. Scale was not applied."
        Exit Function
    End If
    TryLoadWorksheetRecipe = True
End Function

' Observe both Scale branches without substituting a successful owner result.
Public Function Execute(ByVal owner As frmProduction, ByVal context As String, _
                        ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean, _
                        ByVal scaleText As String, ByRef report As String) As Boolean
    Dim action As cProductionWorksheetAction, priorLoading As Boolean, loaded As Boolean, designs As Boolean
    Dim scalePercent As Double, number As Long, source As String, description As String
    Execute = True
    If loading Or busy Then Exit Function
    loaded = modProductionReusableRun.ReusableRunIsLoaded()
    designs = modConfig.GetBool("DesignsEnabled", False) Or modConfig.GetBool("FF_DesignsEnabled", False)
    On Error GoTo Failed
    priorLoading = loading: busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_SCALE", context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    action.OutcomeCode = "REJECTED"
    If Not loaded And designs Then
        report = "Load a released recipe before applying scale."
        GoTo Done
    End If
    If Not loaded And owner.SelectedRunRecipeText(0) = "" Then
        report = "Select a recipe, enter Batch scale %, then click Apply Scale."
        GoTo Done
    End If
    If Not owner.TryParseBatchScalePercent(scaleText, scalePercent, report) Then GoTo Done
    If Not loaded Then
        action.OutcomeCode = "FAILED": report = ""
        owner.ApplyWorksheetRunScale scalePercent, report, action
        GoTo Done
    End If
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
