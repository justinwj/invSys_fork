Attribute VB_Name = "modCreateWarehousePresentation"
Option Explicit

Public Function FormatSummaryMessage(ByVal messageText As String) As String
    Dim lines As Collection
    Dim parsed As String
    Dim lineValue As Variant

    messageText = Trim$(messageText)
    If messageText = "" Then Exit Function

    parsed = FormatBootstrapReportForSummary(messageText)
    If parsed <> "" Then
        FormatSummaryMessage = parsed
        Exit Function
    End If

    Set lines = New Collection
    AppendWrappedSummaryLine lines, messageText, 92

    For Each lineValue In lines
        If FormatSummaryMessage <> "" Then FormatSummaryMessage = FormatSummaryMessage & vbCrLf
        FormatSummaryMessage = FormatSummaryMessage & CStr(lineValue)
    Next lineValue
End Function

Private Function FormatBootstrapReportForSummary(ByVal reportText As String) As String
    Dim parts() As String
    Dim i As Long
    Dim keyText As String
    Dim valueText As String
    Dim lines As Collection
    Dim lineValue As Variant
    Dim eqPos As Long

    reportText = Trim$(reportText)
    If reportText = "" Then Exit Function
    If InStr(1, reportText, "|", vbBinaryCompare) = 0 Then Exit Function

    parts = Split(reportText, "|")
    Set lines = New Collection

    If UBound(parts) >= 0 Then
        If StrComp(Trim$(parts(0)), "OK", vbTextCompare) = 0 Then
            AppendWrappedSummaryLine lines, "Status: OK", 92
        Else
            AppendWrappedSummaryLine lines, Trim$(parts(0)), 92
        End If
    End If

    For i = 1 To UBound(parts)
        eqPos = InStr(1, parts(i), "=", vbBinaryCompare)
        If eqPos > 1 Then
            keyText = Trim$(Left$(parts(i), eqPos - 1))
            valueText = Trim$(Mid$(parts(i), eqPos + 1))
            If valueText <> "" Then
                valueText = Replace$(valueText, "COPIED:", "copied to ", 1, 1, vbTextCompare)
                valueText = Replace$(valueText, "SKIPPED", "already current", 1, 1, vbTextCompare)
                AppendWrappedSummaryLine lines, FriendlySummaryLabel(keyText) & ": " & valueText, 92
            End If
        ElseIf Trim$(parts(i)) <> "" Then
            AppendWrappedSummaryLine lines, Trim$(parts(i)), 92
        End If
    Next i

    For Each lineValue In lines
        If FormatBootstrapReportForSummary <> "" Then FormatBootstrapReportForSummary = FormatBootstrapReportForSummary & vbCrLf
        FormatBootstrapReportForSummary = FormatBootstrapReportForSummary & CStr(lineValue)
    Next lineValue
End Function

Private Function FriendlySummaryLabel(ByVal keyText As String) As String
    Select Case UCase$(Trim$(keyText))
        Case "CONFIG"
            FriendlySummaryLabel = "Config artifact"
        Case "DISCOVERY"
            FriendlySummaryLabel = "Discovery file"
        Case Else
            FriendlySummaryLabel = Trim$(keyText)
    End Select
End Function

Private Sub AppendWrappedSummaryLine(ByVal lines As Collection, ByVal textLine As String, ByVal maxChars As Long)
    Dim working As String
    Dim breakPos As Long

    working = Trim$(textLine)
    If working = "" Then
        lines.Add vbNullString
        Exit Sub
    End If

    Do While Len(working) > maxChars
        breakPos = InStrRev(Left$(working, maxChars), " ")
        If breakPos <= 0 Then breakPos = maxChars
        lines.Add Trim$(Left$(working, breakPos))
        working = Trim$(Mid$(working, breakPos + 1))
    Loop

    If working <> "" Then lines.Add working
End Sub
