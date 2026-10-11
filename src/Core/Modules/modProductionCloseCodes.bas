Attribute VB_Name = "modProductionCloseCodes"
Option Explicit
Option Private Module

' Catalog 21 observes dismissal only, never saved work or Domain application.
Public Function Control(ByVal id As String) As Object
    If id <> "PRODUCTION_CLOSE" Then Exit Function
    Set Control = modFormCloseCodes.Control(id)
End Function

Public Function Outcome(ByVal code As String) As Object
    Set Outcome = modFormCloseCodes.Outcome("PRODUCTION_CLOSE", code)
End Function
