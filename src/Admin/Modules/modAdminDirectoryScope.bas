Attribute VB_Name = "modAdminDirectoryScope"
Option Explicit
Option Private Module

' Ambient discovery inputs are frozen before optional recording; preference roots
' are read after the ordinary owner write, preserving their existing ordering.
Public Function Capture(ByVal adminRoot As String, ByVal configRoot As String, _
                        ByVal overrideRoot As String, ByVal knownRoots As Collection, _
                        ByVal defaultRoot As String) As Object
    Dim result As Object, roots As New Collection, item As Variant
    Set result = CreateObject("Scripting.Dictionary")
    roots.Add adminRoot: roots.Add configRoot: roots.Add overrideRoot
    For Each item In knownRoots: roots.Add CStr(item): Next item
    result.Add "AmbientRoots", roots: result.Add "DefaultRoot", defaultRoot
    Set Capture = result
End Function

Public Function Roots(ByVal inputs As Object, ByVal remembered As Collection) As Collection
    Dim result As New Collection, item As Variant
    For Each item In inputs("AmbientRoots")
        modAdminConsole.AddWarehouseDirectoryScanRoot result, CStr(item)
    Next item
    For Each item In remembered
        modAdminConsole.AddWarehouseDirectoryScanRoot result, CStr(item)
    Next item
    modAdminConsole.AddWarehouseDirectoryScanRoot result, CStr(inputs("DefaultRoot"))
    Set Roots = result
End Function
