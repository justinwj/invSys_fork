# Measure in Excel: GDI font handles belong to the process that created them.
# The bounded timer only observes the generated confirmation; it never sends input.
function Install-AdminReconcileFontProbe($Module) {
    $Module.InsertLines($Module.CountOfDeclarationLines+1,@'
Private Declare PtrSafe Function RfSetTimer Lib "user32" Alias "SetTimer" (ByVal h As LongPtr, ByVal id As LongPtr, ByVal ms As Long, ByVal proc As LongPtr) As LongPtr
Private Declare PtrSafe Function RfKillTimer Lib "user32" Alias "KillTimer" (ByVal h As LongPtr, ByVal id As LongPtr) As Long
Private Declare PtrSafe Function RfForeground Lib "user32" Alias "GetForegroundWindow" () As LongPtr
Private Declare PtrSafe Function RfText Lib "user32" Alias "GetWindowTextW" (ByVal h As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function RfChild Lib "user32" Alias "FindWindowExW" (ByVal h As LongPtr, ByVal after As LongPtr, ByVal kind As LongPtr, ByVal caption As LongPtr) As LongPtr
Private Declare PtrSafe Function RfMessage Lib "user32" Alias "SendMessageW" (ByVal h As LongPtr, ByVal msg As Long, ByVal w As LongPtr, ByVal l As LongPtr) As LongPtr
Private Declare PtrSafe Function RfDC Lib "user32" Alias "GetDC" (ByVal h As LongPtr) As LongPtr
Private Declare PtrSafe Function RfRelease Lib "user32" Alias "ReleaseDC" (ByVal h As LongPtr, ByVal dc As LongPtr) As Long
Private Declare PtrSafe Function RfSelect Lib "gdi32" Alias "SelectObject" (ByVal dc As LongPtr, ByVal obj As LongPtr) As LongPtr
Private Declare PtrSafe Function RfStock Lib "gdi32" Alias "GetStockObject" (ByVal index As Long) As LongPtr
Private Declare PtrSafe Function RfDpi Lib "user32" Alias "SetThreadDpiAwarenessContext" (ByVal context As LongPtr) As LongPtr
Private Type RfRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare PtrSafe Function RfClient Lib "user32" Alias "GetClientRect" (ByVal h As LongPtr, ByRef rect As RfRect) As Long
Private Declare PtrSafe Function RfDraw Lib "user32" Alias "DrawTextW" (ByVal dc As LongPtr, ByVal text As LongPtr, ByVal count As Long, ByRef rect As RfRect, ByVal flags As Long) As Long
Private mRfTimer As LongPtr, mRfTries As Long, mRfExpected As String, mRfProof As String
'@)
    $Module.AddFromString(@'
Private Function RfWindowText(ByVal h As LongPtr) As String
    Dim value As String, count As Long
    value = String$(4096, vbNullChar)
    count = RfText(h, StrPtr(value), Len(value))
    RfWindowText = Left$(value, count)
End Function
Public Sub StartNativeFontProbe(ByVal prompt As String)
    StopNativeFontProbe
    mRfExpected = prompt: mRfProof = "": mRfTries = 0
    mRfTimer = RfSetTimer(0, 0, 50, AddressOf NativeFontTick)
    If mRfTimer = 0 Then Err.Raise 5, , "Native font observer unavailable"
End Sub
Public Sub StopNativeFontProbe()
    If mRfTimer <> 0 Then RfKillTimer 0, mRfTimer
    mRfTimer = 0
End Sub
Public Function NativeFontProof() As String
    NativeFontProof = mRfProof
End Function
Public Sub NativeFontTick(ByVal hwnd As LongPtr, ByVal msg As Long, ByVal id As LongPtr, ByVal tick As Long)
    On Error GoTo Done
    Dim dialog As LongPtr, prompt As LongPtr, dc As LongPtr, font As LongPtr, previous As LongPtr, dpi As LongPtr
    Dim client As RfRect, measured As RfRect, height As Long, kind As String, value As String
    mRfTries = mRfTries + 1
    If mRfTries > 100 Then GoTo Done
    dialog = RfForeground()
    If RfWindowText(dialog) <> "invSys Admin - Confirm Shipment Reconcile" Then Exit Sub
    kind = "Static"
    Do
        prompt = RfChild(dialog, prompt, StrPtr(kind), 0)
        If prompt = 0 Then Exit Sub
        value = RfWindowText(prompt)
    Loop Until value = mRfExpected
    dpi = RfDpi(-4)
    If dpi = 0 Then GoTo Done
    If RfClient(prompt, client) = 0 Then GoTo Done
    dc = RfDC(prompt)
    If dc = 0 Then GoTo Done
    font = RfMessage(prompt, &H31, 0, 0)
    If font = 0 Then font = RfStock(13)
    previous = RfSelect(dc, font)
    If previous = 0 Then GoTo Done
    measured.Right = client.Right
    height = RfDraw(dc, StrPtr(value), Len(value), measured, &H400 Or &H10 Or &H800)
    If height > 0 And client.Right > 0 Then mRfProof = CStr(Len(value)) & "|" & CStr(client.Right) & "|" & CStr(client.Bottom) & "|" & CStr(measured.Bottom)
Done:
    On Error Resume Next
    If previous <> 0 Then RfSelect dc, previous
    If dc <> 0 Then RfRelease prompt, dc
    If dpi <> 0 Then RfDpi dpi
    StopNativeFontProbe
End Sub
'@)
}
