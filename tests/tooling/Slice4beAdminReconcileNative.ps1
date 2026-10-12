# Actual generated-fixture native selection/cancel evidence. Never logs row values.
# Screenshots contain only disposable fixture data and remain ignored runtime files.
function Install-AdminReconcileNativeProbe($Module) {
    $Module.InsertLines($Module.CountOfDeclarationLines+1,'Private mNativeLastChoiceEvent As String')
    $Module.AddFromString(@'
Public Function PrepareNativeHistory() As Boolean
    Dim item As Object, eventId As String, i As Long
    If EntityQty(mKeyA) <> 21 Or EntityQty(mKeyB) <> 55 Then Exit Function
    For i = 1 To 20
        Set item = HistoryItem(mKeyA, 1, "CURRENT-A", "DAMAGED", "SHIP")
        eventId = modRoleEventWriter.CreateSystemKey()
        If i = 1 Then mNativeLastChoiceEvent = eventId
        If Not ApplyFixtureEvent("SHIP", eventId, item) Then Exit Function
    Next i
    PrepareNativeHistory = (EntityQty(mKeyA) = 1 And EntityQty(mKeyB) = 55 And StageCanary())
End Function
Public Function NativeLastChoiceEvent() As String
    NativeLastChoiceEvent = mNativeLastChoiceEvent
End Function
Public Function InvokeNativeSelection(Optional ByVal inspectLast As Boolean = False) As Boolean
    On Error GoTo Failed
    mNativePresentation = True
    InvokeCallback
    mNativePresentation = False
    InvokeNativeSelection = (mPrompts = 1 And Not mQueuedNotice And mBoundaryNotice = "")
    If inspectLast Then InvokeNativeSelection = (mPrompts = 1 And Not mQueuedNotice And mBoundaryNotice = "Current quantity is not higher than the selected entity's balance after that shipment. No correction was queued.")
    Exit Function
Failed:
    mNativePresentation = False
    Err.Raise 5, , "Native Reconcile callback failed"
End Function
Public Function NativeFixturePreserved() As Boolean
    NativeFixturePreserved = (EntityQty(mKeyA) = 1 And EntityQty(mKeyB) = 55 And mInventory.Saved And CanaryPreserved())
End Function
'@)
}

function Invoke-ReconcileNativeCancel([string]$LastEvent,[string]$SystemKey) {
    if($script:reconcileSelectionFormPresent){return Invoke-ReconcileSelectionNative}
    $handle=[long]$excel.Hwnd
    $ready=Join-Path $runRoot ('reconcile-native-'+[guid]::NewGuid().ToString('N')+'.ready')
    $capture=Join-Path $reportRoot 'reconcile-native-selection.png'
    $job=Start-Job -ArgumentList $handle,$ready,$capture,$LastEvent,$SystemKey -ScriptBlock {
        param($handle,$ready,$capture,$lastEvent,$systemKey)
        $ErrorActionPreference='Stop'
        Add-Type -ReferencedAssemblies System.Drawing @'
using System;using System.Text;using System.Text.RegularExpressions;using System.Collections.Generic;using System.Drawing;using System.Runtime.InteropServices;
public static class ReconcileNativeDialog {
 public delegate bool CB(IntPtr h,IntPtr p);
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [DllImport("user32.dll")]static extern bool EnumWindows(CB c,IntPtr p);
 [DllImport("user32.dll")]static extern bool EnumChildWindows(IntPtr h,CB c,IntPtr p);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")]static extern bool GetWindowRect(IntPtr h,out R r);
 [DllImport("user32.dll")]static extern bool GetClientRect(IntPtr h,out R r);
 [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
 [DllImport("user32.dll",SetLastError=true)]static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
 [DllImport("user32.dll")]static extern IntPtr GetDC(IntPtr h);
 [DllImport("user32.dll")]static extern int ReleaseDC(IntPtr h,IntPtr dc);
 [DllImport("gdi32.dll")]static extern IntPtr SelectObject(IntPtr dc,IntPtr o);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int DrawText(IntPtr dc,string s,int n,ref R r,uint f);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]static extern int GetSystemMetrics(int i);
 [DllImport("user32.dll")]static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetCursorPos(out P p);
 public static uint Owner(IntPtr h){uint p;GetWindowThreadProcessId(h,out p);return p;}
 static string Text(IntPtr h){var b=new StringBuilder(32768);GetWindowText(h,b,b.Capacity);return b.ToString();}
 static string Class(IntPtr h){var b=new StringBuilder(128);GetClassName(h,b,b.Capacity);return b.ToString();}
 public static IntPtr Find(uint owner){IntPtr found=IntPtr.Zero;EnumWindows((h,p)=>{if(Owner(h)==owner&&IsWindowVisible(h)&&Text(h)=="invSys Admin - Shipment Reconcile")found=h;return found==IntPtr.Zero;},IntPtr.Zero);return found;}
 static IntPtr Child(IntPtr h,string kind,string starts){IntPtr found=IntPtr.Zero;EnumChildWindows(h,(c,p)=>{if(IsWindowVisible(c)&&Class(c)==kind&&Text(c).Replace("&","").StartsWith(starts))found=c;return true;},IntPtr.Zero);return found;}
 public class Result{public bool ExactDialog,LastChoicePresent,ExactKeyPresent,HumanReadableTimestamp,TextFits,WindowWithinDisplay,Captured,CancelDispatched,ForegroundBeforeCapture,ForegroundAfterCapture,ForegroundRequestAccepted;public int PromptCharacters,DesktopError;}
 public static Result Inspect(IntPtr h,uint owner,string lastEvent,string key,string path){
  var result=new Result();IntPtr previous=SetThreadDpiAwarenessContext(new IntPtr(-4));
  if(previous==IntPtr.Zero)throw new Exception("Physical pixel context unavailable.");
  try{
   var cancel=Child(h,"Button","Cancel");var prompt=Child(h,"Static","Enter the number of the server shipment deduction log line to inspect.");
   result.ExactDialog=Owner(h)==owner&&Find(owner)==h&&cancel!=IntPtr.Zero&&prompt!=IntPtr.Zero;
   if(!result.ExactDialog)throw new Exception("Expected native selection controls unavailable.");
   try{
    P cursor;if(!GetCursorPos(out cursor)){result.DesktopError=Marshal.GetLastWin32Error();return result;}
    string text=Text(prompt);result.PromptCharacters=text.Length;
    result.LastChoicePresent=text.Contains("20. "+lastEvent);result.ExactKeyPresent=text.Contains("System_Key: "+key);
    result.HumanReadableTimestamp=Regex.IsMatch(text,@"\b20\d{2}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b");
    R bounds,client;GetWindowRect(h,out bounds);GetClientRect(prompt,out client);
    var dc=GetDC(prompt);var font=SendMessage(prompt,0x31,IntPtr.Zero,IntPtr.Zero);var old=SelectObject(dc,font);
    try{R measured=new R{L=0,T=0,Right=client.Right,Bottom=0};int height=DrawText(dc,text,text.Length,ref measured,0x400|0x10|0x800);result.TextFits=height>0&&client.Right>0&&measured.Bottom<=client.Bottom;}
    finally{SelectObject(dc,old);ReleaseDC(prompt,dc);}
    int left=GetSystemMetrics(76),top=GetSystemMetrics(77),right=left+GetSystemMetrics(78),bottom=top+GetSystemMetrics(79);
    result.WindowWithinDisplay=bounds.L>=left&&bounds.T>=top&&bounds.Right<=right&&bounds.Bottom<=bottom;
    int x=Math.Max(left,bounds.L),y=Math.Max(top,bounds.T),w=Math.Min(right,bounds.Right)-x,z=Math.Min(bottom,bounds.Bottom)-y;
    result.ForegroundBeforeCapture=GetForegroundWindow()==h;
    if(result.ForegroundBeforeCapture&&w>0&&z>0){using(var b=new Bitmap(w,z)){using(var g=Graphics.FromImage(b)){g.CopyFromScreen(x,y,0,0,b.Size);}result.ForegroundAfterCapture=GetForegroundWindow()==h;if(result.ForegroundAfterCapture&&Find(owner)==h){b.Save(path,System.Drawing.Imaging.ImageFormat.Png);result.Captured=true;}}}
   }finally{result.CancelDispatched=PostMessage(cancel,0xF5,IntPtr.Zero,IntPtr.Zero);}
   return result;
  }finally{SetThreadDpiAwarenessContext(previous);}
 }
}
'@
        $owner=[ReconcileNativeDialog]::Owner([IntPtr]$handle)
        if($owner -eq 0){throw 'Generated Excel owner unavailable.'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($ready,'Ready')
        $deadline=[DateTime]::UtcNow.AddSeconds(60)
        while([DateTime]::UtcNow -lt $deadline){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Generated Excel identity changed.'}
            $dialog=[ReconcileNativeDialog]::Find($owner)
            if($dialog -ne [IntPtr]::Zero){
                $foregroundAccepted=[ReconcileNativeDialog]::SetForegroundWindow($dialog)
                Start-Sleep -Milliseconds 250
                $result=[ReconcileNativeDialog]::Inspect($dialog,$owner,$lastEvent,$systemKey,$capture)
                $result.ForegroundRequestAccepted=$foregroundAccepted
                return $result
            }
            Start-Sleep -Milliseconds 100
        }
        throw 'Native Reconcile selection did not appear.'
    }
    try{
        for($i=0;$i -lt 100 -and -not(Test-Path -LiteralPath $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path -LiteralPath $ready)){throw 'Native observer not ready.'}
        $returned=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeNativeSelection')
        [void](Wait-Job $job -Timeout 5)
        $observed=Receive-Job $job -ErrorAction Stop
        if($null -eq $observed){throw 'Native result unavailable.'}
        $observed|Select-Object ExactDialog,LastChoicePresent,ExactKeyPresent,HumanReadableTimestamp,TextFits,WindowWithinDisplay,Captured,CancelDispatched,ForegroundBeforeCapture,ForegroundAfterCapture,ForegroundRequestAccepted,PromptCharacters,DesktopError|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'native-selection.json')
        if($observed.DesktopError -ne 0){throw ('Desktop Win32 error '+$observed.DesktopError+' at '+[DateTimeOffset]::UtcNow.ToString('o'))}
        return [pscustomobject]@{Returned=$returned;Observed=$observed}
    }finally{
        if($job.State -eq 'Running'){Stop-Job $job}
        Remove-Job $job
    }
}

function Test-AdminReconcileNative($Events,$State) {
    # Owner-generated data windows are hidden by the normal services. Keep an
    # unsaved, disposable operator window visible during native form interaction.
    $nativeHost=$excel.Workbooks.Add()
    try{
    $nativeHost.Windows.Item(1).Visible=$true
    $excel.Visible=$true
    $nativeHost.Activate()
    $ready=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareNativeHistory')
    Check 'Reconcile.Native.TwentyOwnerAppliedShipments' $ready
    if(-not $ready){throw 'Native history fixture failed; not behavioral RED.'}
    $last=[string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeLastChoiceEvent')
    $bytes=Hash $State.InventoryPath;$rows=$Events.ListRows.Count
    $visible=[bool]$excel.Visible -and [bool]$nativeHost.Windows.Item(1).Visible
    Check 'Reconcile.Native.VisibleExcelHost' $visible
    if(-not $visible){throw 'Visible operator host prerequisite failed; native entry withheld.'}
    $result=Invoke-ReconcileNativeCancel $last $State.KeyA
    Check 'Reconcile.Native.ActualSelectionControlsObserved' ([bool]$result.Observed.ExactDialog)
    Check 'Reconcile.Native.LastChoiceAndExactKeyPresent' ($result.Observed.LastChoicePresent -and $result.Observed.ExactKeyPresent)
    Check 'Reconcile.Native.SelectionTextFitsVisibleDialog' ($result.Observed.TextFits -and $result.Observed.WindowWithinDisplay)
    Check 'Reconcile.Native.HistoryTimestampReadable' ([bool]$result.Observed.HumanReadableTimestamp)
    Check 'Reconcile.Native.OperatorCapture' ([bool]$result.Observed.Captured)
    Check 'Reconcile.Native.CancelReturnsWithoutQueue' ($result.Returned -and $result.Observed.CancelDispatched -and $Events.ListRows.Count -eq $rows)
    Check 'Reconcile.Native.InventoryBytesPreserved' ((Hash $State.InventoryPath) -ceq $bytes)
    Check 'Reconcile.Native.QuantitiesCustomValuesAndSavedStatePreserved' ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved'))
    if($script:reconcileSelectionFormPresent){
        Check 'Reconcile.Native.Layout.AllRequiredSizesObserved' ((@($result.Layouts.Mode) -join '|') -ceq 'Grow|Minimum|Maximize|Restore')
        foreach($layout in $result.Layouts){
            Check ('Reconcile.Native.Layout.'+$layout.Mode) ($layout.Geometry -and $layout.CompleteHistory -and $layout.Scrollable -and $layout.Captured -and $layout.WithinDisplay)
        }
        [void](Run 'invSys.Inventory.Domain.xlam' 'TestReconcileReadLease.Reset')
        $inspected=Invoke-ReconcileSelectionNative -InspectLast
        Check 'Reconcile.Native.InspectLast.ActualEntryAndHandler' ($inspected.Observed.InspectDispatched -and $inspected.Observed.SelectionIs20)
        Check 'Reconcile.Native.InspectLast.ExactEventAndKeyRead' ([bool](Run 'invSys.Inventory.Domain.xlam' 'TestReconcileReadLease.ExactStateRequest' @($last,$State.KeyA)))
        Check 'Reconcile.Native.InspectLast.ConclusionWithoutQueue' ($inspected.Returned -and $Events.ListRows.Count -eq $rows)
        Check 'Reconcile.Native.InspectLast.InventoryPreserved' (((Hash $State.InventoryPath) -ceq $bytes) -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved'))
    }
    }finally{try{$nativeHost.Close($false)}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($nativeHost)}
}
