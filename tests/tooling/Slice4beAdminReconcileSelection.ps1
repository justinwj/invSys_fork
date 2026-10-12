# Disposable instrumentation: deterministic entries use the same Inspect/Cancel
# handlers; native mode leaves the actual modal form and keyboard input intact.
function Install-ReconcileSelectionProbe($Project,$Probe,$Owner) {
    $forms=@($Project.VBComponents|Where-Object Name -CEQ 'frmShipmentReconcileSelection')
    $script:reconcileSelectionFormPresent=($forms.Count -eq 1)
    if($forms.Count -eq 0){return}
    $form=$forms[0].CodeModule
    $form.AddFromString(@'
Public Sub SelectionActionForTest(ByVal entry As String)
    mSelection.Value = entry
    If entry = "" Then mCancel_Click Else mInspect_Click
End Sub
Public Function GeometryForTest() As Boolean
    Dim a As Object, b As Object
    For Each a In Me.Controls
        If a.Left < 0 Or a.Top < 0 Or a.Width <= 0 Or a.Height <= 0 Then Exit Function
        If a.Left + a.Width > Me.InsideWidth Or a.Top + a.Height > Me.InsideHeight Then Exit Function
        For Each b In Me.Controls
            If a.Name <> b.Name Then
                If a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And a.Top < b.Top + b.Height And b.Top < a.Top + a.Height Then Exit Function
            End If
        Next b
    Next a
    GeometryForTest = True
End Function
'@)
    $Probe.AddFromString(@'
Public Function PresentSelectionForTest(ByVal dialog As Object, ByVal prompt As String) As Boolean
    Dim entry As String
    If mNativePresentation Then
        mPrompts = mPrompts + 1: mListPrompt = prompt
    Else
        entry = PromptAnswer("invSys Admin - Shipment Reconcile", prompt)
        dialog.SelectionActionForTest entry
        PresentSelectionForTest = True
    End If
End Function
'@)
    $start=$Owner.ProcStartLine('SelectShipmentHistory',0)
    $count=$Owner.ProcCountLines('SelectShipmentHistory',0)
    $found=0
    for($line=$start;$line -lt $start+$count;$line++){
        if($Owner.Lines($line,1).Trim() -ceq 'dialog.Show vbModal'){
            $Owner.ReplaceLine($line,'    If Not TestReconcileIdentity.PresentSelectionForTest(dialog, prompt) Then dialog.Show vbModal')
            $found++
        }
    }
    if($found -ne 1){throw 'Exactly one selection presentation boundary required.'}
    if($AdminReconcileNativePresentation){Install-ReconcileSelectionNativeProbe $Probe $form}
}

function Install-ReconcileSelectionNativeProbe($Probe,$Form) {
    $Probe.InsertLines($Probe.CountOfDeclarationLines+1,'Private mFormReadyPath As String, mFormClosedPath As String, mSelectionShown As Boolean, mSelectionRevision As Long')
    $Probe.AddFromString(@'
Public Sub SetSelectionProbePaths(ByVal readyPath As String, ByVal closedPath As String)
    mFormReadyPath = readyPath: mFormClosedPath = closedPath
    mSelectionShown = False: mSelectionRevision = 0
End Sub
Public Sub ObserveSelectionShown(ByVal dialog As Object)
    If Not mNativePresentation Then Exit Sub
    dialog.Controls("txtHistory").SetFocus
    mSelectionShown = True
    WriteSelectionFacts dialog, mFormReadyPath
End Sub
Public Sub ObserveSelectionLayout(ByVal dialog As Object)
    If mNativePresentation And mSelectionShown Then WriteSelectionFacts dialog, mFormReadyPath
End Sub
Public Sub ObserveSelectionClosed(ByVal dialog As Object)
    If mNativePresentation Then WriteSelectionFacts dialog, mFormClosedPath
End Sub
Private Sub WriteSelectionFacts(ByVal dialog As Object, ByVal path As String)
    Dim history As Object, text As String, complete As Boolean, atEnd As Boolean, scrollable As Boolean, f As Integer
    Dim observedPosition As Long, endPosition As Long
    Set history = dialog.Controls("txtHistory"): text = CStr(history.Value)
    complete = (InStr(1, text, "20. " & mNativeLastChoiceEvent & " | System_Key: " & mKeyA, vbBinaryCompare) > 0)
    observedPosition = history.SelStart: endPosition = -1
    If path = mFormClosedPath Then
        ' Calibrate this control's caret indexing only after native capture/input.
        ' Preserve the observed position and compare it to the clamped end.
        history.SelStart = Len(text): endPosition = history.SelStart
        history.SelStart = observedPosition
    End If
    atEnd = (observedPosition = endPosition And Len(text) > 1023)
    scrollable = (history.Locked And history.MultiLine And history.WordWrap And history.ScrollBars = 2)
    mSelectionRevision = mSelectionRevision + 1
    f = FreeFile
    Open path For Output As #f
    Print #f, "{""CompleteHistory"":" & LCase$(CStr(complete)) & ",""AtEnd"":" & LCase$(CStr(atEnd)) & _
        ",""Scrollable"":" & LCase$(CStr(scrollable)) & ",""Geometry"":" & LCase$(CStr(dialog.GeometryForTest())) & _
        ",""Timestamp"":" & LCase$(CStr(text Like "*20##-##-## ##:##:##*")) & ",""Characters"":" & CStr(Len(text)) & _
        ",""Window"":" & CStr(modUserFormResizeWin.GetUserFormWindowHandle(dialog)) & _
        ",""ObservedPosition"":" & CStr(observedPosition) & ",""EndPosition"":" & CStr(endPosition) & _
        ",""Width"":" & CStr(CLng(dialog.Width)) & ",""Height"":" & CStr(CLng(dialog.Height)) & ",""Revision"":" & CStr(mSelectionRevision) & _
        ",""SelectionIs20"":" & LCase$(CStr(CStr(dialog.Controls("txtSelection").Value) = "20")) & "}"
    Close #f
End Sub
'@)
    $line=$Form.ProcStartLine('UserForm_Activate',0)+$Form.ProcCountLines('UserForm_Activate',0)-1
    while($Form.Lines($line,1).Trim() -cne 'End Sub'){$line--}
    $Form.InsertLines($line,'    TestReconcileIdentity.ObserveSelectionShown Me')
    $line=$Form.ProcStartLine('UserForm_Layout',0)+$Form.ProcCountLines('UserForm_Layout',0)-1
    while($Form.Lines($line,1).Trim() -cne 'End Sub'){$line--}
    $Form.InsertLines($line,'    TestReconcileIdentity.ObserveSelectionLayout Me')
    foreach($handler in @('mCancel_Click','mInspect_Click')){
        $line=$Form.ProcBodyLine($handler,0)
        $Form.InsertLines($line+1,'    TestReconcileIdentity.ObserveSelectionClosed Me')
    }
}

function Invoke-ReconcileSelectionNative([switch]$InspectLast) {
    $prefix=Join-Path $reportRoot ('selection-'+[guid]::NewGuid().ToString('N'))
    $ready=$prefix+'-form.json';$closed=$prefix+'-closed.json';$jobReady=$prefix+'-job.ready'
    $capture=Join-Path $reportRoot $(if($InspectLast){'reconcile-selection-inspect.png'}else{'reconcile-selection-last.png'})
    [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.SetSelectionProbePaths' @($ready,$closed))
    $job=Start-Job -ArgumentList ([long]$excel.Hwnd),$ready,$jobReady,$capture,([bool]$InspectLast) -ScriptBlock {
        param($handle,$ready,$jobReady,$capture,[bool]$inspectLast)
        $ErrorActionPreference='Stop'
        function Stage([string]$name){[IO.File]::AppendAllText($jobReady+'.stages.jsonl',('{"UTC":"'+[DateTimeOffset]::UtcNow.ToString('o')+'","Stage":"'+$name+'"}'+[Environment]::NewLine))}
        Add-Type -ReferencedAssemblies System.Drawing @'
using System;using System.Text;using System.Drawing;using System.Runtime.InteropServices;
public static class ReconcileSelectionInput {
 public delegate bool CB(IntPtr h,IntPtr p);
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [StructLayout(LayoutKind.Sequential)]public struct M{public int X,Y;public uint Data,Flags,Time;public UIntPtr Extra;}
 [StructLayout(LayoutKind.Sequential)]public struct K{public ushort Key,Scan;public uint Flags,Time;public UIntPtr Extra;}
 [StructLayout(LayoutKind.Explicit)]public struct U{[FieldOffset(0)]public M Mouse;[FieldOffset(0)]public K Key;}
 [StructLayout(LayoutKind.Sequential)]public struct I{public uint Type;public U Data;}
 [DllImport("user32.dll")]static extern bool EnumWindows(CB cb,IntPtr p);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]static extern bool GetWindowRect(IntPtr h,out R r);
 [DllImport("user32.dll")]static extern bool SetWindowPos(IntPtr h,IntPtr after,int x,int y,int w,int z,uint flags);
 [DllImport("user32.dll")]static extern bool ShowWindow(IntPtr h,int state);
 [DllImport("dwmapi.dll")]static extern int DwmGetWindowAttribute(IntPtr h,int attr,out R r,int size);
 [DllImport("user32.dll")]static extern int GetSystemMetrics(int i);
 [DllImport("user32.dll")]static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetCursorPos(out P p);
 [DllImport("user32.dll",SetLastError=true)]static extern uint SendInput(uint n,I[] input,int size);
 [DllImport("user32.dll",SetLastError=true)]static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
 public static uint Owner(IntPtr h){uint p;GetWindowThreadProcessId(h,out p);return p;}
 public static IntPtr Find(uint owner){IntPtr found=IntPtr.Zero;EnumWindows((h,p)=>{var title=new StringBuilder(256);GetWindowText(h,title,256);if(Owner(h)==owner&&IsWindowVisible(h)&&title.ToString()=="invSys Admin - Shipment Reconcile")found=h;return found==IntPtr.Zero;},IntPtr.Zero);return found;}
 static void Guard(IntPtr h,uint owner){if(Find(owner)!=h||GetForegroundWindow()!=h)throw new Exception("Owned selection form is not foreground.");}
 static I Key(ushort key,bool up){return new I{Type=1,Data=new U{Key=new K{Key=key,Flags=up?2u:0u}}};}
 static void Keys(I[] keys){uint n=SendInput((uint)keys.Length,keys,Marshal.SizeOf(typeof(I)));int e=n==keys.Length?0:Marshal.GetLastWin32Error();if(n!=keys.Length)throw new Exception("Keyboard delivery failed; Win32 error "+e);}
 public static void ScrollEnd(IntPtr h,uint owner){Guard(h,owner);try{Keys(new[]{Key(17,false),Key(36,false),Key(36,true),Key(35,false),Key(35,true),Key(17,true)});}finally{Keys(new[]{Key(36,true),Key(35,true),Key(17,true)});}}
 public static bool Cancel(IntPtr h,uint owner){Guard(h,owner);Keys(new[]{Key(27,false),Key(27,true)});return true;}
 public static bool InspectLast(IntPtr h,uint owner){Guard(h,owner);Keys(new[]{Key(9,false),Key(9,true),Key(50,false),Key(50,true),Key(48,false),Key(48,true),Key(13,false),Key(13,true)});return true;}
 public static R Bounds(IntPtr h){R r;if(!GetWindowRect(h,out r))throw new Exception("Selection bounds unavailable.");return r;}
 public static void Resize(IntPtr h,uint owner,R original,double scale){Guard(h,owner);if(!SetWindowPos(h,IntPtr.Zero,original.L,original.T,(int)((original.Right-original.L)*scale),(int)((original.Bottom-original.T)*scale),4))throw new Exception("Native resize failed.");}
 public static void WindowState(IntPtr h,uint owner,int state){Guard(h,owner);ShowWindow(h,state);}
 public static void CloseOwned(IntPtr h,uint owner){var cls=new StringBuilder(128);GetClassName(h,cls,128);if(h!=IntPtr.Zero&&Owner(h)==owner&&cls.ToString().StartsWith("Thunder"))PostMessage(h,0x10,IntPtr.Zero,IntPtr.Zero);}
 public class Result{public bool ExactDialog,WindowWithinDisplay,Captured,CancelDispatched,InspectDispatched,ForegroundRequestAccepted;public int DesktopError;}
 public static Result Capture(IntPtr h,uint owner,string path){
  var r=new Result();Guard(h,owner);var cls=new StringBuilder(128);GetClassName(h,cls,128);r.ExactDialog=cls.ToString().StartsWith("Thunder");
  P p;if(!GetCursorPos(out p)){r.DesktopError=Marshal.GetLastWin32Error();return r;}
  IntPtr previous=SetThreadDpiAwarenessContext(new IntPtr(-4));if(previous==IntPtr.Zero)throw new Exception("Physical pixel context unavailable.");
  try{R b;if(DwmGetWindowAttribute(h,9,out b,Marshal.SizeOf(typeof(R)))!=0&&!GetWindowRect(h,out b))throw new Exception("Selection bounds unavailable.");int left=GetSystemMetrics(76),top=GetSystemMetrics(77);r.WindowWithinDisplay=b.L>=left&&b.T>=top&&b.Right<=left+GetSystemMetrics(78)&&b.Bottom<=top+GetSystemMetrics(79);
   if(r.WindowWithinDisplay){using(var bitmap=new Bitmap(b.Right-b.L,b.Bottom-b.T)){using(var g=Graphics.FromImage(bitmap)){g.CopyFromScreen(b.L,b.T,0,0,bitmap.Size);}Guard(h,owner);bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);r.Captured=true;}}
  }finally{SetThreadDpiAwarenessContext(previous);}return r;
 }
}
'@
        $owner=[ReconcileSelectionInput]::Owner([IntPtr]$handle)
        $started=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($jobReady,'Ready')
        Stage 'ObserverReady'
        $window=[IntPtr]::Zero
        try{
            $deadline=[DateTime]::UtcNow.AddSeconds(60)
            while([DateTime]::UtcNow -lt $deadline){
                if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $started){throw 'Generated Excel identity changed.'}
                $window=[ReconcileSelectionInput]::Find($owner)
                if($window -ne [IntPtr]::Zero -and (Test-Path -LiteralPath $ready)){break}
                Start-Sleep -Milliseconds 100
            }
            if($window -eq [IntPtr]::Zero -or -not(Test-Path -LiteralPath $ready)){throw 'Native selection form unavailable.'}
            Stage 'FormObserved'
            $foreground=[ReconcileSelectionInput]::SetForegroundWindow($window)
            Stage 'ForegroundRequested'
            Start-Sleep -Milliseconds 250
            [ReconcileSelectionInput]::ScrollEnd($window,$owner)
            Stage 'ScrollInputDelivered'
            Start-Sleep -Milliseconds 250
            $result=[ReconcileSelectionInput]::Capture($window,$owner,$capture)
            Stage 'CaptureReturned'
            $original=[ReconcileSelectionInput]::Bounds($window)
            $layouts=@()
            $modes=@();if(-not $inspectLast){$modes=@('Grow','Minimum','Maximize','Restore')}
            foreach($mode in $modes){
                $before=Get-Content -LiteralPath $ready -Raw|ConvertFrom-Json
                switch($mode){
                    'Grow'{[ReconcileSelectionInput]::Resize($window,$owner,$original,1.25)}
                    'Minimum'{[ReconcileSelectionInput]::Resize($window,$owner,$original,0.75)}
                    'Maximize'{[ReconcileSelectionInput]::WindowState($window,$owner,3)}
                    'Restore'{[ReconcileSelectionInput]::WindowState($window,$owner,9);[ReconcileSelectionInput]::Resize($window,$owner,$original,1.0)}
                }
                $layout=$null
                for($i=0;$i -lt 30;$i++){
                    Start-Sleep -Milliseconds 100
                    try{$layout=Get-Content -LiteralPath $ready -Raw|ConvertFrom-Json}catch{continue}
                    if($layout.Revision -gt $before.Revision){break}
                }
                if($null -eq $layout -or $layout.Revision -le $before.Revision){throw 'Native resize did not produce fresh layout evidence.'}
                [ReconcileSelectionInput]::ScrollEnd($window,$owner)
                Start-Sleep -Milliseconds 200
                $picture=[ReconcileSelectionInput]::Capture($window,$owner,($capture.Replace('.png','-'+$mode.ToLowerInvariant()+'.png')))
                $layouts+=[pscustomobject]@{Mode=$mode;Width=$layout.Width;Height=$layout.Height;Geometry=$layout.Geometry;CompleteHistory=$layout.CompleteHistory;Scrollable=$layout.Scrollable;Captured=$picture.Captured;WithinDisplay=$picture.WindowWithinDisplay}
                Stage ('Layout'+$mode)
            }
            ConvertTo-Json -InputObject @($layouts)|Set-Content ($jobReady+'.layouts.json')
            $result.ForegroundRequestAccepted=$foreground
            if($inspectLast){$result.InspectDispatched=[ReconcileSelectionInput]::InspectLast($window,$owner);Stage 'InspectInputDelivered'}
            else{$result.CancelDispatched=[ReconcileSelectionInput]::Cancel($window,$owner);Stage 'CancelInputDelivered'}
            for($i=0;$i -lt 50 -and [ReconcileSelectionInput]::Find($owner) -ne [IntPtr]::Zero;$i++){Start-Sleep -Milliseconds 100}
            if([ReconcileSelectionInput]::Find($owner) -ne [IntPtr]::Zero){throw 'Native action did not close the selection form.'}
            return $result
        }catch{
            Stage 'ObserverFailed'
            # The actual generated form supplies its handle; cleanup must also work
            # when a caption defect prevents the stricter acceptance lookup.
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -eq $started){
                if($window -eq [IntPtr]::Zero -and (Test-Path -LiteralPath $ready)){
                    $shown=Get-Content -LiteralPath $ready -Raw|ConvertFrom-Json
                    $window=[IntPtr][long]$shown.Window
                }
                [ReconcileSelectionInput]::CloseOwned($window,$owner)
            }
            throw
        }
    }
    try{
        for($i=0;$i -lt 100 -and -not(Test-Path -LiteralPath $jobReady);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path -LiteralPath $jobReady)){throw 'Selection keyboard observer not ready.'}
        $returned=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeNativeSelection' @([bool]$InspectLast))
        [void](Wait-Job $job -Timeout 5)
        $observed=Receive-Job $job -ErrorAction Stop
        if($null -eq $observed){throw 'Selection keyboard receipt unavailable.'}
        if($observed.DesktopError -ne 0){throw ('Desktop Win32 error '+$observed.DesktopError+' at '+[DateTimeOffset]::UtcNow.ToString('o'))}
        $facts=Get-Content -LiteralPath $closed -Raw|ConvertFrom-Json
        $receipt=[pscustomobject]@{ExactDialog=$observed.ExactDialog;LastChoicePresent=($facts.CompleteHistory -and $facts.AtEnd);ExactKeyPresent=$facts.CompleteHistory;HumanReadableTimestamp=$facts.Timestamp;TextFits=($facts.Geometry -and $facts.Scrollable);WindowWithinDisplay=$observed.WindowWithinDisplay;Captured=$observed.Captured;CancelDispatched=$observed.CancelDispatched;PromptCharacters=$facts.Characters;DesktopError=$observed.DesktopError;NativeKeyboardReachedEnd=$facts.AtEnd;ForegroundRequestAccepted=$observed.ForegroundRequestAccepted}
        $receipt|Add-Member -NotePropertyName InspectDispatched -NotePropertyValue ([bool]$observed.InspectDispatched)
        $receipt|Add-Member -NotePropertyName SelectionIs20 -NotePropertyValue ([bool]$facts.SelectionIs20)
        $receipt|ConvertTo-Json|Set-Content (Join-Path $reportRoot $(if($InspectLast){'native-selection-inspect.json'}else{'native-selection.json'}))
        $layouts=Get-Content ($jobReady+'.layouts.json') -Raw|ConvertFrom-Json
        return [pscustomobject]@{Returned=$returned;Observed=$receipt;Layouts=$layouts}
    }finally{if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job}
}
