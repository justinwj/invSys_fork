# Visible evidence from the callback's existing failure-message branch.
# Selection/confirmation remain the established automation seam; results are real.
. (Join-Path $PSScriptRoot 'Slice4beAdminReconcileFont.ps1')
function Install-AdminSeedFeedbackProbe($Module) {
    # Measure with Excel's own GDI font handle; that handle is process-local.
    Install-AdminReconcileFontProbe $Module
    $fontStart=$Module.ProcStartLine('NativeFontTick',0);$fontCount=$Module.ProcCountLines('NativeFontTick',0)
    $fontSource=$Module.Lines($fontStart,$fontCount)
    if(-not $fontSource.Contains('invSys Admin - Confirm Shipment Reconcile')){throw 'Native font observer title boundary unavailable.'}
    $Module.DeleteLines($fontStart,$fontCount);$Module.InsertLines($fontStart,$fontSource.Replace('invSys Admin - Confirm Shipment Reconcile','invSys Admin'))
    $Module.InsertLines($Module.CountOfDeclarationLines+1,'Public NativeFeedback As Boolean')
    $Module.AddFromString(@'
Public Function InvokeWithNativeFeedback(ByVal action As String) As Boolean
    On Error GoTo Failed
    NativeFeedback = True
    InvokeWithNativeFeedback = Invoke(action)
    NativeFeedback = False
    Exit Function
Failed:
    NativeFeedback = False
    Err.Raise vbObjectError + 911, , "Native Seed callback did not return"
End Function
'@)
    $code=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdmin').CodeModule
    $start=$code.ProcStartLine('Seed_DemoInventory',0);$count=$code.ProcCountLines('Seed_DemoInventory',0)
    $source=$code.Lines($start,$count)
    $needle="    Else`r`n        mSeedCallbackLastResult = `"FAIL|`" & report`r`n        If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, `"invSys Admin`""
    if(([regex]::Matches($source,[regex]::Escape($needle))).Count -ne 1){throw 'One callback result presentation boundary required.'}
    $replacement=$needle.Replace('        If Not mSeedCallbackAutomationEnabled Then',"        If TestSeedOutcome.NativeFeedback Then TestSeedOutcome.StartNativeFontProbe report`r`n        If Not mSeedCallbackAutomationEnabled Or TestSeedOutcome.NativeFeedback Then")
    $replacement+="`r`n        If TestSeedOutcome.NativeFeedback Then TestSeedOutcome.StopNativeFontProbe"
    $changed=$source.Replace($needle,$replacement)
    $code.DeleteLines($start,$count);$code.InsertLines($start,$changed)
}

function Invoke-AdminSeedFeedback([ValidateSet('SEED','DELETE')][string]$Action) {
    $hostBook=$excel.Workbooks.Add();$oldVisible=[bool]$excel.Visible
    $job=$null
    try {
        $hostBook.Windows.Item(1).Visible=$true;$excel.Visible=$true;$hostBook.Activate()
        $handle=[long]$excel.Hwnd
        $ready=Join-Path $runRoot ('seed-feedback-'+[guid]::NewGuid().ToString('N')+'.ready')
        $capture=Join-Path $reportRoot ('seed-feedback-'+$Action.ToLowerInvariant()+'.png')
        $expected=if($Action -ceq 'SEED'){'Demo inventory event was queued; application was not confirmed.'}else{'Delete event was queued; application was not confirmed.'}
        $job=Start-Job -ArgumentList $handle,$ready,$capture,$expected -ScriptBlock {
            param($handle,$ready,$capture,$expected)
            $ErrorActionPreference='Stop'
            Add-Type -ReferencedAssemblies System.Drawing @'
using System;using System.Text;using System.Drawing;using System.Runtime.InteropServices;
public static class SeedFeedbackWindow {
 public delegate bool CB(IntPtr h,IntPtr p);
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [DllImport("user32.dll")]static extern bool EnumWindows(CB cb,IntPtr p);
 [DllImport("user32.dll")]static extern bool EnumChildWindows(IntPtr h,CB cb,IntPtr p);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")]static extern bool GetWindowRect(IntPtr h,out R r);
 [DllImport("user32.dll",SetLastError=true)]static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]static extern int GetSystemMetrics(int index);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetCursorPos(out P p);
 public static uint Owner(IntPtr h){uint p;GetWindowThreadProcessId(h,out p);return p;}
 static string Text(IntPtr h){var b=new StringBuilder(32768);GetWindowText(h,b,b.Capacity);return b.ToString();}
 static string Class(IntPtr h){var b=new StringBuilder(128);GetClassName(h,b,b.Capacity);return b.ToString();}
 public static IntPtr Find(uint owner){IntPtr found=IntPtr.Zero;EnumWindows((h,p)=>{if(Owner(h)==owner&&IsWindowVisible(h)&&Class(h)=="#32770"&&Text(h)=="invSys Admin")found=h;return found==IntPtr.Zero;},IntPtr.Zero);return found;}
 static IntPtr Child(IntPtr h,string kind,string prefix){IntPtr found=IntPtr.Zero;EnumChildWindows(h,(c,p)=>{if(IsWindowVisible(c)&&Class(c)==kind&&Text(c).Replace("&","").StartsWith(prefix))found=c;return true;},IntPtr.Zero);return found;}
 public class Result {public bool ExpectedMessage,FontAvailable,TextFits,WindowWithinDisplay,Captured,ForegroundBefore,ForegroundAfter,OkDispatched;public int Characters,DesktopError,ClientWidth,ClientHeight,MeasuredHeight;}
 public static Result Inspect(IntPtr h,uint owner,string expected,string path){
  if(Owner(h)!=owner||Find(owner)!=h)throw new Exception("Generated callback dialog identity changed.");
  var ok=Child(h,"Button","OK");var prompt=Child(h,"Static",expected.StartsWith("Demo")?"Demo inventory":"Delete event");
  if(ok==IntPtr.Zero||prompt==IntPtr.Zero)throw new Exception("Expected callback message controls unavailable.");
  var r=new Result();IntPtr previous=SetThreadDpiAwarenessContext(new IntPtr(-4));
  try{
   if(previous==IntPtr.Zero)throw new Exception("Physical pixel context unavailable.");
   P point;if(!GetCursorPos(out point)){r.DesktopError=Marshal.GetLastWin32Error();return r;}
   string text=Text(prompt);r.Characters=text.Length;r.ExpectedMessage=text.StartsWith(expected+" ",StringComparison.Ordinal);
   R bounds;GetWindowRect(h,out bounds);
   int left=GetSystemMetrics(76),top=GetSystemMetrics(77),right=left+GetSystemMetrics(78),bottom=top+GetSystemMetrics(79);
   r.WindowWithinDisplay=bounds.L>=left&&bounds.T>=top&&bounds.Right<=right&&bounds.Bottom<=bottom;
   r.ForegroundBefore=GetForegroundWindow()==h;
   if(r.ForegroundBefore&&r.WindowWithinDisplay){using(var bitmap=new Bitmap(bounds.Right-bounds.L,bounds.Bottom-bounds.T)){using(var g=Graphics.FromImage(bitmap)){g.CopyFromScreen(bounds.L,bounds.T,0,0,bitmap.Size);}r.ForegroundAfter=GetForegroundWindow()==h;if(r.ForegroundAfter&&Find(owner)==h){bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);r.Captured=true;}}}
   return r;
  }finally{r.OkDispatched=PostMessage(ok,0xF5,IntPtr.Zero,IntPtr.Zero);if(previous!=IntPtr.Zero)SetThreadDpiAwarenessContext(previous);}
 }
}
'@
            $owner=[SeedFeedbackWindow]::Owner([IntPtr]$handle)
            if($owner -eq 0){throw 'Generated Excel owner unavailable.'}
            $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
            [IO.File]::WriteAllText($ready,'Ready')
            $deadline=[DateTime]::UtcNow.AddSeconds(60)
            while([DateTime]::UtcNow -lt $deadline){
                if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Generated Excel owner changed.'}
                $dialog=[SeedFeedbackWindow]::Find($owner)
                if($dialog -ne [IntPtr]::Zero){
                    [void][SeedFeedbackWindow]::SetForegroundWindow($dialog)
                    Start-Sleep -Milliseconds 250
                    return [SeedFeedbackWindow]::Inspect($dialog,$owner,$expected,$capture)
                }
                Start-Sleep -Milliseconds 100
            }
            throw 'Native callback feedback did not appear.'
        }
        for($i=0;$i -lt 100 -and -not(Test-Path -LiteralPath $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path -LiteralPath $ready)){throw 'Feedback observer not ready.'}
        $ack=[bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWithNativeFeedback' @($Action))
        [void](Wait-Job $job -Timeout 5)
        $observed=Receive-Job $job -ErrorAction Stop
        if($null -eq $observed){throw 'Native feedback observation unavailable.'}
        if($observed.DesktopError -ne 0){throw ('Desktop Win32 error '+$observed.DesktopError+' at '+[DateTimeOffset]::UtcNow.ToString('o'))}
        $fontProof=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.NativeFontProof')
        $fontParts=@($fontProof.Split('|')|ForEach-Object {[int]$_})
        if($fontParts.Count -ne 4 -or $fontParts[0] -ne $observed.Characters){throw 'In-process message font measurement unavailable.'}
        $observed.FontAvailable=$fontParts[1] -gt 0 -and $fontParts[3] -gt 0
        $observed.ClientWidth=$fontParts[1];$observed.ClientHeight=$fontParts[2];$observed.MeasuredHeight=$fontParts[3]
        $observed.TextFits=$observed.FontAvailable -and $fontParts[3] -le $fontParts[2]
        $observed|Select-Object ExpectedMessage,FontAvailable,TextFits,WindowWithinDisplay,Captured,ForegroundBefore,ForegroundAfter,OkDispatched,Characters,DesktopError,ClientWidth,ClientHeight,MeasuredHeight|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('seed-feedback-'+$Action.ToLowerInvariant()+'.json'))
        return [pscustomobject]@{Action=$Action;Acknowledged=$ack;Observed=$observed}
    }finally{
        if($null -ne $job){if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job}
        $hostBook.Close($false);$excel.Visible=$oldVisible
    }
}
