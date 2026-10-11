# D18 native maximize/restore evidence for disposable packaged B0 forms.
# ShowWindowAsync posts a request; IsZoomed and layout readback prove its result.
# https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-showwindowasync
function Test-B0NativeGuideResize([string]$Form,[string]$Title,[string]$Prefix,[string]$ImagePrefix) {
    Initialize-SettingsCapture
    if(-not ('InvSysB0NativeResize' -as [type])){
        Add-Type @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
public static class InvSysB0NativeResize {
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
 [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr h);
 [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll",EntryPoint="GetWindowLongW")] static extern int GetWindowLong(IntPtr h,int index);
 [DllImport("user32.dll",SetLastError=true)] static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr h,int command);
 public static bool Owned(IntPtr h,IntPtr excel) {
  uint actual,expected; GetWindowThreadProcessId(h,out actual); GetWindowThreadProcessId(excel,out expected);
  return IsWindow(h) && actual!=0 && actual==expected;
 }
 public static bool MaximizeOffered(IntPtr h) { return (GetWindowLong(h,-16)&0x10000)!=0; }
 public static Rect Bounds(IntPtr h) {
  Rect r; if(!GetWindowRect(h,out r))throw new Win32Exception(Marshal.GetLastWin32Error(),"Native form bounds unavailable.");
  return r;
 }
}
'@
    }
    $excelWindow=[IntPtr]$excel.Hwnd
    $window=[InvSysSettingsCapture]::OwnedVisibleForm($Title,$excelWindow)
    if(-not [InvSysB0NativeResize]::Owned($window,$excelWindow) -or [InvSysB0NativeResize]::IsZoomed($window)){
        throw 'Native resize requires one owned, visible, restored fixture form; not behavioral RED.'
    }
    $before=[InvSysB0NativeResize]::Bounds($window)
    Check ($Prefix+'.NativeResize.MaximizeOffered') ([InvSysB0NativeResize]::MaximizeOffered($window))
    function Read-NativeResizeState([bool]$Maximized) {
        $fits=$false;$state=$false
        for($attempt=0;$attempt -lt 40;$attempt++){
            if(-not [InvSysB0NativeResize]::Owned($window,$excelWindow)){throw 'Native resize lost its owned window.'}
            # Existing read-only probe repaints/pumps the real resize route; no sizes are assigned.
            $fits=(BoundControl '' 'Fit' 'Current' $Form) -ceq 'True'
            $state=[InvSysB0NativeResize]::IsZoomed($window) -eq $Maximized
            if($state){break}
            Start-Sleep -Milliseconds 50
        }
        [pscustomobject]@{StateObserved=$state;Fits=$fits;Bounds=[InvSysB0NativeResize]::Bounds($window)}
    }
    function Save-NativeResizeEvidence([string]$State) {
        # Bounds/IsZoomed can settle before the visible restore animation. Publish
        # only repeated screen frames; do not repaint, alter DPI or change layout.
        $name=$ImagePrefix+'-'+$State
        CaptureOwnedFormByCaptionEvidence $Title ($name+'-initial.png')
        $initial=Join-Path $reportRoot ($name+'-initial.png')
        $lastHash=(Get-FileHash $initial).Hash
        $lastBounds=[InvSysB0NativeResize]::Bounds($window)|ConvertTo-Json -Compress
        $matches=1;$settled=$false
        $candidate=Join-Path $reportRoot ($name+'-sample.png')
        $frames=[Collections.Generic.List[object]]::new()
        for($sample=1;$sample -le 15;$sample++){
            if(-not [InvSysB0NativeResize]::Owned($window,$excelWindow)){throw 'Native capture lost its owned window.'}
            [InvSysSettingsCapture]::SaveVisibleWindow($window,$candidate)
            $hash=(Get-FileHash $candidate).Hash
            $bounds=[InvSysB0NativeResize]::Bounds($window)|ConvertTo-Json -Compress
            if($hash -ceq $lastHash -and $bounds -ceq $lastBounds){$matches++}else{$matches=1}
            $frames.Add([pscustomobject]@{Sample=$sample;Sha256=$hash;Bounds=$bounds;ConsecutiveMatches=$matches})
            if($matches -ge 3){$settled=$true;break}
            $lastHash=$hash;$lastBounds=$bounds
        }
        [pscustomobject]@{Form=$Form;State=$State;Settled=$settled;InitialSha256=(Get-FileHash $initial).Hash;Frames=$frames.ToArray()}|
            ConvertTo-Json -Depth 5 -Compress|Add-Content (Join-Path $reportRoot 'b0-native-capture-settling.jsonl')
        Check ($Prefix+'.NativeResize.'+$State+'CaptureSettled') $settled
        if(-not $settled){throw 'Native capture did not settle; no acceptance image published.'}
        Copy-Item -LiteralPath $candidate -Destination (Join-Path $reportRoot ($name+'.png'))
    }
    try {
        if(-not [InvSysB0NativeResize]::ShowWindowAsync($window,3)){throw 'Native maximize request was not posted; not behavioral RED.'}
        $max=Read-NativeResizeState $true
        Check ($Prefix+'.NativeResize.Maximized') $max.StateObserved
        Check ($Prefix+'.NativeResize.MaximizedControlsFit') $max.Fits
        Save-NativeResizeEvidence 'maximized'
        if(-not [InvSysB0NativeResize]::ShowWindowAsync($window,9)){throw 'Native restore request was not posted; not behavioral RED.'}
        $restored=Read-NativeResizeState $false
        $same=$true
        foreach($edge in @('Left','Top','Right','Bottom')){
            $same=$same -and [Math]::Abs($before.$edge-$restored.Bounds.$edge) -le 1
        }
        Check ($Prefix+'.NativeResize.Restored') $restored.StateObserved
        Check ($Prefix+'.NativeResize.RestoredBounds') $same
        Check ($Prefix+'.NativeResize.RestoredControlsFit') $restored.Fits
        Save-NativeResizeEvidence 'restored'
        if($NativeResizeCaptureDiagnostic){
            $series=[Collections.Generic.List[object]]::new()
            $clock=[Diagnostics.Stopwatch]::StartNew()
            $initial=Join-Path $reportRoot ($ImagePrefix+'-restored.png')
            $series.Add([pscustomobject]@{Stage='Initial';ElapsedMilliseconds=0;Sha256=(Get-FileHash $initial).Hash;Bounds=[InvSysB0NativeResize]::Bounds($window)})
            for($sample=1;$sample -le 5;$sample++){
                Start-Sleep -Milliseconds 250
                if(-not [InvSysB0NativeResize]::Owned($window,$excelWindow)){throw 'Native capture lost its owned window.'}
                $path=Join-Path $reportRoot ($ImagePrefix+'-later-'+$sample+'.png')
                # Preserve the published image. These extra screen reads use the
                # unchanged capture helper without another caption click or COM repaint.
                [InvSysSettingsCapture]::SaveVisibleWindow($window,$path)
                $series.Add([pscustomobject]@{Stage='Later'+$sample;ElapsedMilliseconds=$clock.ElapsedMilliseconds;Sha256=(Get-FileHash $path).Hash;Bounds=[InvSysB0NativeResize]::Bounds($window)})
            }
            $stable=$series[4].Sha256 -ceq $series[5].Sha256
            Check ($Prefix+'.NativeCapture.LaterFramesAgree') $stable
            Check ($Prefix+'.NativeCapture.InitialMatchesLater') ($stable -and $series[0].Sha256 -ceq $series[5].Sha256)
            [void](Read-NativeResizeState $false)
            $path=Join-Path $reportRoot ($ImagePrefix+'-after-repaint.png')
            [InvSysSettingsCapture]::SaveVisibleWindow($window,$path)
            $series.Add([pscustomobject]@{Stage='AfterExistingRepaint';ElapsedMilliseconds=$clock.ElapsedMilliseconds;Sha256=(Get-FileHash $path).Hash;Bounds=[InvSysB0NativeResize]::Bounds($window)})
            Check ($Prefix+'.NativeCapture.ExistingRepaintPreservesFrame') ($stable -and $series[5].Sha256 -ceq $series[6].Sha256)
            [pscustomobject]@{Form=$Form;Frames=$series.ToArray();DiagnosticOnly=$true}|ConvertTo-Json -Depth 6 -Compress|
                Add-Content (Join-Path $reportRoot 'b0-native-capture-series.jsonl')
        }
        [pscustomobject]@{Form=$Form;Before=$before;Maximized=$max;Restored=$restored;HelperRequestsDpiContextChanges=$false}|
            ConvertTo-Json -Depth 5 -Compress|Add-Content (Join-Path $reportRoot 'b0-native-resize.jsonl')
    } finally {
        if([InvSysB0NativeResize]::Owned($window,$excelWindow) -and [InvSysB0NativeResize]::IsZoomed($window)){
            [void][InvSysB0NativeResize]::ShowWindowAsync($window,9)
        }
    }
}
