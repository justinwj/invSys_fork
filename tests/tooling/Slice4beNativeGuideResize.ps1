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
    try {
        if(-not [InvSysB0NativeResize]::ShowWindowAsync($window,3)){throw 'Native maximize request was not posted; not behavioral RED.'}
        $max=Read-NativeResizeState $true
        Check ($Prefix+'.NativeResize.Maximized') $max.StateObserved
        Check ($Prefix+'.NativeResize.MaximizedControlsFit') $max.Fits
        CaptureOwnedFormByCaptionEvidence $Title ($ImagePrefix+'-maximized.png')
        if(-not [InvSysB0NativeResize]::ShowWindowAsync($window,9)){throw 'Native restore request was not posted; not behavioral RED.'}
        $restored=Read-NativeResizeState $false
        $same=$true
        foreach($edge in @('Left','Top','Right','Bottom')){
            $same=$same -and [Math]::Abs($before.$edge-$restored.Bounds.$edge) -le 1
        }
        Check ($Prefix+'.NativeResize.Restored') $restored.StateObserved
        Check ($Prefix+'.NativeResize.RestoredBounds') $same
        Check ($Prefix+'.NativeResize.RestoredControlsFit') $restored.Fits
        CaptureOwnedFormByCaptionEvidence $Title ($ImagePrefix+'-restored.png')
        [pscustomobject]@{Form=$Form;Before=$before;Maximized=$max;Restored=$restored;HelperRequestsDpiContextChanges=$false}|
            ConvertTo-Json -Depth 5 -Compress|Add-Content (Join-Path $reportRoot 'b0-native-resize.jsonl')
    } finally {
        if([InvSysB0NativeResize]::Owned($window,$excelWindow) -and [InvSysB0NativeResize]::IsZoomed($window)){
            [void][InvSysB0NativeResize]::ShowWindowAsync($window,9)
        }
    }
}
