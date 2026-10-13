# Capture only the fixed owner-generated warning in the disposable Excel process.
function Invoke-AdminRootFeedbackNative([string]$Mode,[string]$Root) {
    Initialize-SettingsCapture
    $handle=[long]$excel.Hwnd
    $ready=Join-Path $runRoot ('root-feedback-'+[guid]::NewGuid().ToString('N')+'.ready')
    $path=Join-Path $reportRoot ('root-feedback-'+$Mode+'.png')
    $failed=$Mode -ceq 'RefreshFailed'
    $prompt="Warehouse root preference could not be confirmed.`r`n"+$(if($failed){'Generated directory refresh failure.'}else{'Warehouse directory refreshed.'})
    $job=Start-Job -ArgumentList $handle,$ready,$path,$prompt,${function:Initialize-SettingsCapture}.ToString() -ScriptBlock {
        param($handle,$ready,$path,$prompt,$captureDefinition)
        $ErrorActionPreference='Stop'
        & ([scriptblock]::Create($captureDefinition))
        Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class RootFeedbackDialog {
    public delegate bool EnumProc(IntPtr h,IntPtr p);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h,EnumProc f,IntPtr p);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
    public static uint Owner(IntPtr h) {uint p;GetWindowThreadProcessId(h,out p);return p;}
    static string Text(IntPtr h) {var s=new StringBuilder(512);GetWindowText(h,s,s.Capacity);return s.ToString();}
    static string Class(IntPtr h) {var s=new StringBuilder(128);GetClassName(h,s,s.Capacity);return s.ToString();}
    static IntPtr FindOK(IntPtr h) {
        IntPtr found=IntPtr.Zero;int count=0;
        EnumChildWindows(h,(c,p)=>{if(Class(c)=="Button"&&Text(c).Replace("&","")=="OK"){found=c;count++;}return true;},IntPtr.Zero);
        return count==1?found:IntPtr.Zero;
    }
    public static bool Exact(IntPtr h,uint owner,string prompt) {
        if(Owner(h)!=owner||Class(h)!="#32770"||Text(h)!="invSys Admin")return false;
        bool text=false;int buttons=0;
        EnumChildWindows(h,(c,p)=>{if(Class(c)=="Static"&&Text(c)==prompt)text=true;if(Class(c)=="Button")buttons++;return true;},IntPtr.Zero);
        var ok=FindOK(h);
        return text&&buttons==1&&Owner(ok)==owner&&Text(ok).Replace("&","")=="OK";
    }
    public static bool Acknowledge(IntPtr h,uint owner,string prompt) {
        return Exact(h,owner,prompt)&&PostMessage(FindOK(h),0x00F5,IntPtr.Zero,IntPtr.Zero);
    }
}
'@
        $owner=[RootFeedbackDialog]::Owner([IntPtr]$handle)
        if($owner -eq 0){throw 'Generated Excel owner unavailable'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($ready,'Ready')
        $deadline=[DateTime]::UtcNow.AddSeconds(45)
        while([DateTime]::UtcNow -lt $deadline){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Generated Excel identity changed'}
            $dialog=[InvSysSettingsCapture]::OwnedVisibleForm('invSys Admin',[IntPtr]$handle)
            if($dialog -ne [IntPtr]::Zero -and [RootFeedbackDialog]::Exact($dialog,$owner,$prompt)){
                $captured=$false;$acknowledged=$false
                try{
                    [void][InvSysSettingsCapture]::ActivateByCaptionClick($dialog,[IntPtr]$handle)
                    Start-Sleep -Milliseconds 250
                    [InvSysSettingsCapture]::SaveVisibleWindow($dialog,$path)
                    $captured=[RootFeedbackDialog]::Exact($dialog,$owner,$prompt)
                }finally{$acknowledged=[RootFeedbackDialog]::Acknowledge($dialog,$owner,$prompt)}
                return [pscustomobject]@{ExactDialog=$true;Captured=$captured;Acknowledged=$acknowledged}
            }
            Start-Sleep -Milliseconds 100
        }
        throw 'Expected root warning did not appear'
    }
    try{
        for($i=0;$i -lt 100 -and -not(Test-Path $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path $ready)){throw 'Root feedback observer not ready'}
        [void](Invoke 'NativeRootFeedback' @($true))
        $facts=([string](Invoke 'InvokeRoot' @($Root,$failed,$(if($failed){''}else{'ReadFailure'}),$failed))) -split '\|'
        [void](Wait-Job $job -Timeout 5)
        $result=Receive-Job $job -ErrorAction Stop
        if($null -eq $result){throw 'Native root feedback result unavailable'}
        $prefix='WarehouseRoot.Native.'+$Mode+'.'
        Check ($prefix+'ExactOwnerWarning') ([bool]$result.ExactDialog)
        Check ($prefix+'VisibleCapture') ([bool]$result.Captured -and (Test-Path $path))
        Check ($prefix+'AcknowledgedAndHandlerReturned') ([bool]$result.Acknowledged -and $facts.Count -eq 14 -and $facts[13] -ceq $prompt.Replace("`r`n",' '))
    }finally{
        [void](Invoke 'NativeRootFeedback' @($false))
        if($job.State -eq 'Running'){Stop-Job $job}
        Remove-Job $job
    }
}
