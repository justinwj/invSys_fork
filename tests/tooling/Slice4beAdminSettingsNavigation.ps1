# Observe actual MSForms events: owned SendInput clicks and focused keyboard messages.
function Install-AdminSettingsNavigationProbe($TestModule,$FormCode) {
    foreach($field in @('mPages','mTrackingSections')){
        for($line=1;$line -le $FormCode.CountOfDeclarationLines;$line++){
            if($FormCode.Lines($line,1).Trim() -ceq ('Private '+$field+' As MSForms.MultiPage')){$FormCode.ReplaceLine($line,('Private WithEvents '+$field+' As MSForms.MultiPage'));break}
        }
        foreach($event in @('Change','Click','MouseDown','MouseUp','KeyDown','KeyUp')){
            $parameters=if($event -in @('MouseDown','MouseUp')){'ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single'}elseif($event -in @('KeyDown','KeyUp')){'ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer'}elseif($event -ceq 'Click'){'ByVal Index As Long'}else{''}
            $name=$field+'_'+$event;$start=0
            try{$start=$FormCode.ProcStartLine($name,0)}catch{}
            $trace='    mEntryNavigationTrace = mEntryNavigationTrace & "'+$name+';"'
            if($start){$body=$FormCode.ProcBodyLine($name,0);while($FormCode.Lines($body,1).TrimEnd().EndsWith('_')){$body++};$FormCode.InsertLines($body+1,$trace)}
            else{$FormCode.AddFromString("Private Sub $name($parameters)`r`n$trace`r`nEnd Sub")}
        }
    }
    $FormCode.InsertLines(1,'Private mEntryNavigationTrace As String')
    $FormCode.AddFromString(@'
Public Sub EntryNavigationPrepare(ByVal inner As Boolean, ByVal source As Long)
    If inner Then
        mPages.Value = 1: mTrackingSections.Value = source
    Else
        mPages.Value = source
    End If
    mEntryNavigationTrace = ""
End Sub
Public Function EntryNavigationFocus(ByVal inner As Boolean) As Boolean
    On Error GoTo Unavailable
    If inner Then mTrackingSections.SetFocus Else mPages.SetFocus
    EntryNavigationFocus = True
Unavailable:
End Function
Public Function EntryNavigationState(ByVal inner As Boolean) As Long
    If inner Then EntryNavigationState = mTrackingSections.Value Else EntryNavigationState = mPages.Value
End Function
Public Function EntryNavigationTrace() As String
    EntryNavigationTrace = mEntryNavigationTrace
End Function
Public Function EntryNavigationGeometry(ByVal inner As Boolean) As String
    Dim x As Double, y As Double
    x = mPages.Left: y = mPages.Top
    If inner Then x = x + mTrackingSections.Left: y = y + 23 + mTrackingSections.Top
    EntryNavigationGeometry = CStr(x) & "|" & CStr(y) & "|" & CStr(Me.InsideWidth) & "|" & CStr(Me.InsideHeight)
End Function
Public Function EntryNavigationHandle() As Double
    EntryNavigationHandle = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(Me))
End Function
Public Function EntryNavigationNotice() As String
    EntryNavigationNotice = mLblStatus.Caption
End Function
'@)
    $TestModule.CodeModule.AddFromString(@'
Public Sub EntryNavPrepare(ByVal inner As Boolean, ByVal source As Long)
    mForm.EntryNavigationPrepare inner, source
End Sub
Public Function EntryNavFocus(ByVal inner As Boolean) As Boolean
    EntryNavFocus = mForm.EntryNavigationFocus(inner)
End Function
Public Function EntryNavState(ByVal inner As Boolean) As Long
    EntryNavState = mForm.EntryNavigationState(inner)
End Function
Public Function EntryNavTrace() As String
    EntryNavTrace = mForm.EntryNavigationTrace()
End Function
Public Function EntryNavGeometry(ByVal inner As Boolean) As String
    EntryNavGeometry = mForm.EntryNavigationGeometry(inner)
End Function
Public Function EntryNavHandle() As Double
    EntryNavHandle = mForm.EntryNavigationHandle()
End Function
Public Function EntryNavNotice() As String
    EntryNavNotice = mForm.EntryNavigationNotice()
End Function
Public Sub EntryNavReactivate()
    mForm.Hide
    mForm.Show vbModeless
End Sub
'@)
}

function Test-AdminSettingsNavigation($Fixture) {
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingNavigation.ps1')
    Initialize-ReceivingNavigationInput
    Add-Type @'
using System;using System.Runtime.InteropServices;
public static class SettingsNavigationCoordinates {
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [DllImport("user32.dll")]static extern bool GetClientRect(IntPtr h,out R r);
 [DllImport("user32.dll")]static extern bool GetWindowRect(IntPtr h,out R r);
 [DllImport("user32.dll")]static extern bool ClientToScreen(IntPtr h,ref P p);
 [StructLayout(LayoutKind.Sequential)]public struct Mouse{public int X,Y;public uint Data,Flags,Time;public UIntPtr Extra;}
 [StructLayout(LayoutKind.Sequential)]public struct Input{public uint Type;public Mouse Mouse;}
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]static extern IntPtr WindowFromPoint(P p);
 [DllImport("user32.dll")]static extern IntPtr GetAncestor(IntPtr h,uint flags);
 [DllImport("user32.dll")]static extern IntPtr SetThreadDpiAwarenessContext(IntPtr value);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetCursorPos(out P p);
 [DllImport("user32.dll",SetLastError=true)]static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll",SetLastError=true)]static extern uint SendInput(uint count,Input[] input,int size);
 static uint Owner(IntPtr h){uint p;GetWindowThreadProcessId(h,out p);return p;}
 public static void Click(IntPtr form,IntPtr excel,double x,double y,double w,double height){
  var oldDpi=SetThreadDpiAwarenessContext(new IntPtr(-4));if(oldDpi==IntPtr.Zero)throw new Exception("Native DPI coordinates unavailable.");
  try{R r;P original;if(!GetClientRect(form,out r))throw new Exception("Settings client bounds unavailable.");
   if(!GetCursorPos(out original))throw new Exception("Cursor read failed; Win32 error "+Marshal.GetLastWin32Error());
   var p=new P{X=(int)(x*(r.Right-r.L)/w),Y=(int)(y*(r.Bottom-r.T)/height)};
   if(!ClientToScreen(form,ref p)||Owner(form)==0||Owner(form)!=Owner(excel)||GetAncestor(WindowFromPoint(p),2)!=form)throw new Exception("Native point is not on the owned Settings form.");
   try{
    if(!SetCursorPos(p.X,p.Y))throw new Exception("SetCursorPos failed; Win32 error "+Marshal.GetLastWin32Error());
    P actual;bool read=GetCursorPos(out actual);int readError=read?0:Marshal.GetLastWin32Error();
    if(!read)throw new Exception("Cursor verification read failed; Win32 error "+readError);
    if(actual.X!=p.X||actual.Y!=p.Y)throw new Exception("SetCursorPos succeeded; unexpected coordinates. Expected "+p.X+","+p.Y+"; observed "+actual.X+","+actual.Y+" at "+DateTime.UtcNow.ToString("o"));
    if(GetAncestor(WindowFromPoint(actual),2)!=form)throw new Exception("Settings point became covered.");
    try{if(SendInput(1,new[]{new Input{Mouse=new Mouse{Flags=2}}},Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Mouse press failed; Win32 error "+Marshal.GetLastWin32Error());}
    finally{if(SendInput(1,new[]{new Input{Mouse=new Mouse{Flags=4}}},Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Mouse release failed; Win32 error "+Marshal.GetLastWin32Error());}
    System.Threading.Thread.Sleep(150);
   }finally{if(!SetCursorPos(original.X,original.Y))throw new Exception("Cursor restore failed; Win32 error "+Marshal.GetLastWin32Error());}
  }finally{SetThreadDpiAwarenessContext(oldDpi);}
 }
 public static string Read(IntPtr h,double x,double y,double w,double height){R r,b;GetClientRect(h,out r);GetWindowRect(h,out b);var p=new P{X=(int)(x*(r.Right-r.L)/w),Y=(int)(y*(r.Bottom-r.T)/height)};ClientToScreen(h,ref p);return (p.X-b.L)+"|"+(p.Y-b.T)+"|"+x+"|"+y+"|"+w+"|"+height;}
}
'@
    function Act([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestD5Commands.'+$Name) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    $cases=@(@('GENERAL',$false,1,0,22),@('EVENT_TRACKING',$false,0,1,88),@('TRACKING',$true,1,0,22),@('EVENT_DETAIL',$true,0,1,75),@('ACTION_PATHS',$true,1,2,140))
    $traceRows=@();$book=$null
    try{
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminEntry.EnablePolicy')){throw 'Explicit navigation policy unavailable; not behavioral RED.'}
        $book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Cells.Item(1,1).Value2='Navigation extra'
        $path=Join-Path $runRoot 'settings-navigation-host.xlsb';$book.SaveAs($path,50)
        $book.Close($false);$book=$null;$hostHash=(Get-FileHash $path).Hash
        $book=$excel.Workbooks.Open($path)
        $configHash=(Get-FileHash $Fixture.Config).Hash;$initial=@(Files)
        [void](Act 'OpenSettings');[void](Act 'ShowSettings')
        Check 'AdminNavigation.InternalShowExcluded' (@(Files).Count -eq $initial.Count)
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'admin-navigation-start.png'
        $handle=[IntPtr][long][double](Act 'EntryNavHandle')
        foreach($mode in @('Mouse','Keyboard')){
            foreach($case in $cases){
                $id='ADMIN_SETTINGS_PAGE_'+$case[0];$label='AdminNavigation.'+$mode+'.'+$case[0]
                $before=@(Files);[void](Act 'EntryNavPrepare' @($case[1],[int]$case[2]))
                Check ($label+'.ProgrammaticPreparationExcluded') (@(Files).Count -eq $before.Count)
                if($mode -ceq 'Mouse' -and $case[0] -ceq 'TRACKING'){
                    CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'admin-navigation-nested.png'
                }
                if($mode -ceq 'Keyboard'){
                    if(-not [bool](Act 'EntryNavFocus' @($case[1]))){throw 'Native tab focus unavailable; not behavioral RED.'}
                    # SetFocus may enter a child on a nested page. A same-tab click
                    # puts focus on the tab header without creating a selection.
                    $g=([string](Act 'EntryNavGeometry' @($case[1]))).Split('|')
                    $sourceX=if($case[1]){@(22,75,140)[[int]$case[2]]}else{@(22,65)[[int]$case[2]]}
                    [SettingsNavigationCoordinates]::Click($handle,[IntPtr]$excel.Hwnd,([double]$g[0]+$sourceX),([double]$g[1]+7),[double]$g[2],[double]$g[3])
                    Start-Sleep -Milliseconds 150
                    Check ($label+'.SameTabFocusExcluded') (@(Files).Count -eq $before.Count)
                    [void](Act 'EntryNavPrepare' @($case[1],[int]$case[2]))
                    [ReceivingNavigationInput]::Key($handle,[IntPtr]$excel.Hwnd,$(if($case[3] -lt $case[2]){37}else{39}))
                }else{
                    $g=([string](Act 'EntryNavGeometry' @($case[1]))).Split('|')
                    [SettingsNavigationCoordinates]::Read($handle,([double]$g[0]+$case[4]),([double]$g[1]+7),[double]$g[2],[double]$g[3])|Add-Content (Join-Path $reportRoot 'navigation-coordinates.txt')
                    [SettingsNavigationCoordinates]::Click($handle,[IntPtr]$excel.Hwnd,([double]$g[0]+$case[4]),([double]$g[1]+7),[double]$g[2],[double]$g[3])
                }
                Start-Sleep -Milliseconds 250
                $state=[int](Act 'EntryNavState' @($case[1]));$trace=[string](Act 'EntryNavTrace')
                $traceRows+=[pscustomobject]@{ControlId=$id;Mode=$mode;SelectedIndex=$state;Trace=$trace}
                $traceRows|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'navigation-native-trace.json')
                $event=if($mode -ceq 'Keyboard'){'_KeyDown;'}else{'_MouseDown;'}
                if($state -ne $case[3] -or -not $trace.Contains($event)){throw ('Native navigation did not reach intended tab: '+$id+' '+$mode+'; not behavioral RED.')}
                Check ($label+'.NativeInputSelectedDestination') $true
                $records=@(Files|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
                $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ 'SELECTED')
                $valid=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
                foreach($record in $records){$valid=$valid -and $record.ControlId -ceq $id -and $record.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $record.CatalogVersion -eq 33 -and $record.WarehouseId -ceq $Fixture.Warehouse -and @($record.SourceEventRefs).Count -eq 0}
                Check ($label+'.OneExactNavigationPair') $valid
            }
        }
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'admin-navigation-finish.png'
        [void](Act 'CloseSettings');$book.Close($false);$book=$null
        Check 'AdminNavigation.ConfigAndSavedHostPreserved' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash -and (Get-FileHash $path).Hash -ceq $hostHash)
    }finally{[void](Act 'CloseSettings');if($null -ne $book){$book.Close($false)}}
}
