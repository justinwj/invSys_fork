# Extends the same native keyboard/capture observer. Only disposable fixture data
# may appear in screenshots; receipts contain booleans and counts, not row values.
function Get-ReconcileConfirmationSource {
@'
public static partial class ReconcileSelectionInput {
 const string ConfirmationTitle="invSys Admin - Confirm Shipment Reconcile";
 [DllImport("user32.dll")]static extern bool EnumChildWindows(IntPtr h,CB cb,IntPtr p);
 [DllImport("user32.dll",EntryPoint="SendMessageTimeoutW",CharSet=CharSet.Unicode,SetLastError=true)]static extern IntPtr ReadTextMessage(IntPtr h,uint m,UIntPtr capacity,StringBuilder text,uint flags,uint timeout,out UIntPtr result);
 static string Text(IntPtr h){var s=new StringBuilder(32768);GetWindowText(h,s,s.Capacity);return s.ToString();}
 static string Class(IntPtr h){var s=new StringBuilder(128);GetClassName(h,s,s.Capacity);return s.ToString();}
 static IntPtr Child(IntPtr h,string kind,string starts){IntPtr found=IntPtr.Zero;EnumChildWindows(h,(c,p)=>{if(IsWindowVisible(c)&&Class(c)==kind&&Text(c).Replace("&","").StartsWith(starts))found=c;return true;},IntPtr.Zero);return found;}
 public static bool CloseWindow(IntPtr h,uint owner){Guard(h,owner);return PostMessage(h,0x10,IntPtr.Zero,IntPtr.Zero);}
 public static bool InspectFirst(IntPtr h,uint owner){Guard(h,owner);Keys(new[]{Key(9,false),Key(9,true),Key(49,false),Key(49,true),Key(13,false),Key(13,true)});return true;}
 public static IntPtr FindConfirmation(uint owner){return FindTitle(owner,ConfirmationTitle);}
 static bool PreviewNumber(string text,string label,double expected){var match=System.Text.RegularExpressions.Regex.Match(text,"(?m)^"+System.Text.RegularExpressions.Regex.Escape(label)+@"\s*([-0-9.]+)\r?$");double value;return match.Success&&Double.TryParse(match.Groups[1].Value,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out value)&&value==expected;}
 public class ConfirmationResult {public bool ExactDialog,ExactKey,ExactEvent,ExactPreview,TextFits,Captured,WithinDisplay,NarrativeExact,ActionDispatched,FontAvailable;public int Characters,ClientWidth,ClientHeight,MeasuredHeight;}
 public static ConfirmationResult ReadConfirmation(IntPtr h,uint owner,string eventId,string systemKey,string path){
  GuardTitle(h,owner,ConfirmationTitle);var prompt=Child(h,"Static","Type one line explaining why this admin correction is being made.");var edit=Child(h,"Edit","");
  var r=new ConfirmationResult();r.ExactDialog=prompt!=IntPtr.Zero&&edit!=IntPtr.Zero&&Child(h,"Button","OK")!=IntPtr.Zero&&Child(h,"Button","Cancel")!=IntPtr.Zero;
  if(!r.ExactDialog)throw new Exception("Expected confirmation controls unavailable.");
  string text=Text(prompt);r.Characters=text.Length;r.ExactKey=text.Contains("System_Key: "+systemKey);r.ExactEvent=text.Contains("Corrects server shipment deduction event: "+eventId);
  r.ExactPreview=PreviewNumber(text,"Current Qty:",4)&&PreviewNumber(text,"Qty after that event:",1)&&PreviewNumber(text,"Queued correction delta:",-3);
  var picture=Capture(h,owner,path,ConfirmationTitle);RequireDesktop(picture);r.Captured=picture.Captured;r.WithinDisplay=picture.WindowWithinDisplay;return r;
 }
 public static void TypeNarrative(IntPtr h,uint owner,string narrative){GuardTitle(h,owner,ConfirmationTitle);var keys=new System.Collections.Generic.List<I>();foreach(char c in narrative){keys.Add(new I{Type=1,Data=new U{Key=new K{Scan=c,Flags=4}}});keys.Add(new I{Type=1,Data=new U{Key=new K{Scan=c,Flags=6}}});}Keys(keys.ToArray());}
 public static bool NarrativeMatches(IntPtr h,uint owner,string text){GuardTitle(h,owner,ConfirmationTitle);var edit=Child(h,"Edit","");var value=new StringBuilder(4096);UIntPtr count;if(ReadTextMessage(edit,0xD,(UIntPtr)value.Capacity,value,2,1000,out count)==IntPtr.Zero)throw new Exception("Native narrative readback unavailable.");return value.ToString()==text;}
 public static bool ConfirmAction(IntPtr h,uint owner,bool submit){GuardTitle(h,owner,ConfirmationTitle);ushort key=(ushort)(submit?13:27);Keys(new[]{Key(key,false),Key(key,true)});return true;}
 public static void RequireDesktop(Result result){if(result.DesktopError!=0)throw new Exception("Desktop Win32 error "+result.DesktopError+" at "+DateTime.UtcNow.ToString("o"));}
}
'@
}

function Test-AdminReconcileConfirmation($Events,$State) {
    $prepared=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareNativeConfirmation')
    Check 'Reconcile.Native.Confirmation.OwnerPreparedHigherQuantity' $prepared
    if(-not $prepared){throw 'Native confirmation fixture failed; not behavioral RED.'}
    $first=[string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFirstChoiceEvent')
    $bytes=Hash $State.InventoryPath
    foreach($mode in @('Cancel','Submit')){
        $before=$Events.ListRows.Count
        $result=Invoke-ReconcileSelectionNative -Confirmation $mode -ExpectedEvent $first -ExpectedKey $State.KeyA
        $observed=$result.Confirmation
        $prefix='Reconcile.Native.Confirmation.'+$mode
        Check ($prefix+'.ExactReadablePreview') ($observed.ExactDialog -and $observed.ExactKey -and $observed.ExactEvent -and $observed.ExactPreview -and $observed.TextFits -and $observed.Captured -and $observed.WithinDisplay)
        Check ($prefix+'.NativeActionAndConclusion') ($result.Returned -and $observed.ActionDispatched -and ($mode -ceq 'Cancel' -or $observed.NarrativeExact))
        $added=if($mode -ceq 'Submit'){1}else{0}
        Check ($prefix+'.ExpectedQueueCount') ($Events.ListRows.Count -eq $before+$added)
        Check ($prefix+'.InventoryPreserved') (((Hash $State.InventoryPath) -ceq $bytes) -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)))
        if($mode -ceq 'Submit'){
            $valid=$false
            if($Events.ListRows.Count -eq $before+1){
            $row=$Events.ListRows.Item($Events.ListRows.Count)
            $payload=[string]$row.Range.Cells.Item(1,$Events.ListColumns.Item('PayloadJson').Index).Value2
            $items=ConvertFrom-Json -InputObject $payload
            $valid=$items -is [array] -and $items.Count -eq 1
            if($valid){$item=$items[0];$names=@($item.PSObject.Properties.Name);$valid=@('System_Key','CorrectedShipEventId','Qty','RepairNarrative'|Where-Object {$_ -cnotin $names}).Count -eq 0}
            if($valid){$valid=$item.System_Key -ceq $State.KeyA -and $item.CorrectedShipEventId -ceq $first -and $item.Qty -eq -3 -and $item.RepairNarrative -ceq 'Generated native training correction' -and @($names|Where-Object {$_ -ieq 'ROW'}).Count -eq 0}
            }
            Check ($prefix+'.ExactQueuedPayload') $valid
        }
    }
}
