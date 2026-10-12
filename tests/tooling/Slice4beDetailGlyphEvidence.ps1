# Compare actual native rendering with extra label height in the disposable form.
# The same caption/font/position is retained; restore before any later assertion.
function Install-DetailGlyphProbe {
    $code=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $code.AddFromString(@'
Public Function DetailGlyphPaddingForTest(ByVal restore As Boolean) As Boolean
    Static originalHeight As Single, originalWidth As Single, originalAuto As Boolean, originalText As String, active As Boolean
    Dim instance As Object, value As Object
    Set instance = DetailFixtureFormForTest()
    Set value = instance.Controls("fraDetailMultiline").Controls("lblMultilineValue1")
    If restore Then
        If Not active Then Err.Raise 5, , "No glyph diagnostic to restore."
        value.AutoSize = False: value.Height = originalHeight: value.AutoSize = originalAuto
        DetailGlyphPaddingForTest = (Abs(value.Height - originalHeight) < 0.1 And Abs(value.Width - originalWidth) < 0.1 And value.AutoSize = originalAuto And StrComp(value.Caption, originalText, vbBinaryCompare) = 0)
        active = False
    Else
        If active Then Err.Raise 5, , "Glyph diagnostic already active."
        originalHeight = value.Height: originalWidth = value.Width: originalAuto = value.AutoSize: originalText = value.Caption
        active = True
        value.AutoSize = False: value.Height = originalHeight + 12
        DetailGlyphPaddingForTest = (value.Height > originalHeight And value.Width = originalWidth And StrComp(value.Caption, originalText, vbBinaryCompare) = 0)
    End If
End Function
'@)
}

function Measure-DetailAddedInk([string]$Before,[string]$After,[double[]]$Bounds) {
    Add-Type -AssemblyName System.Drawing
    $first=$null;$last=$null
    try {
        $first=[Drawing.Bitmap]::new($Before);$last=[Drawing.Bitmap]::new($After)
        if($first.Width -ne $last.Width -or $first.Height -ne $last.Height){throw 'Glyph capture dimensions changed.'}
        $sx=$first.Width/$Bounds[4];$sy=$first.Height/$Bounds[5]
        $left=[int][Math]::Ceiling($Bounds[0]*$sx);$top=[int][Math]::Ceiling($Bounds[1]*$sy)
        $right=[int][Math]::Floor($Bounds[2]*$sx);$bottom=[int][Math]::Floor($Bounds[3]*$sy)
        if($left -lt 0 -or $top -lt 0 -or $right -ge $first.Width -or $bottom -ge $first.Height -or $right -le $left -or $bottom -le $top){throw 'Glyph comparison region unavailable.'}
        $added=0
        for($y=$top;$y -lt $bottom;$y++){
            for($x=$left;$x -lt $right;$x++){
                $a=$first.GetPixel($x,$y);$b=$last.GetPixel($x,$y)
                if($a.R -gt 200 -and $a.G -gt 200 -and $a.B -gt 200 -and $b.R -lt 100 -and $b.G -lt 100 -and $b.B -lt 100){$added++}
            }
        }
        [pscustomobject]@{AddedInkPixels=$added;ComparedPixels=($right-$left)*($bottom-$top);Unclipped=($added -eq 0)}
    } finally {if($first){$first.Dispose()};if($last){$last.Dispose()}}
}

function Test-DetailGlyphRendering([long]$Window) {
    if(-not ('DetailGlyphBounds' -as [type])){
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class DetailGlyphBounds {
 [StructLayout(LayoutKind.Sequential)] struct Rect {public int Left,Top,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)] struct Point {public int X,Y;}
 [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h,ref Point p);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
 [DllImport("user32.dll")] static extern int GetSystemMetricsForDpi(int index,uint dpi);
 public static double[] Read(IntPtr form,IntPtr excel,double[] b){
  uint owner,expected;GetWindowThreadProcessId(form,out owner);GetWindowThreadProcessId(excel,out expected);
  Rect whole,client;Point origin=new Point();
  if(owner==0 || owner!=expected || b.Length!=6 || b[4]<=0 || b[5]<=0 || !GetWindowRect(form,out whole) || !GetClientRect(form,out client) || !ClientToScreen(form,ref origin))throw new Exception("Owned glyph geometry unavailable.");
  double sx=client.Right/b[4],sy=client.Bottom/b[5];uint dpi=GetDpiForWindow(form);
  return new double[]{origin.X-whole.Left+b[0]*sx+4,origin.Y-whole.Top+b[1]*sy+4,
   origin.X-whole.Left+(b[0]+b[2])*sx-GetSystemMetricsForDpi(2,dpi)-4,
   origin.Y-whole.Top+(b[1]+b[3])*sy-GetSystemMetricsForDpi(3,dpi)-4,
   whole.Right-whole.Left,whole.Bottom-whole.Top};
 }
}
'@
    }
    $geometry=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailGeometryForTest')
    $rows=@($geometry -split "`n"|Where-Object {$_ -like "fraDetailMultiline`t*"})
    if($rows.Count -ne 1){throw 'Glyph frame geometry is ambiguous.'}
    [double[]]$frame=@(($rows[0] -split "`t")[1..6]|ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::CurrentCulture)})
    $bounds=[DetailGlyphBounds]::Read([IntPtr]$Window,[IntPtr]$excel.Hwnd,$frame)
    $before=Join-Path $reportRoot 'event-detail-glyph-original.png'
    $after=Join-Path $reportRoot 'event-detail-glyph-padded-diagnostic.png'
    Capture-DetailOwnedEvidence (Split-Path $before -Leaf) $Window
    try {
        if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailGlyphPaddingForTest' @($false))){throw 'Glyph padding diagnostic did not preserve text/width.'}
        Capture-DetailOwnedEvidence (Split-Path $after -Leaf) $Window
        $ink=Measure-DetailAddedInk $before $after $bounds
        $ink|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'detail-glyph-ink.json')
        Check 'EventDetail.Multiline.FinalGlyphsUnclipped' $ink.Unclipped
    } finally {
        Check 'EventDetail.Multiline.GlyphDiagnosticRestored' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailGlyphPaddingForTest' @($true)))
    }
}
