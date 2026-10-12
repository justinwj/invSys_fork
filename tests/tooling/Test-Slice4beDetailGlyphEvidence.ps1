[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Slice4beDetailGlyphEvidence.ps1')
Add-Type -AssemblyName System.Drawing
$root=Join-Path (Resolve-Path $RepoRoot).Path ('reports/runtime/detail-glyph-detector/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$results=[Collections.Generic.List[object]]::new()
foreach($scale in @(1,1.5,2)){
    $paths=@();$label=([int](100*$scale)).ToString()
    foreach($kind in @('original','clipped','noise')){
        $path=Join-Path $root ($label+'-'+$kind+'.png');$paths+=$path
        $bitmap=[Drawing.Bitmap]::new([int](100*$scale),[int](100*$scale));$drawing=[Drawing.Graphics]::FromImage($bitmap)
        try{
            $drawing.Clear([Drawing.Color]::White)
            if($kind -ceq 'clipped'){$drawing.FillRectangle([Drawing.Brushes]::Black,[single](30*$scale),[single](50*$scale),[single](20*$scale),[single](2*$scale))}
            if($kind -ceq 'noise'){$drawing.FillRectangle([Drawing.Brushes]::Black,0,0,[single](10*$scale),[single](10*$scale))}
            $bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png)
        }finally{$drawing.Dispose();$bitmap.Dispose()}
    }
    foreach($case in @(@('Clipped',1,$false),@('Unchanged',0,$true),@('OutsideNoise',2,$true))){
        $actual=Measure-DetailAddedInk $paths[0] $paths[$case[1]] @(20,20,80,80,100,100)
        $passed=$actual.Unclipped -eq $case[2]
        $results.Add([pscustomobject]@{Check=$case[0]+'.'+$label;Passed=$passed})
    }
}
$results|ConvertTo-Json|Set-Content (Join-Path $root 'results.json')
$results|Format-Table -AutoSize
Write-Output ('REPORT_ROOT='+$root)
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
