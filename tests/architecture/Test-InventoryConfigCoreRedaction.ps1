[CmdletBinding()]
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$ErrorActionPreference='Stop'
. (Join-Path $RepoRoot 'tests/tooling/Slice4beInventoryConfigMetadata.ps1')
$script:checks=0;$script:failed=0;$script:observations=@()
function Check-Core([string]$Name,[bool]$Passed){
    $script:checks++
    if($Passed){'PASS '+$Name}else{$script:failed++;'FAIL '+$Name}
}
function Compare-Core([string]$Before,[string]$After){
    $result=Compare-InventoryCoreProperties (Get-InventoryCorePropertySnapshot ([Text.Encoding]::UTF8.GetBytes($Before))) (Get-InventoryCorePropertySnapshot ([Text.Encoding]::UTF8.GetBytes($After)))
    $script:observations+=$result
    return $result
}
$xml='<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>CREATOR-CANARY</dc:creator><cp:lastModifiedBy>AUTHOR-CANARY</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>'
$same=Compare-Core $xml $xml
Check-Core 'Unchanged.NoDifference' ($same.Parsed -and $same.ChangedFields.Count -eq 0 -and -not $same.OtherXmlChanged)
foreach($case in @(@('modified','2020-01-01','2020-02-02'),@('lastModifiedBy','AUTHOR-CANARY','OTHER-AUTHOR-CANARY'),@('revision','<cp:revision>1','<cp:revision>2'))){
    $result=Compare-Core $xml $xml.Replace($case[1],$case[2])
    Check-Core ('KnownField.'+$case[0]) ($result.Parsed -and ($result.ChangedFields -join '|') -ceq $case[0] -and -not $result.OtherXmlChanged)
}
$creator=Compare-Core $xml $xml.Replace('CREATOR-CANARY','OTHER-CREATOR-CANARY')
Check-Core 'OtherProperty.RemainsUnclassified' ($creator.Parsed -and $creator.OtherXmlChanged -and $creator.ChangedFields.Count -eq 0)
$unknown=Compare-Core $xml $xml.Replace('</cp:coreProperties>','<cp:UNKNOWN-CANARY>VALUE-CANARY</cp:UNKNOWN-CANARY></cp:coreProperties>')
Check-Core 'UnknownElement.RemainsUnclassified' ($unknown.Parsed -and $unknown.OtherXmlChanged -and $unknown.ChangedFields.Count -eq 0)
foreach($case in @(
    @('WrongRoot',$xml.Replace('http://schemas.openxmlformats.org/package/2006/metadata/core-properties','urn:UNKNOWN-CANARY')),
    @('Duplicate',$xml.Replace('<cp:revision>1</cp:revision>','<cp:revision>1</cp:revision><cp:revision>2</cp:revision>')),
    @('NestedValue',$xml.Replace('AUTHOR-CANARY','<cp:nested>VALUE-CANARY</cp:nested>')),
    @('UnknownAttribute',$xml.Replace('<cp:revision>','<cp:revision unknown="VALUE-CANARY">')),
    @('ExternalEntity','<!DOCTYPE cp:coreProperties [<!ENTITY remote SYSTEM "file:///NEVER-READ-CANARY">]>'+$xml.Replace('AUTHOR-CANARY','&remote;'))
)){
    $bad=Compare-Core $xml $case[1]
    Check-Core ('Reject.'+$case[0]) (-not $bad.Parsed -and $bad.OtherXmlChanged -and $bad.ChangedFields.Count -eq 0)
}
$bytes=[Text.Encoding]::UTF8.GetBytes($xml);$copy=[byte[]]$bytes.Clone()
[void](Get-InventoryCorePropertySnapshot $bytes)
Check-Core 'Inspection.InputBytesPreserved' ([Convert]::ToBase64String($bytes) -ceq [Convert]::ToBase64String($copy))
$safe=$true
foreach($observation in $script:observations){
    foreach($property in $observation.PSObject.Properties){if($property.Name -cnotin @('Parsed','ChangedFields','OtherXmlChanged')){$safe=$false}}
    foreach($field in $observation.ChangedFields){if($field -cnotin @('lastModifiedBy','modified','revision')){$safe=$false}}
}
$json=$script:observations|ConvertTo-Json -Depth 5 -Compress
Check-Core 'Redaction.FixedLabelsOnly' ($safe -and $json -notmatch 'CANARY|Hash|OuterXml|Value')
'CONFIG_CORE_DIAGNOSTIC passed={0} failed={1} total={2}' -f ($script:checks-$script:failed),$script:failed,$script:checks
if($script:failed){exit 1}
