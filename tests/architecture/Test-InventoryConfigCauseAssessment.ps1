[CmdletBinding()]
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $RepoRoot 'tests/tooling/Slice4beInventoryConfigMetadata.ps1')
$script:checks=0;$script:failed=0
foreach($case in @('KnownSaveWindows','BetweenWindowWriter','EarlierWriter','LaterWriter',
    'UnknownXmlAlongsideModified','ChangedConfigCell','OtherBinaryRecord','MalformedRevision',
    'UnknownCaller','ReadMutation','MissingWindowEvidence','NoChangedParts')){
    $binary=[pscustomobject]@{Parsed=$true;RecordOrderPreserved=$true;RecordHeadersPreserved=$true;BeforeRecordCount=4;AfterRecordCount=4;ChangedRecordTypes=@(3073);UnknownRevisionPointerChange=$false;ChangedRevisionPointerFields=@('uidLastSave')}
    $core=[pscustomobject]@{Parsed=$true;OtherXmlChanged=$false;ChangedFields=@('modified')}
    $phase=[pscustomobject]@{Procedure='modConfig.SaveConfigWorkbookIfWritable';Calls=3;ByteChangesDuringSave=3;UnavailableHashComparisons=0;CallSites=[pscustomobject]@{Receiving=1;Shipping=1;Production=1;Unknown=0;Completed=3};Windows=[pscustomobject]@{FirstBeforeMatches=$true;LastAfterMatches=$true;NoBetweenSaveChange=$true}}
    $read=$true;$cells=0;$parts=@('xl/workbook.bin','docProps/core.xml')
    switch($case){
        'BetweenWindowWriter' {$phase.Windows.NoBetweenSaveChange=$false}
        'EarlierWriter' {$phase.Windows.FirstBeforeMatches=$false}
        'LaterWriter' {$phase.Windows.LastAfterMatches=$false}
        'UnknownXmlAlongsideModified' {$core.OtherXmlChanged=$true}
        'ChangedConfigCell' {$cells=1}
        'OtherBinaryRecord' {$binary.ChangedRecordTypes=@(3073,158)}
        'MalformedRevision' {$binary.UnknownRevisionPointerChange=$true}
        'UnknownCaller' {$phase.CallSites.Shipping=0;$phase.CallSites.Unknown=1}
        'ReadMutation' {$read=$false}
        'MissingWindowEvidence' {$phase.PSObject.Properties.Remove('Windows')}
        'NoChangedParts' {$parts=@()}
    }
    $result=Get-InventoryConfigSaveCause $read $cells $true $parts 0 $binary $core $phase
    $script:checks++
    if($result -eq ($case -ceq 'KnownSaveWindows')){'PASS '+$case}else{$script:failed++;'FAIL '+$case}
}
'CONFIG_CAUSE_DIAGNOSTIC passed={0} failed={1} total={2}' -f ($script:checks-$script:failed),$script:failed,$script:checks
if($script:failed){exit 1}
