[CmdletBinding()]
param(
    [string]$RepoRoot='.',
    [string]$DeployRoot='deploy/validation-inventory-receipts-01',
    [Parameter(Mandatory=$true)][string]$DesktopMonitorReport,
    [switch]$LoadOnly,
    [switch]$Surfaces,
    [switch]$Domains,
    [switch]$Components,
    [switch]$Startup,
    [switch]$FullInitial,
    [switch]$AuthorizedFixture,
    [switch]$BindFixtureContext
)
# Diagnostic only: no acceptance or product RED claim. All probes are unsaved.
# Record counts/flags, never context handles, identities or source field values.
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
if(@(@($LoadOnly,$Surfaces,$Domains,$Components,$Startup,$FullInitial)|Where-Object {[bool]$_}).Count -gt 1){throw 'Choose one diagnostic group; use FullInitial for the exact combined order.'}
if($AuthorizedFixture -and ($LoadOnly -or $Surfaces -or $Domains -or $Components -or $Startup -or $FullInitial)){throw 'Authorized fixture requires the focused form diagnostic.'}
if($BindFixtureContext -and -not $AuthorizedFixture){throw 'Binding requires the generated live fixture.'}
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before this isolated diagnostic.'}
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_phase6_packaged_xlams.ps1'),[ref]$tokens,[ref]$errors)
if($errors){throw 'Packaged validator does not parse.'}
foreach($definition in $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true)){
    . ([scriptblock]::Create($definition.Extent.Text))
}
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class PackagedValidationProcess {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
. (Join-Path $repo 'tests/tooling/Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$closureRoot=Join-Path $repo ('reports/runtime/packaged-inventory-diagnostic/'+[guid]::NewGuid().ToString('N'))
$targetRoot=Join-Path $closureRoot 'fixture'
New-Item -ItemType Directory -Path $targetRoot -Force|Out-Null
$pins=@(Get-ChildItem -LiteralPath $deployPath -Filter '*.xlam'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Exactly five candidate packages required.'}
$script:PackagedDesktopStopped=$false
$resultRows=[Collections.Generic.List[object]]::new()
$packages=@{}; $excel=$null; $completed=$false; $normalClosure=$false
$started=[DateTimeOffset]::UtcNow
Write-Output ('DIAGNOSTIC='+$closureRoot)
try {
    Assert-PackagedDesktop
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AutomationSecurity=1
    $openOrder=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
    if($Startup -or $FullInitial){
        $source=[IO.File]::ReadAllText((Join-Path $repo 'tools/validate_phase6_packaged_xlams.ps1'))
        $start=$source.IndexOf('    $sentinelWb = $excel.Workbooks.Add()', [StringComparison]::Ordinal)
        $end=$source.IndexOf('    $componentSpecs = @(', [StringComparison]::Ordinal)
        if($FullInitial){
            $end=$source.IndexOf('    Close-PackagedValidationSession -Excel $excel -Phase "Initial"', [StringComparison]::Ordinal)
            $specAssignment=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -ceq '$validationSpecs'},$true))
            if($specAssignment.Count -ne 1){throw 'Validation specifications are ambiguous.'}
            . ([scriptblock]::Create($specAssignment[0].Extent.Text))
        }
        if($start -lt 0 -or $end -le $start){throw 'Startup diagnostic boundaries changed.'}
        $targetWorkbooks=[Collections.Generic.List[object]]::new()
        $openedWorkbooks=[Collections.Generic.List[object]]::new()
        $workbookMap=$packages
        $body=$source.Substring($start,$end-$start)
        # The maintained validator may now end its action scope before cleanup.
        if($FullInitial -and $source.Contains('# Finish the action scope before releasing this host.')){
            $body=$body.Substring(0,$body.LastIndexOf('    }', [StringComparison]::Ordinal))
        }
        & ([scriptblock]::Create($body))
    }else{
        foreach($name in $openOrder){
            Assert-PackagedDesktop
            $packages[$name]=$excel.Workbooks.Open((Join-Path $deployPath $name))
            if($name -ceq 'invSys.Core.xlam'){Set-PackagedValidationRoot -Excel $excel -Phase 'Initial'}
        }
    }
    if($Surfaces){
        $specAssignment=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -ceq '$validationSpecs'},$true))
        if($specAssignment.Count -ne 1){throw 'Validation specifications are ambiguous.'}
        . ([scriptblock]::Create($specAssignment[0].Extent.Text))
        foreach($spec in $validationSpecs){
            Assert-PackagedDesktop
            $book=$excel.Workbooks.Add()
            $book.SaveAs((Join-Path $targetRoot $spec.TargetFile),51)
            $book.Activate()
            Run-WorkbookMacro -Excel $excel -WorkbookName $spec.File -MacroName $spec.InitMacro
            try {
                Begin-QuietUi $excel
                if($spec.SafeMacro -ne ''){Run-WorkbookMacro -Excel $excel -WorkbookName $spec.File -MacroName $spec.SafeMacro}
                if($spec.ContainsKey('FormSmokeMacro')){
                    $formResult=[string](Run-WorkbookMacro1 -Excel $excel -WorkbookName $spec.File -MacroName $spec.FormSmokeMacro -Argument1 $book)
                    Add-ResultRow -Rows $resultRows -Check ($spec.Name+'.Form') -Passed ($formResult.StartsWith('OK|'))
                }
            } finally {End-QuietUi $excel}
            $surface=Test-WorkbookSurface -Workbook $book -TableSpecs $spec.Tables
            Add-ResultRow -Rows $resultRows -Check ($spec.Name+'.Surface') -Passed ($surface -ceq 'OK')
            $book.Save()
        }
    }
    if($Domains){
        foreach($macro in @('InventoryZeroCreateContractForAutomation','InventoryRetirementContractForAutomation')){
            Assert-PackagedDesktop
            $result=[string]$excel.Run("'invSys.Inventory.Domain.xlam'!modInventoryApply."+$macro)
            Add-ResultRow -Rows $resultRows -Check $macro -Passed ($result.StartsWith('OK|'))
        }
        foreach($domain in @('Inventory','Designs')){
            Assert-PackagedDesktop
            $packages['invSys.'+$domain+'.Domain.xlam'].Close($false)
            $result=[string]$excel.Run("'invSys.Core.xlam'!mod"+$domain+'DomainBridge.Diagnose'+$domain+'DomainBridge')
            Add-ResultRow -Rows $resultRows -Check ($domain+'.PeerAutoLoad') -Passed (-not [string]::IsNullOrWhiteSpace($result) -and -not $result.StartsWith($domain+' Domain unavailable'))
        }
    }
    if($Components){
        $specAssignment=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -ceq '$componentSpecs'},$true))
        if($specAssignment.Count -ne 1){throw 'Component specifications are ambiguous.'}
        . ([scriptblock]::Create($specAssignment[0].Extent.Text))
        foreach($spec in $componentSpecs){
            Assert-PackagedDesktop
            $presence=Test-VbComponentPresence -Workbook $packages[$spec.File] -ComponentName $spec.Component -ShouldExist $spec.Exists
            Add-ResultRow -Rows $resultRows -Check ($spec.File+'.'+$spec.Component) -Passed ($presence -ceq 'OK')
        }
    }
    if(-not $LoadOnly -and -not $Surfaces -and -not $Domains -and -not $Components -and -not $Startup -and -not $FullInitial){
        $project=$packages['invSys.Core.xlam'].VBProject
        $probe=$project.VBComponents.Add(1); $probe.Name='TestSmokeAuthority'
        $probe.CodeModule.AddFromString(@'
Option Explicit
Private mCalls As Long, mEmpty As Long
Private mCaptureExit As Long, mCaptureError As Long
Private mSourceExit As Long
Public Sub Reset()
    mCalls = 0: mEmpty = 0
End Sub
Public Sub Observe(ByVal key As String)
    mCalls = mCalls + 1
    If key = "" Then mEmpty = mEmpty + 1
End Sub
Public Function Facts() As String
    Facts = CStr(mCalls) & "|" & CStr(mEmpty)
End Function
Public Sub CaptureExit(ByVal line As Long, Optional ByVal errorCode As Long = 0)
    mCaptureExit = line: mCaptureError = errorCode
End Sub
Public Function CaptureFacts() As String
    CaptureFacts = CStr(mCaptureExit) & "|" & CStr(mCaptureError) & "|" & CStr(modAuth.IsSignedIn()) & "|" & CStr(mSourceExit)
End Function
Public Sub SourceExit(ByVal line As Long)
    mSourceExit = line
End Sub
'@)
        $module=$project.VBComponents.Item('modInventoryCommands').CodeModule
        $line=$module.ProcStartLine('CheckForm',0); $count=$module.ProcCountLines('CheckForm',0)
        $body=$module.Lines($line,$count)
        $boundary='    On Error GoTo Refused'
        if([regex]::Matches($body,[regex]::Escape($boundary)).Count -ne 1){throw 'Observation boundary changed; not product RED.'}
        $module.DeleteLines($line,$count)
        $module.InsertLines($line,$body.Replace($boundary,($boundary+"`r`n"+'    TestSmokeAuthority.Observe key')))
        if($AuthorizedFixture){
            $line=$module.ProcStartLine('Capture',0); $count=$module.ProcCountLines('Capture',0)
            $captureLines=$module.Lines($line,$count) -split '\r?\n'
            for($i=0;$i -lt $captureLines.Count;$i++){
                $captureLines[$i]=$captureLines[$i].Replace('Exit Function',('TestSmokeAuthority.CaptureExit '+($i+1)+': Exit Function'))
                if($captureLines[$i] -ceq 'Failed:'){$captureLines[$i]+="`r`n"+'    TestSmokeAuthority.CaptureExit 999, Err.Number'}
            }
            $module.DeleteLines($line,$count); $module.InsertLines($line,($captureLines -join "`r`n"))
            $sourceModule=$project.VBComponents.Item('cInventoryCommandInputs').CodeModule
            $line=$sourceModule.ProcStartLine('CaptureSource',0); $count=$sourceModule.ProcCountLines('CaptureSource',0)
            $sourceLines=$sourceModule.Lines($line,$count) -split '\r?\n'
            for($i=0;$i -lt $sourceLines.Count;$i++){$sourceLines[$i]=$sourceLines[$i].Replace('Exit Function',('TestSmokeAuthority.SourceExit '+($i+1)+': Exit Function'))}
            $sourceModule.DeleteLines($line,$count); $sourceModule.InsertLines($line,($sourceLines -join "`r`n"))
            . (Join-Path $repo 'tests/tooling/PackagedInventorySmokeFixture.ps1')
            try {$fixture=Initialize-PackagedInventorySmokeFixture -Excel $excel -AdminPackage $packages['invSys.Admin.xlam'] -Root $targetRoot -BindContext ([bool]$BindFixtureContext)}
            catch {
                Assert-PackagedDesktop
                $facts=([string]$excel.Run("'invSys.Core.xlam'!TestSmokeAuthority.CaptureFacts")).Split('|')
                [pscustomobject]@{CaptureExitLine=[int]$facts[0];ErrorCode=[int]$facts[1];SignedIn=$facts[2] -ceq 'True';SourceExitLine=[int]$facts[3]}|ConvertTo-Json|Set-Content (Join-Path $closureRoot 'context-failure.json')
                throw
            }
            Add-ResultRow -Rows $resultRows -Check 'Fixture.LiveCapturedAdmin' -Passed $fixture.ContextReady
        }
        Assert-PackagedDesktop
        $book=$excel.Workbooks.Add()
        $book.SaveAs((Join-Path $targetRoot 'Admin.Diagnostic.xlsx'),51)
        $book.Activate()
        foreach($case in @('AddVisibilityDropdown','Delete','Worksheet')){
            Assert-PackagedDesktop
            [void]$excel.Run("'invSys.Core.xlam'!TestSmokeAuthority.Reset")
            if($AuthorizedFixture){
                $action=if($case -ceq 'AddVisibilityDropdown'){'Add'}else{$case}
                $result=[string]$excel.Run("'invSys.Admin.xlam'!TestPackagedInventoryFixture."+$action,[string]$book.Name)
            }else{
                $macro="'invSys.Admin.xlam'!modAdmin.Inventory"+$case+'ContractForAutomation'
                if($case -ceq 'Worksheet'){$result=[string]$excel.Run($macro,$book)}else{$result=[string]$excel.Run($macro)}
            }
            Assert-PackagedDesktop
            $counts=([string]$excel.Run("'invSys.Core.xlam'!TestSmokeAuthority.Facts")).Split('|')
            $calls=[int]$counts[0]; $empty=[int]$counts[1]
            if($AuthorizedFixture){
                Add-ResultRow -Rows $resultRows -Check ($case+'.CapturedHandleObserved') -Passed ($calls -gt 0 -and $empty -eq 0)
                Add-ResultRow -Rows $resultRows -Check ($case+'.ExistingHelperSucceeded') -Passed ($result.StartsWith('OK|'))
                # Report only fixed boolean/count fields from the legacy envelope.
                $facts=@([regex]::Matches($result,'(?:^|\|)([A-Za-z]+)=(True|False)(?=\||$)')|ForEach-Object {$_.Groups[1].Value+'='+$_.Groups[2].Value})
                [pscustomobject]@{Case=$case;Calls=$calls;EmptyHandles=$empty;Facts=$facts}|ConvertTo-Json -Compress|Add-Content (Join-Path $closureRoot 'form-observations.jsonl')
            }else{
                Add-ResultRow -Rows $resultRows -Check ($case+'.EmptyHandleObserved') -Passed ($calls -gt 0 -and $calls -eq $empty) -Detail ('Calls='+$calls+';Empty='+$empty)
                Add-ResultRow -Rows $resultRows -Check ($case+'.ExistingHelperRefused') -Passed ($result.StartsWith('FAIL|')) -Detail 'Only the fixed success/failure prefix is retained.'
            }
        }
    }
    $completed=$true
} finally {
    try {
        if($null -ne $excel){Close-PackagedValidationSession -Excel $excel -Phase 'Diagnostic'; $normalClosure=$true}
    } finally {
        # The process handle can signal before Windows' process enumeration
        # drops the host. Retain the in-memory snapshot through that interval.
        Wait-RecordingCleanup -Creator $null -Worker $null
        $restored=Restore-InvSysTestSettingsSnapshot $settings
        $preserved=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
        $closed=-not [bool](Get-Process EXCEL -ErrorAction SilentlyContinue)
        $failed=@($resultRows|Where-Object {-not $_.Passed}).Count
        $resultRows.ToArray()|Select-Object Check,Passed|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $closureRoot 'checks.json')
        [pscustomobject]@{StartUTC=$started.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');LoadOnly=[bool]$LoadOnly;Surfaces=[bool]$Surfaces;Domains=[bool]$Domains;Components=[bool]$Components;Startup=[bool]$Startup;FullInitial=[bool]$FullInitial;Completed=$completed;NormalClosure=$normalClosure;ExcelClosed=$closed;SettingsRestored=$restored;PackagesPreserved=$preserved;Passed=$resultRows.Count-$failed;Failed=$failed;Acceptance=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $closureRoot 'closure.json')
        if(-not($restored -and $preserved -and $closed)){throw 'Diagnostic preservation or closure failed.'}
    }
}
Get-Content -LiteralPath (Join-Path $closureRoot 'closure.json')
if($failed -or -not $completed -or -not $normalClosure){exit 1}
