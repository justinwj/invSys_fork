# Diagnostic only: real source recording and Create guide handler, fixed labels/ticks.
function Install-AdminSeedGuideTiming {
    . (Join-Path $PSScriptRoot 'Slice4beGuideViewTrace.ps1')
    Install-Slice4beGuideViewTrace @{
        Operations=@{
            modGuideEditor=@('OpenForRun')
            frmActionPathGuide=@('BindContext','LoadDraft','RefreshExpectation','ReadStepInstruction')
            modInventoryViewer=@('GuideDraftControlForTest')
        }
        Core=@{
            modActionGuideDraft=@('CanCreate','OpenDraft','Guard','ReadDraft','ReadText','ReadExpectation')
            modGuideDraftSource=@('ReadSelected')
            modPathExpectation=@('ReadSelected')
            modRecordingReader=@('ReadRun')
            modActivityPolicy=@('ReadPolicy')
            modTrackingPolicyModel=@('Project')
            modConfig=@('LoadConfig','CloseTransientConfigAfterLoad')
            modAuth=@('CanPerform')
        }
    }
}

function Test-AdminSeedGuideTiming($Fixture,$Source) {
    CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
    if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $Source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Exact source recording unavailable.'}
    $before=BoundPins $journalRoot
    foreach($role in @('Core','Operations')){[void](Run ('invSys.'+$role+'.xlam') 'TestViewTrace.Mark' @('SeedTiming.CreateGuide|Begin'))}
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $result=BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths'
    $clock.Stop()
    foreach($role in @('Core','Operations')){[void](Run ('invSys.'+$role+'.xlam') 'TestViewTrace.Mark' @('SeedTiming.CreateGuide|End'))}
    [pscustomobject]@{Boundary='ActualCreateGuide';ElapsedMilliseconds=$clock.ElapsedMilliseconds;DiagnosticOnly=$true}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'seed-guide-timing.json')
    Check 'SeedTiming.ActualCreateGuide' ($result -ceq 'DELIVERED')
    # Engineering regression budget for this two-step generated fixture, not a
    # permission/policy waiver or a claim about all warehouse sizes.
    Check 'SeedTiming.CreateGuideWithinFiveSeconds' ($clock.ElapsedMilliseconds -le 5000)
    Check 'SeedTiming.TwoObservedSteps' ((BoundControl 'lstGuideSteps' 'Rows' '' 'frmActionPathGuide') -ceq '2')
    Check 'SeedTiming.SourceEvidencePreserved' (BoundSame $before (BoundPins $journalRoot))
    CloseRecordingViewer
}
