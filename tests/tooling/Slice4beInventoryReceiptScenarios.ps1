# Test selection only. Default sets preserve the existing receipt assertions.
function Get-InventoryReceiptScenarios {
    param(
        [Parameter(Mandatory=$true)][ValidateSet('Phase','Result','Canonical','Ownership')][string]$Cohort,
        [string]$Scenario=''
    )
    $cases=switch($Cohort){
        'Phase' {@('CatalogOnly','NoopQuantity','QuantityChange','QueueFailure')}
        'Result' {@('OperatorSaveThrows','OperatorSaveThrowsLaterQuantity','AfterApplyAuthorityLost')}
        'Canonical' {@('BeforeCatalogSaveThrows','AfterCatalogSaveThrows','BeforeCatalogSaveAuthorityLost','AfterCatalogSaveAuthorityLost')}
        'Ownership' {@('QueueFailure')}
    }
    if($Scenario -ne ''){
        if($Scenario -cnotin $cases){throw 'Receipt scenario does not belong to the selected cohort.'}
        return $Scenario
    }
    return $cases
}
