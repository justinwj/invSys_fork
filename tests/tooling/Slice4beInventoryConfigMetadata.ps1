# Read-only OPC core-property comparison. Values and hashes remain memory-only.
# Only three fixed save-metadata labels may leave this comparison boundary.
function Get-InventoryCorePropertySnapshot([byte[]]$Bytes) {
    $reader=$null;$stream=$null;$sha=[Security.Cryptography.SHA256]::Create()
    try{
        if($null -eq $Bytes -or $Bytes.Length -eq 0){throw 'Unavailable part.'}
        $settings=[Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing=[Xml.DtdProcessing]::Prohibit;$settings.XmlResolver=$null
        $stream=[IO.MemoryStream]::new($Bytes,$false)
        $reader=[Xml.XmlReader]::Create($stream,$settings)
        $document=[Xml.XmlDocument]::new();$document.XmlResolver=$null
        $document.Load($reader)
        $core='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        $terms='http://purl.org/dc/terms/'
        $root=$document.DocumentElement
        if($null -eq $root -or $root.LocalName -cne 'coreProperties' -or $root.NamespaceURI -cne $core){throw 'Unknown root.'}
        $fields=@{}
        foreach($node in @($root.ChildNodes)){
            $known=($node.NodeType -eq [Xml.XmlNodeType]::Element) -and (
                ($node.LocalName -ceq 'modified' -and $node.NamespaceURI -ceq $terms) -or
                ($node.LocalName -cin @('lastModifiedBy','revision') -and $node.NamespaceURI -ceq $core))
            if(-not $known){continue}
            if($fields.ContainsKey($node.LocalName)){throw 'Duplicate metadata field.'}
            foreach($child in $node.ChildNodes){
                if($child.NodeType -notin @([Xml.XmlNodeType]::Text,[Xml.XmlNodeType]::Whitespace,[Xml.XmlNodeType]::SignificantWhitespace)){throw 'Unexpected metadata structure.'}
            }
            foreach($attribute in $node.Attributes){
                if($node.LocalName -cne 'modified' -or $attribute.LocalName -cne 'type' -or
                   $attribute.NamespaceURI -cne 'http://www.w3.org/2001/XMLSchema-instance'){throw 'Unexpected metadata attribute.'}
                $type=$attribute.Value.Split(':')
                if($type.Count -ne 2 -or $type[1] -cne 'W3CDTF' -or $node.GetNamespaceOfPrefix($type[0]) -cne $terms){throw 'Unknown date type.'}
            }
            $fields[$node.LocalName]=[Convert]::ToBase64String($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($node.OuterXml)))
            [void]$root.RemoveChild($node)
        }
        [pscustomobject]@{Parsed=$true;Fields=$fields;OtherXmlHash=[Convert]::ToBase64String($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($document.OuterXml)))}
    }catch{
        # Never export exception text: an XML exception can contain input values.
        [pscustomobject]@{Parsed=$false;Fields=@{};OtherXmlHash=''}
    }finally{
        if($null -ne $reader){$reader.Dispose()}
        if($null -ne $stream){$stream.Dispose()}
        $sha.Dispose()
    }
}

function Compare-InventoryCoreProperties($Before,$After) {
    if($null -eq $Before -or $null -eq $After -or -not $Before.Parsed -or -not $After.Parsed){
        return [pscustomobject]@{Parsed=$false;ChangedFields=@();OtherXmlChanged=$true}
    }
    $changed=@()
    foreach($name in @('lastModifiedBy','modified','revision')){
        if($Before.Fields.ContainsKey($name) -ne $After.Fields.ContainsKey($name) -or $Before.Fields[$name] -cne $After.Fields[$name]){$changed+=$name}
    }
    [pscustomobject]@{Parsed=$true;ChangedFields=$changed;OtherXmlChanged=($Before.OtherXmlHash -cne $After.OtherXmlHash)}
}

# A narrowly established byte-change cause, never blanket workbook semantics or
# an authorization verdict. Unknown parts, structures or save windows stay open.
function Get-InventoryConfigSaveCause([bool]$ReadPreserved,[int]$ChangedCells,
        [bool]$ShapesPreserved,[string[]]$ChangedParts,[int]$UnknownParts,$Binary,$Core,$Phase) {
    if(-not $ReadPreserved -or $ChangedCells -ne 0 -or -not $ShapesPreserved -or $UnknownParts -ne 0){return $false}
    if($null -eq $ChangedParts -or $ChangedParts.Count -eq 0 -or @($ChangedParts|Where-Object {$_ -cnotin @('xl/workbook.bin','docProps/core.xml')}).Count){return $false}
    if($null -eq $Phase -or $Phase.Procedure -cne 'modConfig.SaveConfigWorkbookIfWritable' -or $Phase.Calls -ne 3 -or
       $Phase.ByteChangesDuringSave -lt 1 -or $Phase.ByteChangesDuringSave -gt 3 -or $Phase.UnavailableHashComparisons -ne 0){return $false}
    if($null -eq $Phase.PSObject.Properties['CallSites'] -or $null -eq $Phase.PSObject.Properties['Windows']){return $false}
    $sites=$Phase.CallSites;$windows=$Phase.Windows
    if($sites.Receiving -ne 1 -or $sites.Shipping -ne 1 -or $sites.Production -ne 1 -or $sites.Unknown -ne 0 -or $sites.Completed -ne 3){return $false}
    if(-not $windows.FirstBeforeMatches -or -not $windows.LastAfterMatches -or -not $windows.NoBetweenSaveChange){return $false}
    if('xl/workbook.bin' -cin $ChangedParts){
        if($null -eq $Binary -or -not $Binary.Parsed -or -not $Binary.RecordOrderPreserved -or -not $Binary.RecordHeadersPreserved -or
           $Binary.BeforeRecordCount -ne $Binary.AfterRecordCount -or ($Binary.ChangedRecordTypes -join '|') -cne '3073' -or
           $Binary.UnknownRevisionPointerChange -or $Binary.ChangedRevisionPointerFields.Count -eq 0 -or
           @($Binary.ChangedRevisionPointerFields|Where-Object {$_ -cnotin @('xrevidLastSaved','documentId','CoauthVersionLast','CoauthVersionMax','uidLastSave')}).Count){return $false}
    }
    if('docProps/core.xml' -cin $ChangedParts){
        if($null -eq $Core -or -not $Core.Parsed -or $Core.OtherXmlChanged -or $Core.ChangedFields.Count -eq 0 -or
           @($Core.ChangedFields|Where-Object {$_ -cnotin @('lastModifiedBy','modified','revision')}).Count){return $false}
    }
    return $true
}
