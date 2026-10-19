# Fixed synthetic wire controls for the existing private Admin payload builder.
# These supplement the architectural RED and real form-handler regressions.
function Install-InventoryPrimitivePayloadProbe {
    $module=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdmin').CodeModule
    $module.AddFromString(@'
Public Function InventoryPrimitivePayloadForTest(ByVal kind As String) As String
    Dim fields As Object, qty As Double, picture As String, description As String
    description = "Fixture description"
    If kind <> "NoCustomFields" Then
        Set fields = CreateObject("Scripting.Dictionary")
        fields.CompareMode = vbTextCompare
        fields("TRACK_QTY") = "TRUE": fields("ITEM_KIND") = "STOCK"
    End If
    If kind = "ScalarsAndPictures" Then
        qty = 1.25
        picture = "one.png" & vbLf & "two.png"
        description = "First" & vbLf & "Second" & vbTab & Chr$(34) & "quoted" & Chr$(34) & " " & ChrW$(937)
        fields("CUSTOM_TEXT") = "back\slash"
        fields("CUSTOM_NUMBER") = 12.5
        fields("CUSTOM_BOOL") = True
        fields("CUSTOM_NULL") = Null
        fields("CUSTOM_EMPTY") = Empty
    ElseIf kind = "NonCounted" Then
        fields("TRACK_QTY") = "FALSE": fields("ITEM_KIND") = "UTILITY"
        fields("CUSTOM_TEXT") = "retained extension"
    ElseIf kind <> "NoCustomFields" Then
        Err.Raise vbObjectError + 2190, , "Unknown payload control."
    End If
    InventoryPrimitivePayloadForTest = BuildInventoryCreatePayloadAdmin( _
        "TEST-KEY", "FIXTURE-SKU", "Fixture item", "EA", "BAY-A", qty, _
        description, "Vendor", "V-CODE", "CAT", "EXT", picture, fields)
End Function
'@)
}

function Test-InventoryPrimitivePayload {
    # Literal expected wire content, independent of the production serializer.
    $prefix='[{"System_Key":"TEST-KEY","SKU":"FIXTURE-SKU","Qty":'
    $middle=',"Location":"BAY-A","Condition":"GOOD","AttributesJson":"","Note":"Admin add inventory item","IoType":"CREATE","ITEM_CODE":"FIXTURE-SKU","ITEM":"Fixture item","UOM":"EA","TOTAL INV":'
    $tail=',"VENDOR(s)":"Vendor","VENDOR_CODE":"V-CODE","CATEGORY":"CAT","CATALOG_STATE":"ACTIVE","EXTERNAL_CODE":"EXT","IMAGE_PATH":'
    $expected=@{
        NoCustomFields=($prefix+'0'+$middle+'0,"QtyAvailable":0,"DESCRIPTION":"Fixture description"'+$tail+'""}]')
        ScalarsAndPictures=($prefix+'1.25'+$middle+'1.25,"QtyAvailable":1.25,"DESCRIPTION":"First\nSecond\t\"quoted\" '+[char]937+'"'+$tail+'"one.png","IMAGE_PATH_2":"two.png","TRACK_QTY":"TRUE","ITEM_KIND":"STOCK","CUSTOM_TEXT":"back\\slash","CUSTOM_NUMBER":12.5,"CUSTOM_BOOL":true,"CUSTOM_NULL":null,"CUSTOM_EMPTY":null}]')
        NonCounted=($prefix+'0'+$middle+'0,"QtyAvailable":0,"DESCRIPTION":"Fixture description"'+$tail+'"","TRACK_QTY":"FALSE","ITEM_KIND":"UTILITY","CUSTOM_TEXT":"retained extension"}]')
    }
    foreach($case in @('NoCustomFields','ScalarsAndPictures','NonCounted')){
        $actual=[string](Run 'invSys.Admin.xlam' 'modAdmin.InventoryPrimitivePayloadForTest' @($case))
        if([string]::IsNullOrEmpty($actual)){throw 'Existing Admin payload builder unavailable; not architectural RED.'}
        Check ('InventoryPrimitivePayload.'+$case+'.ExistingWirePreserved') ($actual -ceq $expected[$case])
    }
}
