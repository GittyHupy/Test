
<#
  # Connect to Zeiterfassung Site
  Connect-PnPOnline -Url https://foerdech.sharepoint.com/zeiterfassung -UseWebLogin # limbo@foerde.ch
  function Create_Mitarbeiterliste {
    ## Create Mitarbeiterliste
  New-PnPList     -Title      "Mitarbeiterliste" -Template GenericList    
  Set-PnPList     -Identity   "Mitarbeiterliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

    ## Rename Title Column
  Set-PnPField    -List       "Mitarbeiterliste" -Identity "Title" -Values @{Title="email";Description="auticon Email Adresse, wird zum setzen der Sharepoint Berechtigungen benötigt"}

    ## Adding Columns
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Name"                      -InternalName "Name"            -Type Text      -AddToDefaultView
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Name"                      -Values @{Description="Name des Mitarbeiters"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Pensum"                    -InternalName "Pensum"          -Type Text      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Pensum"                    -Values @{Description="Arbeitspensum des Mitarbeiters"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Bemerkung"                 -InternalName "Bemerkung"       -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Bemerkung"                 -Values @{Description="Notizen zu Pensum Ferientage uvm."}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Ferientage"                -InternalName "Ferientage"      -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Ferientage"                -Values @{Description="Anzahl Ferien Tage im aktuellen Jahr"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Ferientage_inStunden"      -InternalName "Ferientage_inStunden"    -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Ferientage_inStunden"      -Values @{Description="Berechnet mit Pensum in Stunden"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "FerienSaldo_LastMonth"     -InternalName "FerienSaldo_LastMonth"   -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "FerienSaldo_LastMonth"     -Values @{Description="Der Ferien Saldo am letzten Monatsende"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Ferien_ThisMonth"          -InternalName "Ferien_ThisMonth"        -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Ferien_ThisMonth"          -Values @{Description="Ferien diesen Monat"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "StundenSaldo_LastMonth"    -InternalName "StundenSaldo_LastMonth"  -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "StundenSaldo_LastMonth"    -Values @{Description="Stundensaldo am Ende des letzten Monats"}
  Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Stunden_ThisMonth"         -InternalName "Stunden_ThisMonth"       -Type Note      -AddToDefaultView 
  Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Stunden_ThisMonth"         -Values @{Description="Stunden diesen Monat"}  

    ## Create View "Ansicht"
  Add-PnPView     -List       "Mitarbeiterliste" -Title "Ansicht 1" -Fields "email","Name" ,"Bemerkung", "Ferientage", "Ferientage_inStunden", "FerienSaldo_LastMonth", "Ferien_ThisMonth", "StundenSaldo_LastMonth", "Stunden_ThisMonth"   -SetAsDefault  
}

Create_Mitarbeiterliste

Add-PnPField    -List       "Kundenliste" -DisplayName  "Kundennummer4"  -InternalName "Kundennummer4" -Type Text   -AddToDefaultView | Format-List  Title, FieldTypeKind, InternalName, Description, ID, Required, AutoIndexed , SchemaXml
Add-PnPField    -List       "Kundenliste" -DisplayName  "Kundennummer4"  -InternalName "Kundennummer4" -Type Text   -AddToDefaultView | Out-Null
Set-PnPField    -List       "Kundenliste" -Identity     "Kundennummer4"  -Values @{Description="Interne Kunden Nummer (Stephan)"} | Format-List  Title, FieldTypeKind, InternalName, Description, ID, Required, AutoIndexed , SchemaXml

Get-PnPField -List "Kundenliste" -Identity "Kundennummer3" | Format-List  Title, InternalName, FieldTypeKind, Description, ID, Required, AutoIndexed , SchemaXml
Get-PnPField -List "Kundenliste" -Identity "Kundennummer3" | Format-List  Title, SchemaXml

Get-PnPList -Identity "Kundenliste" | Get-ChildItem

#>

$column = "Kundenliste"
Write-Host $column
ForEach ($field in $column.Fields)
{
    Write-Host $field + "  " $column
    Get-PnPField -List "Kundenliste" -Identity $field | Format-List  Title, SchemaXml | Out-File .\log.txt
   # $fieldSettings = Get-PnPField -List "Kundenliste" -Identity $field | Format-List  Title, SchemaXml
   # Write-Output $fieldSettings
   
}


#Read more: https://www.sharepointdiary.com/2016/04/get-list-fields-in-sharepoint-using-powershell.html#ixzz6hfreeNkP
Write-Host "OK"

<#region Testing stuff

Get-PnPField -List Typenliste -Identity "Title" | Format-list ID, Title, InternalName, Description, Scope, TypeShortDescription, SchemaXML
Get-PnPField -List Typenliste -Identity "Title" | Format-list *

Remove-PnPList -Identity "Typenliste" -Force
Remove-PnPList -Identity "Kundenliste" -Force
Remove-PnPList -Identity Auftragsliste -Force
Remove-PnPList -Identity Mitarbeiterliste -Force
Remove-Item Function:\Create_LookupColumn 
Remove-Item Function:\Create_Auftragsliste




function Create_LookupColumn {
    [CmdletBinding()]
      param ($tableName,$fieldName, $lookupTableName 
      )
      $tableName = Get-PnPList -Identity $tableName 
      $fieldName = Get-PnPField -List $tableName -Identity $fieldName
      $lookupTableName = Get-PnPList -Identity $lookupTableName
            
      ## SchemaXML String for the Kunde Lookup Column 
      $schemaXML = '<Field Type="Lookup" DisplayName="Kunde" Description="Kunde aus Kundenliste" Required="FALSE" EnforceUniqueValues="FALSE" List="{'+$lookupTableName.Id+'}" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{'+$fieldName.Id+'}" SourceID="{'+$tableName.Id+'}" StaticName="Kunde" Name="Kunde" ColName="int1" RowOrdinal="0" />'   
  
      return $schemaXML
  }
#>
  

