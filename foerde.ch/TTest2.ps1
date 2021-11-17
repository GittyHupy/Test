## TTest2 file
##region Beginn

  ## Connect to Zeiterfassung Site
Connect-PnPOnline -Url https://foerdech.sharepoint.com/zeiterfassung -UseWebLogin # limbo@foerde.ch

  ## Write-Verbose Status
$VerbosePreference = "Continue" # "SilentlyContinue", "Stop", "Continue", "Inquire", "Ignore", "Suspend"
Write-Host "VerbosePreference is set to: $VerbosePreference"

##endregion 


function Create_Zeiterfassungsliste {
    Write-Verbose "Function Create_Zeiterfassungsliste called"##
  $item = New-PnPList     -Title      "Zeiterfassungsliste" -Template GenericList    
  $item = Set-PnPList     -Identity   "Zeiterfassungsliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

    Write-Verbose "Rename Title Column"##
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity "Title" -Values @{Title="Arbeitsbeschrieb";Description="Kurzbeschrieb der erledigten Arbeit"}

    Write-Verbose "Adding Columns"##
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Datum"        -InternalName "Datum"              -Type DateTime  -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Datum" -Values @{Description="Arbeitsdatum"}
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Beginn"       -InternalName "Beginn"             -Type DateTime  -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Beginn" -Values @{Description="Arbeitsbeginn"}
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Ende"         -InternalName "Ende"               -Type DateTime  -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Ende" -Values @{Description="Arbeitsende"}
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Dauer"        -InternalName "Dauer"              -Type DateTime  -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Dauer" -Values @{Description="Arbeitsdauer in hh:mm"}

    Write-Verbose "Create & Configuring Kunde Lookup "##
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName  "Kunde"        -InternalName "Kunde"              -Type Lookup    -AddToDefaultView 
  $schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Kunde" -lookupTableName "Kundenliste" 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Kunde" -Values @{SchemaXml=$schemaXML;Description="Kunde aus Kundenliste zu Kundenname"} 

    Write-Verbose "Create & Configuring Auftrag Lookup "##
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName  "Auftrag"      -InternalName "Auftrag"            -Type Lookup    -AddToDefaultView 
  $schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Auftrag" -lookupTableName "Auftragsliste" 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Auftrag" -Values @{SchemaXml=$schemaXML;Description="Auftragsbezeichnung aus Auftragsliste zu Auftrag"} 
 
    Write-Verbose "Create & Configuring Typ Lookup "##
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Typ"          -InternalName "Typ"                 -Type Lookup    -AddToDefaultView 
  $schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Typ" -lookupTableName "Typenliste" 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Typ" -Values @{SchemaXml=$schemaXML;Description="Typ aus Typenliste zu Typ"} 

   Write-Verbose "Create & Configuring Arbeitsort Lookup "##
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Arbeitsort"   -InternalName "Arbeitsort"          -Type Lookup    -AddToDefaultView 
  $schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Arbeitsort" -lookupTableName "Arbeitsorteliste" 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Arbeitsort" -Values @{SchemaXml=$schemaXML;Description="Arbeitsort aus Arbeitsorteliste zu Arbeitsort"} 

  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Projekt"        -InternalName "Projekt"           -Type Text      -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Projekt" -Values @{Description="Kunden Projektname oder Überbegriff der Arbeit eingeben"}
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Stunden"        -InternalName "Stunden"           -Type DateTime  -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Stunden" -Values @{Description='Stunden dezimal 4,5 ist 4:30'}

    Write-Verbose "Create & Configuring Mitarbeiter Lookup "##
  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Mitarbeiter"   -InternalName "Mitarbeiter"        -Type Lookup      -AddToDefaultView 
  $schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Mitarbeiter" -lookupTableName "Mitarbeiterliste" 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Mitarbeiter" -Values @{SchemaXml=$schemaXML;Description="Mitarbeiter aus Mitarbeiterliste zu Mitarbeiter"} 

  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Pensum"        -InternalName "Pensum"           -Type Zahl          -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Pensum" -Values @{Description="Arbeitspensum wird via Workflow aus Mitarbeiterliste Pensum übernommen"}

  $item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Notiz"        -InternalName "Notiz"            -Type Note           -AddToDefaultView 
  $item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Notiz" -Values @{Description="Diese Notiz für Consultant gedacht um Lust & Frust zu verarbeiten. Kann vom Job Coach gelesen werden"}

    Write-Verbose "Create View 'Ansicht 1'"##
  $item = Add-PnPView     -List       "Zeiterfassungsliste" -Title "Ansicht 1" -Fields "Datum", "Dauer", "Kunde", "Auftrag", "Typ", "Arbeitsort", "Projekt", "Stunden", "Arbeitsbeschrieb", "Mitarbeiter", "Pensum", "Notiz"    -SetAsDefault
    Write-Verbose "Create View 'Ansicht 2'"##
  $item = Add-PnPView     -List       "Zeiterfassungsliste" -Title "Ansicht 2" -Fields "Datum", "Beginn", "Ende", "Dauer", "Kunde", "Auftrag", "Typ", "Arbeitsort", "Projekt", "Stunden", "Arbeitsbeschrieb", "Mitarbeiter", "Pensum", "Notiz"    -SetAsDefault

  return  
}

function Create_LookupColumn {
  [CmdletBinding()]
    param ($tableName,$fieldName, $lookupTableName 
    )
    $tableName = Get-PnPList -Identity $tableName 
    $fieldName = Get-PnPField -List $tableName -Identity $fieldName
    $lookupTableName = Get-PnPList -Identity $lookupTableName
    
      ## SchemaXML String for the Kunde Lookup Column 
    $schemaXML = '<Field Type="Lookup" DisplayName="'+$fieldName.Title+'" Description="'+$fieldName.Description+'" Required="FALSE" EnforceUniqueValues="FALSE" List="{'+$lookupTableName.Id+'}" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{'+$fieldName.Id+'}" SourceID="{'+$tableName.Id+'}" StaticName="'+$fieldName.Title+'" Name="'+$fieldName.Title+'" ColName="int1" RowOrdinal="0" />'   
                                  
    return $schemaXML
}

Create_Zeiterfassungsliste


