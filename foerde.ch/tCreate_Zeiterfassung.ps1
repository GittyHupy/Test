## Script to setup Zeiterfassung in Sharepoint with Powershell
## Tools used 
   ## Powershell PSVersion 5.1.19041.610         \> $PSVersionTable
     ## Module 
       ## PowerShellGet 2.2.5                    \> Get-Module -Name *ShellGet*  -ListAvailable | Select Name,Version
       ## PnP.PowerShell 0.3.14                  \> Get-Module -Name *pnp*       -ListAvailable | Select Name,Version
       ## ImportExcel 7.1.1                      \> Get-Module -Name *Excel*     -ListAvailable | Select Name,Version


$VerbosePreference = "Continue" # "SilentlyContinue", "Stop", "Continue", "Inquire", "Ignore", "Suspend"
Write-Host "Write-Verbose Status:" $VerbosePreference

 
  ## Connect to new Site 
Connect-PnPOnline -Url https://foerdech.sharepoint.com/zeiterfassung -UseWebLogin


function Create_Kundenliste {
  Write-Verbose "Function Create_Kundenliste called"##
$item = New-PnPList     -Title      "Kundenliste" -Template GenericList       
$item = Set-PnPList     -Identity   "Kundenliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

  Write-Verbose "Rename Title Column"## 
$item = Set-PnPField    -List       "Kundenliste" -Identity "Title"             -Values @{Title="Kundenname"; Description="In der Liste sind die Kunden zu erfassen, auf welche rapportiert werden soll."}

  Write-Verbose "Adding Columns"##
$item = Add-PnPField    -List       "Kundenliste" -DisplayName  "Kundennummer"  -InternalName "Kundennummer" -Type Text  -AddToDefaultView 
$item = Set-PnPField    -List       "Kundenliste" -Identity     "Kundennummer"  -Values @{Description="Interne Kunden Nummer (Stephan)"}

  Write-Verbose "Create View 'Ansicht 1'"##
$item = Add-PnPView     -List       "Kundenliste" -Title "Ansicht 1" -Fields "Kundennummer","Kundenname"  -SetAsDefault  

  Write-Verbose "Filling Table"## 
  ## Import Typen into Typenliste 88- original Headername "Title" is necessary instead of "Typ" -88
$Kunden = Import-Excel -Path .\Zeiterfassung_Import-Daten.xlsx -WorksheetName Kundenliste -HeaderName "Title" , "Kundennummer" -StartRow 2
foreach ($kunde in $Kunden)
{
    ## Write-Verbose $kunde -Verbose
    Add-PnPListItem -List "Kundenliste" -Values @{"Title"=$kunde.Title; "Kundennummer"=$kunde.Kundennummer}
}
return
}

function Create_Auftragsliste {
  Write-Verbose "Function Create_Auftragsliste called"##
$item = New-PnPList     -Title      "Auftragsliste" -Template GenericList    
$item = Set-PnPList     -Identity   "Auftragsliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

  Write-Verbose "Rename Title Column"##
$item = Set-PnPField    -List       "Auftragsliste" -Identity "Title" -Values @{Title="Auftrag";Description="Kurze Auftragsbezeichung"}

  Write-Verbose "Adding & Configuring Lookup Column"##
$item = Add-PnPField    -List       "Auftragsliste" -DisplayName  "Kunde"             -InternalName "Kunde"               -Type Lookup    -AddToDefaultView 
$schemaXML = Create_LookupColumn -tableName "Auftragsliste" -fieldName "Kunde" -lookupTableName "Kundenliste" -showField "Title"

  Write-Verbose "Adding & Configuring Column"##
$item = Set-PnPField    -List       "Auftragsliste" -Identity     "Kunde"             -Values @{SchemaXml=$schemaXML}
$item = Add-PnPField    -List       "Auftragsliste" -DisplayName  "Auftragsnummer"    -InternalName "Auftragsnummer"      -Type Text      -AddToDefaultView  
$item = Set-PnPField    -List       "Auftragsliste" -Identity     "Auftragsnummer"    -Values @{Description="Interne Kunden Nummer (Stephan)"}
$item = Add-PnPField    -List       "Auftragsliste" -DisplayName  "Rechnungskontakt"  -InternalName "Rechnungskontakt"    -Type Note      -AddToDefaultView  
$item = Set-PnPField    -List       "Auftragsliste" -Identity     "Rechnungskontakt"  -Values @{Description="Name und Adresse Rechnungskontakt"}
$item = Add-PnPField    -List       "Auftragsliste" -DisplayName  "Rechnungstext"     -InternalName "Rechnungstext"       -Type Note      -AddToDefaultView  
$item = Set-PnPField    -List       "Auftragsliste" -Identity     "Rechnungstext"     -Values @{Description="Kundenspezifischer Rechnungstext"}

  Write-Verbose "Create View 'Ansicht 1'"##
$item = Add-PnPView     -List       "Auftragsliste" -Title "Ansicht 1" -Fields "Auftrag","Kunde" ,"Auftragsnummer", "Rechnungskontakt", "Rechnungstext" -SetAsDefault
return  
}

function Create_Typenliste {
Write-Verbose "Function Create_Typenliste called"##
$item = New-PnPList     -Title      "Typenliste" -Template GenericList    
$item = Set-PnPList     -Identity   "Typenliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

Write-Verbose "Rename Title Column"##
$item = Set-PnPField    -List       "Typenliste" -Identity    "Title"     -Values @{Title="Typ";Description="Abrechnungsart"}

Write-Verbose "Adding & Configuring Columns"##
$item = Add-PnPField    -List       "Typenliste" -DisplayName "Beschrieb" -InternalName "Beschrieb"      -Type Text      -AddToDefaultView 
$item = Set-PnPField    -List       "Typenliste" -Identity    "Beschrieb" -Values @{Description="Erklärung des Abrechnungstypen"}

Write-Verbose "Create View 'Ansicht 1'"##
$item = Add-PnPView     -List       "Typenliste" -Title "Ansicht 1" -Fields "Typ", "Beschrieb" -SetAsDefault  

Write-Verbose "Filling Table"##
## Import Typen into Typenliste 88- original Headername "Title" is necessary instead of "Typ" -88
$Typen = Import-Excel -Path .\Zeiterfassung_Import-Daten.xlsx -WorksheetName Typenliste -HeaderName "Title" , "Beschrieb" -StartRow 2
foreach ($typ in $Typen)
{
  ## Write-Verbose $typ -Verbose  
  Add-PnPListItem -List "Typenliste" -Values @{"Title"=$typ.Title; "Beschrieb"=$typ.Beschrieb}
}
return 
}

function Create_Arbeitsorteliste {
Write-Verbose "Function Create_Arbeitsorteliste called"##
$item = New-PnPList     -Title      "Arbeitsorteliste" -Template GenericList    
$item = Set-PnPList     -Identity   "Arbeitsorteliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

Write-Verbose "Rename Title Column"##
$item = Set-PnPField    -List       "Arbeitsorteliste" -Identity    "Title" -Values @{Title="Arbeitsort";Description="Ort an dem gearbeitet wird"}

Write-Verbose "Adding & Configuring Columns"##
$item = Add-PnPField    -List       "Arbeitsorteliste" -DisplayName "Beschrieb"   -InternalName "Beschrieb"           -Type Text      -AddToDefaultView 
$item = Set-PnPField    -List       "Arbeitsorteliste" -Identity    "Beschrieb"   -Values @{Description="Beschrieb Arbeitsort"}

Write-Verbose "Create View 'Ansicht 1'"## 
$item = Add-PnPView     -List       "Arbeitsorteliste" -Title       "Ansicht 1"   -Fields "Arbeitsort", "Beschrieb" -SetAsDefault  

Write-Verbose "Filling Table"## 
## Import Arbeitsorte into Arbeitsorteliste 88- original Headername "Title" is necessary instead of "Arbeitsort" -88
$ArbeitsOrte = Import-Excel -Path .\Zeiterfassung_Import-Daten.xlsx -WorksheetName Arbeitsorteliste -HeaderName "Title" , "Beschrieb" -StartRow 2
foreach ($arbeitsOrt in $ArbeitsOrte)
{
  ## Write-Verbose $arbeitsOrt -Verbose
Add-PnPListItem -List "Arbeitsorteliste" -Values @{"Title"=$arbeitsOrt.Title; "Beschrieb"=$arbeitsOrt.Beschrieb}
}
return 
}

function Create_Mitarbeiterliste {
Write-Verbose "Function Create_Mitarbeiterliste called"## 
$item = New-PnPList     -Title      "Mitarbeiterliste" -Template GenericList    
$item = Set-PnPList     -Identity   "Mitarbeiterliste" -EnableVersioning $true -MajorVersions 10 -EnableAttachments $false

Write-Verbose "Rename Title Column"##
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity "Title" -Values @{Title="email";Description="auticon Email Adresse, wird zum setzen der Sharepoint Berechtigungen benötigt"}

Write-Verbose "Adding & Configuring Columns"##
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Name"                      -InternalName "Name"            -Type Text      -AddToDefaultView 
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Name"                      -Values @{Description="Name des Mitarbeiters"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Pensum"                    -InternalName "Pensum"          -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Pensum"                    -Values @{Description="Arbeitspensum des Mitarbeiters"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Bemerkung"                 -InternalName "Bemerkung"       -Type Note      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Bemerkung"                 -Values @{Description="Notizen zu Pensum Ferientage uvm."}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "UPN"                       -InternalName "UPN"       -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "UPN"                       -Values @{Description="User Principle Name für Sharepoint Berechtigung"}  
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Ferientage"                -InternalName "Ferientage"      -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Ferientage"                -Values @{Description="Anzahl Ferien Tage im aktuellen Jahr"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Ferientage_inStunden"      -InternalName "Ferientage_inStunden"    -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Ferientage_inStunden"      -Values @{Description="Berechnet mit Pensum in Stunden"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "FerienSaldo_LastMonth"     -InternalName "FerienSaldo_LastMonth"   -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "FerienSaldo_LastMonth"     -Values @{Description="Der Ferien Saldo am letzten Monatsende"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Ferien_ThisMonth"          -InternalName "Ferien_ThisMonth"        -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Ferien_ThisMonth"          -Values @{Description="Ferien diesen Monat"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "StundenSaldo_LastMonth"    -InternalName "StundenSaldo_LastMonth"  -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "StundenSaldo_LastMonth"    -Values @{Description="Stundensaldo am Ende des letzten Monats"}
$item = Add-PnPField    -List       "Mitarbeiterliste" -DisplayName   "Stunden_ThisMonth"         -InternalName "Stunden_ThisMonth"       -Type Number      -AddToDefaultView  
$item = Set-PnPField    -List       "Mitarbeiterliste" -Identity      "Stunden_ThisMonth"         -Values @{Description="Stunden diesen Monat"}  

Write-Verbose "Create View 'Ansicht 1'"## 
$item = Add-PnPView     -List       "Mitarbeiterliste" -Title "Ansicht 1" -Fields "email","Name" ,"Pensum", "Bemerkung", "UPN", "Ferientage", "Ferientage_inStunden", "FerienSaldo_LastMonth", "Ferien_ThisMonth", "StundenSaldo_LastMonth", "Stunden_ThisMonth"   -SetAsDefault  

Write-Verbose "Filling Table"## 
## Import Mitarbeiter into Mitarbeiterliste 88- original Headername "Title" is necessary instead of "Email" -88
$MitarbeiterList = Import-Excel -Path .\Zeiterfassung_Import-Daten.xlsx -WorksheetName Mitarbeiterliste -HeaderName "Title" , "Name" , "Pensum" -StartRow 2
foreach ($mitarbeiter in $MitarbeiterList)
{
  ## Write-Verbose $mitarbeiter -Verbose
  Add-PnPListItem -List "Mitarbeiterliste" -Values @{"Title"=$mitarbeiter.Title; "Name"=$mitarbeiter.Name; "Pensum"=$mitarbeiter.Pensum}
}
return 
}
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
$schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Kunde" -lookupTableName "Kundenliste" -showField "Title"
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Kunde" -Values @{SchemaXml=$schemaXML;Description="Kunde aus Kundenliste zu Kundenname"} 

Write-Verbose "Create & Configuring Auftrag Lookup "##
$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName  "Auftrag"      -InternalName "Auftrag"            -Type Lookup    -AddToDefaultView 
$schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Auftrag" -lookupTableName "Auftragsliste" -showField "Title"
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Auftrag" -Values @{SchemaXml=$schemaXML;Description="Auftragsbezeichnung aus Auftragsliste zu Auftrag"} 

Write-Verbose "Create & Configuring Typ Lookup "##
$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Typ"          -InternalName "Typ"                 -Type Lookup    -AddToDefaultView 
$schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Typ" -lookupTableName "Typenliste" -showField "Title"
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Typ" -Values @{SchemaXml=$schemaXML;Description="Typ aus Typenliste zu Typ"} 

Write-Verbose "Create & Configuring Arbeitsort Lookup "##
$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Arbeitsort"   -InternalName "Arbeitsort"          -Type Lookup    -AddToDefaultView 
$schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Arbeitsort" -lookupTableName "Arbeitsorteliste" -showField "Title"
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Arbeitsort" -Values @{SchemaXml=$schemaXML;Description="Arbeitsort aus Arbeitsorteliste zu Arbeitsort"} 

$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Projekt"        -InternalName "Projekt"             -Type Text      -AddToDefaultView 
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Projekt"        -Values @{Description="Kunden Projektname oder Überbegriff der Arbeit eingeben"}
$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Stunden"        -InternalName "Stunden"             -Type Number  -AddToDefaultView 
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Stunden"        -Values @{Description='Stunden dezimal 4,5 ist 4:30'}

Write-Verbose "Create & Configuring Mitarbeiter Lookup "##
$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Mitarbeiter"   -InternalName "Mitarbeiter"        -Type Lookup      -AddToDefaultView 
$schemaXML = Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Mitarbeiter" -lookupTableName "Mitarbeiterliste" -showField "Name"
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity     "Mitarbeiter" -Values @{SchemaXml=$schemaXML;Description="Mitarbeiter aus Mitarbeiterliste zu Mitarbeiter"} 

$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Pensum"        -InternalName "Pensum"           -Type Number        -AddToDefaultView 
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Pensum" -Values @{Description="Arbeitspensum wird via Workflow aus Mitarbeiterliste Pensum übernommen"}

$item = Add-PnPField    -List       "Zeiterfassungsliste" -DisplayName "Notiz"        -InternalName "Notiz"            -Type Note           -AddToDefaultView 
$item = Set-PnPField    -List       "Zeiterfassungsliste" -Identity    "Notiz" -Values @{Description="Diese Notiz für Consultant gedacht um Lust & Frust zu verarbeiten. Kann vom Job Coach gelesen werden"}

Write-Verbose "Create View 'Ansicht 1'"##
$item = Add-PnPView     -List       "Zeiterfassungsliste" -Title "Ansicht 1" -Fields "Datum", "Dauer", "Kunde", "Auftrag", "Typ", "Arbeitsort", "Projekt", "Arbeitsbeschrieb", "Stunden", "Mitarbeiter", "Pensum", "Notiz"    -SetAsDefault
Write-Verbose "Create View 'Ansicht 2'"##
$item = Add-PnPView     -List       "Zeiterfassungsliste" -Title "Ansicht 2" -Fields "Datum", "Beginn", "Ende", "Dauer", "Kunde", "Auftrag", "Typ", "Arbeitsort", "Projekt", "Arbeitsbeschrieb", "Stunden", "Mitarbeiter", "Pensum", "Notiz"    -SetAsDefault
Write-Verbose "Create View 'IMPORT View'"##
$item = Add-PnPView     -List       "Zeiterfassungsliste" -Title "IMPORT View" -Fields "Datum", "Kunde", "Auftrag", "Typ", "Arbeitsort", "Projekt",  "Arbeitsbeschrieb", "Stunden", "Mitarbeiter"   -SetAsDefault


return  
}

function Create_LookupColumn {
[CmdletBinding()]
    param ($tableName, $fieldName, $lookupTableName, $showField 
)
Write-Verbose  "Configure Lookup for $tableName"
$tableName = Get-PnPList  -Identity $tableName 
$fieldName = Get-PnPField -List $tableName -Identity $fieldName
$lookupTableName = Get-PnPList -Identity $lookupTableName
      
  ## SchemaXML String for the Kunde Lookup Column 
$schemaXML = '<Field Type="Lookup" DisplayName="'+$fieldName.Title+'" Description="'+$fieldName.Description+'" Required="FALSE" EnforceUniqueValues="FALSE" List="{'+$lookupTableName.Id+'}" ShowField="'+$showField+'" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{'+$fieldName.Id+'}" SourceID="{'+$tableName.Id+'}" StaticName="'+$fieldName.Title+'" Name="'+$fieldName.Title+'" ColName="int1" RowOrdinal="0" />'   

return $schemaXML
}

Create_Kundenliste  
Create_Auftragsliste
Create_Typenliste
Create_Arbeitsorteliste
Create_Mitarbeiterliste
Create_Zeiterfassungsliste

<#region Cleanup-Action

Remove-PnPList -Identity "Kundenliste" -Force
Remove-PnPList -Identity "Auftragsliste" -Force
Remove-PnPList -Identity "Typenliste" -Force
Remove-PnPList -Identity "Arbeitsorteliste" -Force
Remove-PnPList -Identity "Mitarbeiterliste" -Force
Remove-PnPList -Identity "Zeiterfassungsliste" -Force

Remove-Item Function:\Create_LookupColumn     -ErrorAction SilentlyContinue
Remove-Item Function:\Create_Auftragsliste    -ErrorAction SilentlyContinue
Remove-Item Function:\Create_Typenliste       -ErrorAction SilentlyContinue
Remove-Item Function:\Create_Arbeitsorteliste -ErrorAction SilentlyContinue
Remove-Item Function:\Create_Mitarbeiterliste -ErrorAction SilentlyContinue

#>

