# Script to Manage Zeiterfassung in Sharepoint with Powershell
<# Tools used 
   # Powershell PSVersion 5.1.19041.610         \> $PSVersionTable
     # Module 
       # PowerShellGet 2.2.5                    \> Get-Module -Name *ShellGet*  -ListAvailable | Select Name,Version
       # PnP.PowerShell 0.3.14                  \> Get-Module -Name *pnp*       -ListAvailable | Select Name,Version
       # ImportExcel 7.1.1                      \> Get-Module -Name *Excel*     -ListAvailable | Select Name,Version
#>
#region Sharepoint Conncetions

  # Connect to foerde.ch Sharepoint 
Connect-PnPOnline -Url https://foerdech.sharepoint.com/ -UseWebLogin          # limbo@foerde.ch
 
  # Connect to support-hub.ch Sharepoint 
Connect-PnPOnline -Url https://supporthubch.sharepoint.com/sites/Zeiterfassen -UseWebLogin      # Malim.zurich@support-hub.ch 
Connect-SPOService -Url https://supporthubch.sharepoint.com/sites/Zeiterfassen                  # Malim.zurich@support-hub.ch 
#endregion

  # Connect to Zeiterfassung Site
Connect-PnPOnline -Url https://foerdech.sharepoint.com/zeiterfassung -UseWebLogin # limbo@foerde.ch

  # Connect to Zeiterfassung2 Site
Connect-PnPOnline -Url https://foerdech.sharepoint.com/sites/zeiterfassung2 -UseWebLogin # limbo@foerde.ch 

  # Connect to Zeiterfassung_app Site
Connect-PnPOnline -Url https://foerdech.sharepoint.com/sites/Zeiterfassung2/Zeiterfassung_app/  -UseWebLogin # limbo@foerde.ch 

#region Discovery
Get-PnPList -List Contract_List

Get-PnPField -List Auftragsliste -Identity "Kunde" | Get-Member

  # Info about a Lookup Fieldtype
Get-PnPField -List Auftragsliste -Identity "Kunde" | Format-List  Title, ID, FieldTypeKind, Description, SchemaXml 

Get-PnPField -List Auftragsliste -Identity "Kunde" | Format-List  Title, ID, FieldTypeKind, Description, SchemaXml
#endregion


