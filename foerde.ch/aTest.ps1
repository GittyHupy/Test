## aTest2.ps1 file
##region Begin

    ## Connect to Zeiterfassung Site
    Connect-PnPOnline -Url https://foerdech.sharepoint.com/zeiterfassung -UseWebLogin # limbo@foerde.ch

        ## Write-Verbose Status
    $VerbosePreference = "Continue" # "SilentlyContinue", "Stop", "Continue", "Inquire", "Ignore", "Suspend"
    Write-Host "VerbosePreference is set to: " $VerbosePreference
 
       ## Write-Debug Status
    $DebugPreference = "Continue" # "SilentlyContinue", "Stop", "Continue", "Inquire", "Ignore", "Suspend"
    Write-Host "DebugPreference is set to: " $DebugPreference

##endregion

function Create_LookupColumn {
    [CmdletBinding()]
      param ($tableName,$fieldName, $lookupTableName 
      )
      $tableName = Get-PnPList -Identity $tableName 
      $fieldName = Get-PnPField -List $tableName -Identity $fieldName
      $lookupTableName = Get-PnPList -Identity $lookupTableName
      
       Write-Host $fieldName | Select-Object *
   # Write-Debug $fieldName |Select-Object *
      
  
      ## SchemaXML String for the Kunde Lookup Column 
      $schemaXML = '<Field Type="Lookup" DisplayName="'+$fieldName.Title+'" Description="'+$fieldName.Description+'" Required="FALSE" EnforceUniqueValues="FALSE" List="{'+$lookupTableName.Id+'}" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="None" ID="{'+$fieldName.Id+'}" SourceID="{'+$tableName.Id+'}" StaticName="'+$fieldName.Title+'" Name="'+$fieldName.Title+'" ColName="int1" RowOrdinal="0" />'   
  
      Write-debug -Message halt 
                                     
      return $schemaXML
}

Create_LookupColumn -tableName "Zeiterfassungsliste" -fieldName "Title" -lookupTableName "Auftragsliste" 

Write-Verbose "halloE"


<#
$tableName = "Zeiterfassungsliste"; $fieldName= "Title"
Get-PnPField -List $tableName -Identity $fieldName.Title | Select-Object * | Sort-Object Name
#>


