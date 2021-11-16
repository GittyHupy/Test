# Connect to new Site
Connect-PnPOnline -Url https://foerdech.sharepoint.com/zeiterfassung 

# New-PnPList -Title "Kundenliste" -Template GenericList 
Add-PnPField -list "Kundenliste" -Type Text -DisplayName "Kundennummer" -InternalName "Kundennummer" 
Add-PnPField -list "Kundenliste" -Type Text -DisplayName "Kundennummer" -InternalName "Kundennummer" 