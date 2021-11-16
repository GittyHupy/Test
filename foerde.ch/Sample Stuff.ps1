## Sample Stuff

  ## Write-Verbose Status
  $VerbosePreference = "Continue" # "SilentlyContinue", "Stop", "Continue", "Inquire", "Ignore", "Suspend"

 ## MessageBox
[System.Windows.Forms.MessageBox]::Show("Halo velo" + $PSVersionTable  ,"Titel",0)

Get-PnPField -List Typenliste -Identity "Title" | Get-Member


dir function:*