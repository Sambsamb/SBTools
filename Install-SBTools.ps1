$Destination = "C:\Program Files\WindowsPowerShell\Modules\SBTools"
If ( -not (Test-Path -Path $Destination)) {
    New-Item -Path $Destination -ItemType directory -force -Confirm:$false
}
Copy-Item -Path "SBTools.psm1","SBTools.psd1" -Destination $Destination -Force -Confirm:$false
Unblock-File "$Destination\SBTools.psm1","$Destination\SBTools.psd1" -Confirm:$false
Remove-Module SBTools -ErrorAction SilentlyContinue
Import-Module SBTools -DisableNameChecking
Get-Module
Get-Command -Module SBTools