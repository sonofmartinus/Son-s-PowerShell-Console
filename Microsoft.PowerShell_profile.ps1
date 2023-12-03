<#
.SYNOPSIS
	PowerShell Profile 

.DESCRIPTION
	This is a highly customized PowerShell profile script for doing management tasks in Active Directory and Exchange Online

.EXAMPLE
	Connecting to Exchange Online console using alias

	exoc -UserPrincipalName $adminUPN

.EXAMPLE
	
   
.NOTES
	Version:        1.0
	Author:         Richard Martinez
	Blog:			https://sonofmartinus.com
	Creation Date:  12/3/2023
	
.LINK
	
#>

# Check for latest PowerShell
$PSVersion = $PSVersionTable.PSVersion.Major
if ($PSVersion -lt 7) {
    Write-Output "Old PowerShell version detected. Updating to latest version."
    # Install latest PowerShell using winget
    winget install --id=Microsoft.PowerShell -e
} else {
    Write-Output "Latest PowerShell version already installed."
}

# Check for WindowsCompatibility module
if (!(Get-Module -ListAvailable -Name WindowsCompatibility)) {
    Write-Output "WindowsCompatibility module not found. Installing WindowsCompatibility module."
    # Install WindowsCompatibility module
    Install-Module -Name WindowsCompatibility
} else {
    Write-Output "WindowsCompatibility module already installed."
}

# Check for RSAT
if (!(Get-WindowsCapability -Online | Where-Object {$_.Name -like "*Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0*" -and $_.State -eq "Installed"})) {
    Write-Output "RSAT not found. Installing RSAT."
    # Install RSAT
    Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
} else {
    Write-Output "RSAT already installed."
}

# Check for ExchangeOnlineManagement module
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Output "ExchangeOnlineManagement module not found. Installing ExchangeOnlineManagement module."
    # Install ExchangeOnlineManagement module
    Install-Module -Name ExchangeOnlineManagement
} else {
    Write-Output "ExchangeOnlineManagement module already installed."
}


#First Import some basic modules
Import-Module -Name WindowsCompatibility
Import-Module -Name ActiveDirectory -UseWindowsPowerShell
Import-Module -Name ExchangeOnlineManagement

#Some default variables
$adminUPN= "admin@upn.com"

#Aliases for frequently used commands
Set-Alias im Import-Module
Set-Alias tn Test-NetConnection
Set-Alias exoc Connect-ExchangeOnline
Set-Alias exod Disconnect-ExchangeOnline -Confirm:$false