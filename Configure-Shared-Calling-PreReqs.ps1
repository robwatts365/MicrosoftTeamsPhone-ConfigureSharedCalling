<# 
Configure Shared Calling Pre-Requisites
    Version: v1.0
    Date: 10/01/2024
    Author: Rob Watts - Cloud Solution Architect - Microsoft
    Description: This script will install the pre-requisites needed to configure Shared Calling feature for Microsoft Teams

DISCLAIMER
   THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
   MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES
   OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR
   PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR
   ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS
   INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR
   INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
   BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR
   INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
#>

################### START OF GENERIC SCRIPT FUNCTIONS ###################
function ShowDisclaimer
    {

        # Pop out disclaimer
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [Windows.Forms.MessageBox]::Show("
        THIS CODE IS SAMPLE CODE. 

        THESE SAMPLES ARE PROVIDED 'AS IS' WITHOUT WARRANTY OF ANY KIND.

        MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. 

        THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU.

        IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

        BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.", "***DISCLAIMER***", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)

    }

 ################### END OF GENERIC SCRIPT FUNCTIONS ###################
 
 ################### START OF INSTALL MODULE FUNCTIONS ###################
Function Install-AzureADPreview
    {
    Install-Module AzureADPreview -Force
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-Host "AzureADPreview module installed." -ForegroundColor Green
    }
Function Install-MicrosoftGraph 
    {
    Install-Module Microsoft.Graph
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-Host "Microsoft.Graph module installed." -ForegroundColor Green
    }
Function Install-MSOnline
    {
    Install-Module MSOnline
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-Host "MSOnline module installed." -ForegroundColor Green
    }
Function Install-ExchangeOnlineManagement
    {
    Install-Module ExchangeOnlineManagement
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-Host "ExchangeOnlineManagement module installed." -ForegroundColor Green
    }
Function Install-MicrosoftTeams
    {
    Install-Module MicrosoftTeams
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-Host "MicrosoftTeams module installed." -ForegroundColor Green
    }

################### END OF INSTALL MODULE FUNCTIONS ###################

################### START OF COMMANDS ###################

ShowDisclaimer

Write-Host "Checking for elevated permissions..."
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
} else {
    Write-Host "Script is running as administrator. Moving on..." -ForegroundColor Green
}

# Installs Powershell Modules, if not already installed
Write-Host "Checking for PowerShell Modules" -ForegroundColor Gray -BackgroundColor Black

#AzureADPreview
If (-not(Get-InstalledModule AzureADPreview -ErrorAction silentlycontinue)) {
    Write-Host "AzureADPreview module does not exist" -ForegroundColor DarkRed
    Write-Host "AzureADPreview module will be installed" 
    Install-AzureADPreview
  }
  Else {
    Write-Host "AzureADPreview module exists" -ForegroundColor Green
  }
  
#Microsoft.Graph
  If (-not(Get-InstalledModule Microsoft.Graph -ErrorAction silentlycontinue)) {
    Write-Host "Microsoft.Graph module does not exist" -ForegroundColor DarkRed
    Write-Host "Microsoft.Graph module will be installed"
    Install-MicrosoftGraph
  }
  Else {
    Write-Host "Microsoft.Graph module exists" -ForegroundColor Green
  }
  
#MSOnline
  If (-not(Get-InstalledModule MSOnline -ErrorAction silentlycontinue)) {
    Write-Host "MSOnline module does not exist" -ForegroundColor DarkRed
    Write-Host "MSOnline module will be installed"
    Install-MSOnline
  }
  Else {
    Write-Host "MSOnline module exists" -ForegroundColor Green
  }
  
#ExchangeOnlineManagement
  If (-not(Get-InstalledModule ExchangeOnlineManagement -ErrorAction silentlycontinue)) {
    Write-Host "ExchangeOnlineManagement module does not exist" -ForegroundColor DarkRed
    Write-Host "ExchangeOnlineManagement module will be installed"
    Install-ExchangeOnlineManagement
  }
  Else {
    Write-Host "ExchangeOnlineManagement module exists" -ForegroundColor Green
  }

#MicrosoftTeams
  If (-not(Get-InstalledModule MicrosoftTeams -ErrorAction silentlycontinue)) {
    Write-Host "MicrosoftTeams module does not exist" -ForegroundColor DarkRed
    Write-Host "MicrosoftTeams module will be installed"
    Install-MicrosoftTeams
  }
  Else {
    Write-Host "MicrosoftTeams module exists" -ForegroundColor Green
  } 
################### END OF COMMANDS ###################