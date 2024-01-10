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

### FUNCTION - Write To Log File ###
function Write-LogFileMessage($message) 
    {
        $ActionTime = $(Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
        Write-Debug "$ActionTime :: $message"
    "$ActionTime :: $message" | Out-File -FilePath $LogFile -Append
    }

### FUNCTION - Progress Bar Function ###
Function Set-ScriptSleep($seconds)
    {
        $s = 0;
        Do {
            $p = [math]::Round(100 - (($seconds - $s) / $seconds * 100));
            Write-Progress -Activity "Waiting..." -Status "$p% Complete:" -SecondsRemaining ($seconds - $s) -PercentComplete $p;
            [System.Threading.Thread]::Sleep(1000)
            $s++;
        }
        While($s -lt $seconds); 
    }
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
    Install-Module AzureADPreview
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-LogFileMessage "AzureADPreview Module Installed"
    }
Function Install-MicrosoftGraph 
    {
    Install-Module Microsoft.Graph
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-LogFileMessage "MicrosoftGraph Module Installed"
    }
Function Install-MSOnline
    {
    Install-Module MSOnline
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-LogFileMessage "MSOnline Module Installed"
    }
Function Install-ExchangeOnlineManagement
    {
    Install-Module ExchangeOnlineManagement
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-LogFileMessage "ExchangeOnlineManagement Module Installed"
    }
Function Install-MicrosoftTeams
    {
    Install-Module MicrosoftTeams
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Write-LogFileMessage "MicrosoftTeams Module Installed"
    }

################### END OF INSTALL MODULE FUNCTIONS ###################

################### START OF COMMANDS ###################

# Enable File Saver for Log File
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# Set Log File Location 
Write-Host "Please select the Log file location."
$LogSaver = New-Object -Typename System.Windows.Forms.SaveFileDialog
$LogSaver.initialDirectory = $initialDirectory
$LogSaver.filter = "All files (*.log)| *.log"
$LogSaver.ShowDialog() | Out-Null
$LogFile = $LogSaver.filename

ShowDisclaimer

# Installs Powershell Modules, if not already installed
Write-Host "Checking for PowerShell Modules" -ForegroundColor Gray -BackgroundColor Black
Write-LogFileMessage "Checking for PowerShell Modules"
#AzureADPreview
If (-not(Get-InstalledModule AzureADPreview -ErrorAction silentlycontinue)) {
    Write-Host "AzureADPreview module does not exist"
    Write-LogFileMessage "AzureADPreview module does not exist"
    Install-AzureADPreview
  }
  Else {
    Write-Host "AzureADPreview module exists"
    Write-LogFileMessage "AzureADPreview module exists"
  }
  
#Microsoft.Graph
  If (-not(Get-InstalledModule Microsoft.Graph -ErrorAction silentlycontinue)) {
    Write-Host "Microsoft.Graph module does not exist"
    Write-LogFileMessage "Microsoft.Graph module does not exist"
    Install-MicrosoftGraph
  }
  Else {
    Write-Host "Microsoft.Graph module exists"
    Write-LogFileMessage "Microsoft.Graph module exists"
  }
  
#MSOnline
  If (-not(Get-InstalledModule MSOnline -ErrorAction silentlycontinue)) {
    Write-Host "MSOnline module does not exist"
    Write-LogFileMessage "MSOnline module does not exist"
    Install-MSOnline
  }
  Else {
    Write-Host "MSOnline module exists"
    Write-LogFileMessage "MSOnline module exists"
  }
  
#ExchangeOnlineManagement
  If (-not(Get-InstalledModule ExchangeOnlineManagement -ErrorAction silentlycontinue)) {
    Write-Host "ExchangeOnlineManagement module does not exist"
    Write-LogFileMessage "ExchangeOnlineManagement module does not exist"
    Install-ExchangeOnlineManagement
  }
  Else {
    Write-Host "ExchangeOnlineManagement module exists"
    Write-LogFileMessage "ExchangeOnlineManagement module exists"
  }

#MicrosoftTeams
  If (-not(Get-InstalledModule MicrosoftTeams -ErrorAction silentlycontinue)) {
    Write-Host "MicrosoftTeams module does not exist"
    Write-LogFileMessage "MicrosoftTeams module does not exist"
    Install-MicrosoftTeams
  }
  Else {
    Write-Host "MicrosoftTeams module exists"
    Write-LogFileMessage "MicrosoftTeams module exists"
  } 
################### END OF COMMANDS ###################