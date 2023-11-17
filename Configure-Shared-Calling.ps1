<# 
Configure Shared Calling
    Version: v1.0
    Date: 13/11/2023
    Author: Rob Watts - Cloud Solution Architect - Microsoft
    Description: This script will configure the Shared Calling feature for Microsoft Teams

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

# Pop out disclaimer
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[Windows.Forms.MessageBox]::Show("
THIS CODE IS SAMPLE CODE. 

THESE SAMPLES ARE PROVIDED 'AS IS' WITHOUT WARRANTY OF ANY KIND.

MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. 

THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU.

IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.", "***DISCLAIMER***", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)

# Enable File Saver for Log File
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# Set Log File Location 
Write-Host "Please select the Log file location."
$LogSaver = New-Object -Typename System.Windows.Forms.SaveFileDialog
$LogSaver.initialDirectory = $initialDirectory
$LogSaver.filter = "All files (*.log)| *.log"
$LogSaver.ShowDialog() | Out-Null
$LogFile = $LogSaver.filename


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

#Shows Script Menu
Show-Menu –Title 'Script Options'
 $selection = Read-Host "Please choose your telephony configuration option" 
 switch ($selection)
 {
     '1' {
         Write-Host "Direct Routing selected" -ForegroundColor Gray -BackgroundColor Black
         Write-LogFileMessage "Direct Routing selected"
         New-SharedCallingDirectRoutingConfig
     } '2' {
        Write-Host "Calling Plans selected" -ForegroundColor Gray -BackgroundColor Black
        Write-LogFileMessage "Calling Plans selected"
        New-SharedCallingCallingPlansConfig
    } '3' {
        Write-Host "Operator Connect selected" -ForegroundColor Gray -BackgroundColor Black
        Write-LogFileMessage "Operator Connect selected"
        New-SharedCallingOperatorConnectConfig
     } 'q' {
         return
     }
 }

################### GLOBAL VARIABLES ###################
$SharedCallingDomain = "None"
$SharedCallingAAName = "None"
$SharedCallingAAUPN = $SharedCallingAAName + "@" + $SharedCallingDomain
$EmergencyPolicySelection = "None"
$EmergencyLocation.LocationID = "None"
################### GLOBAL VARIABLES ###################


################### SHARED CALLING TASK FUNCTIONS ###################

# Connects to relevant PowerShell Modules
function Connect-PowerShell
    {
    Connect-AzureAD
    Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All"
    Connect-MSOLService
    Connect-ExchangeOnline
    Connect-MicrosoftTeams
    }


# Creates Resource Account
function New-SharedCallingResourceAccount 
    {
    # Set Domain Name
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    $SharedCallingDomainTitle = 'Domain Name'
    $SharedCallingDomainMsg   = 'Please provide the Domain Name to use (e.g. contoso.com):'
        
    $SharedCallingDomain = [Microsoft.VisualBasic.Interaction]::InputBox( $SharedCallingDomainTitle, $SharedCallingDomainMsg )

    # Write to Host and Log File
    Write-Host "Domain Name: $SharedCallingDomain"
    Write-LogFileMessage "Domain Name: $SharedCallingDomain"

    # Set Resource Account Name
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    $SharedCallingAANameTitle = 'Auto Attendant Name'
    $SharedCallingAANameMsg   = 'Please provide the Auto Attendant Name (e.g. AA-SharedCalling-UK):'
     
    $SharedCallingAAName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAANameTitle, $SharedCallingAANameMsg)
    
    # Write to Host and Log File 
    Write-Host "Resource Account Name: $SharedCallingAAName"
    Write-LogFileMessage "Resource Account Name: $SharedCallingAAName"

    # Set Resource Account UPN
    $SharedCallingAAUPN = $SharedCallingAAName + "@" + $SharedCallingDomain
    
    # Write to Host and Log File
    Write-Host "Resource Account UPN: $SharedCallingAAUPN"
    Write-LogFileMessage "Resource Account UPN: $SharedCallingAAUPN"


    # Create Resource Account 
    New-CsOnlineApplicationInstance -UserPrincipalName "$SharedCallingAAUPN" -ApplicationId ce933385-9390-45d1-9512-c8d228074e07 -DisplayName "$SharedCallingAAName"
    
    # Assign Licensing Usage Location
    Set-ScriptSleep 180 
    Set-MsolUser -UserPrincipalName "$SharedCallingAAUPN" -UsageLocation GB
    
    # Assign Teams Phone Resource Account License
    Set-ScriptSleep 180 
    $TeamsResourceLicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'PHONESYSTEM_VIRTUALUSER'
    Set-MgUserLicense -UserId "$SharedCallingAAUPN" -addLicenses @{SkuId = $TeamsResourceLicenseSku.SkuId} -RemoveLicenses @()
    }

# Function to Configure Shared Calling Emergency Address
function Set-SharedCallingEmergencyAddress
    {
        $EmergencyPolicySelection = [System.Windows.Forms.MessageBox]::Show("Use an existing Emergency Location?" , "Status" , 3, 32)

        if ($EmergencyPolicySelection -eq "Yes") {
            <# Action to perform if the answer is Yes #>
            $EmergencyLocation = Get-CsOnlineLisLocation | Select-Object CompanyName,Description, HouseNumber, StreetName, City, Postcode,CountryOrRegion, LocationID | Out-GridView -OutputMode Single -Title "Please select an Emergency Location"
            Write-Host "Existing Emergency Location selected"
            Write-LogFileMessage "Existing Emergency Location selected"
            Write-Host "Emergency Location: $EmergencyLocation"
            Write-LogFileMessage "Emergency Location: $EmergencyLocation"
        }elseif ($EmergencyPolicySelection -eq "No") {
            Write-Host "New Emergency Location selected"
            Write-LogFileMessage "New Emergency Location selected"
            Write-Host "Emergency Location: $EmergencyLocation"
            Write-LogFileMessage "Emergency Location: $EmergencyLocation"
            
            #Write-Host "No"
        }
        else {
            <# Action when all if and elseif conditions are false #>
            Write-Host "Cancelling Command..."
            Write-LogFileMessage "Cancelling Command..."
            return
        }
    }
    

# Function to Configure Shared Calling for Direct Routing numbers
function New-SharedCallingDirectRoutingConfig
    {
        New-SharedCallingResourceAccount 
        Set-SharedCallingEmergencyAddress
    }

# Function to Configure Shared Calling for Calling Plans numbers
function New-SharedCallingCallingPlansConfig
    {
        New-SharedCallingResourceAccount 
        Set-SharedCallingEmergencyAddress
    }
# Function to Configure Shared Calling for Operator Connect numbers
function New-SharedCallingOperatorConnectConfig
    {
        New-SharedCallingResourceAccount 
        Set-SharedCallingEmergencyAddress
    }

################### SHARED CALLING TASK FUNCTIONS ###################

################### INSTALL MODULE FUNCTIONS ###################
Function Install-AzureADPreview
    {
    Install-Module AzureADPreview
    Set-ExecutionPolicy -ExecutionPolicy ReoteSigned
    Write-LogFileMessage "AzureADPreview Module Installed"
    }
Function Install-MicrosoftGraph 
    {
    Install-Module Microsoft.Graph
    Set-ExecutionPolicy -ExecutionPolicy ReoteSigned
    Write-LogFileMessage "MicrosoftGraph Module Installed"
    }
Function Install-MSOnline
    {
    Install-Module MSOnline
    Set-ExecutionPolicy -ExecutionPolicy ReoteSigned
    Write-LogFileMessage "MSOnline Module Installed"
    }
Function Install-ExchangeOnlineManagement
    {
    Install-Module ExchangeOnlineManagement
    Set-ExecutionPolicy -ExecutionPolicy ReoteSigned
    Write-LogFileMessage "ExchangeOnlineManagement Module Installed"
    }
Function Install-MicrosoftTeams
    {
    Install-Module MicrosoftTeams
    Set-ExecutionPolicy -ExecutionPolicy ReoteSigned
    Write-LogFileMessage "MicrosoftTeams Module Installed"
    }

################### INSTALL MODULE FUNCTIONS ###################

################### GENERIC SCRIPT FUNCTIONS ###################

# Write To Log Function
function Write-LogFileMessage($message) 
    {
        $ActionTime = $(Get-Date).ToString("dd/MM/yyyy HH:mm:ss")
        Write-Debug "$ActionTime :: $message"
    "$ActionTime :: $message" | Out-File -FilePath $LogFile -Append
    }

# Progress Bar Function
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

#  Menu Options Function
function Show-Menu
    {
        param (
            [string]$Title = 'Script Options'
        )
        Write-Host "================ $Title ================" -ForegroundColor Yellow -BackgroundColor Black 
        
        Write-Host "1: Press '1' for Direct Routing." -ForegroundColor White -BackgroundColor Black 
        Write-Host "2: Press '2' for Calling Plans." -ForegroundColor White -BackgroundColor Black 
        Write-Host "3: Press '3' for Operator Connect." -ForegroundColor White -BackgroundColor Black  
        Write-Host "Q: Press 'Q' to quit." -ForegroundColor White -BackgroundColor Black 
    }
################### GENERIC SCRIPT FUNCTIONS ###################