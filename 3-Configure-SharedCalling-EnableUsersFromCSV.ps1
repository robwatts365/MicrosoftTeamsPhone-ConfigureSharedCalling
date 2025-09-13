<# 
Enable Users for Shared Calling in Microsoft Teams using a CSV file
    Version: v1.0
    Date: 10/01/2024
    Author: Rob Watts https://github.com/robwatts365
    Description: This script will users for the Shared Calling feature for Microsoft Teams
#>

# Import Teams Module
Import-Module MicrosoftTeams

# Connects to Microsoft Teams
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams

# Enable File Picker
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# File Picker  (Set File Path - Open File Browser)
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $FilePath = $OpenFileDialog.filename
   
# Store the data from NewUsersFinal.csv in the $ADUsers variable
    Write-Host "Importing CSV..."
    $Users = Import-Csv $FilePath

# Define Teams Shared Calling Routing Policy
    $SharedCallingPolicy = Get-CsTeamsSharedCallingRoutingPolicy | Select-Object Identity, Description, EmergencyLocationId | Out-GridView -OutputMode Single -Title "Please select a Shared Calling Policy"
    Write-Host $SharedCallingPolicy.Identity "is your chosen Shared Calling Policy"
# Define Teams Voice Routing Policy
    $VoiceRoutingPolicy = Get-CsOnlineVoiceRoutingPolicy | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Voice Routing Policy"
    Write-Host $VoiceRoutingPolicy.Identity "is your chosen Voice Routing Policy"
# Define Teams Calling Line Identity Policy
    $CallingLineIdentity = Get-CsCallingLineIdentity | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Calling Line Identity"
    Write-Host $CallingLineIdentity.Identity "is your chosen Calling Line Identity"
# Define Emergency Call Routing Policy
    $EmergencyCallRoutingPolicy = Get-CsTeamsEmergencyCallRoutingPolicy | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select an Emergency Call Routing Policy"
    Write-Host $EmergencyCallRoutingPolicy.Identity "is your chosen Emergency Call Routing Policy"
# Define Teams Dial Plan
    $DialPlan = Get-CsTenantDialPlan | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Dial Plan"
    Write-Host $DialPlan.Identity "is your chosen Dial Plan"

# Loop through each row containing user details in the CSV file
foreach ($User in $Users) {

    # Read user data from each field in each row and assign the data to a variable as below
    $UPN = $User.UPN
        
    Set-CsPhoneNumberAssignment -Identity $UPN -EnterpriseVoiceEnabled $true
    Grant-CsTeamsSharedCallingRoutingPolicy -Identity $UPN -PolicyName $SharedCallingPolicy.Identity
    Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName $VoiceRoutingPolicy.Identity
    Grant-CsCallingLineIdentity -Identity $UPN -PolicyName $CallingLineIdentity.Identity
    Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $UPN -PolicyName $EmergencyCallRoutingPolicy.Identity
    Grant-CsTenantDialPlan -PolicyName $DialPlan.Identity -Identity $UPN
    Write-Host "User" $UPN "has been enabled for Shared Calling."  
    }


