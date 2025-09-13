﻿<# 
Configure Shared Calling
    Version: v1.2
    Date: 23/01/2024
    Author: Rob Watts https://github.com/robwatts365
    Description: This script will configure the Shared Calling feature for Microsoft Teams
#>

################### START OF GLOBAL VARIABLES ###################
$global:SharedCallingDomain = "None"
$global:SharedCallingAAName = "None"
$global:SharedCallingAAName = "None"
$global:SharedCallingDomain = "None"
$global:SharedCallingAAUPN = $global:SharedCallingDomainSharedCallingAAName + "@" + $global:SharedCallingDomain
$global:SharedCallingAANumber = "None"
$global:EmergencyLocation = "None"
################### END OF GLOBAL VARIABLES ###################

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

### FUNCTION - Menu Options Function ###
function ShowMenu
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

################### END OF GENERIC SCRIPT FUNCTIONS ###################

################### START OF SHARED CALLING TASK FUNCTIONS ###################

### FUNCTION - Connects to relevant PowerShell Modules ###
function Connect-PowerShell
    {
    Connect-AzureAD | Out-File -FilePath $LogFile -Append
    Write-Host "AzureAD Module Connected." -ForegroundColor Green
    Write-LogFileMessage "AzureAD Module Connected."

    Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All" | Out-File -FilePath $LogFile -Append
    Write-Host "Microsoft Graph Module Connected." -ForegroundColor Green
    Write-LogFileMessage "Microsoft Graph Module Connected."

    Connect-MSOLService | Out-File -FilePath $LogFile -Append
    Write-Host "MSOnline Module Connected." -ForegroundColor Green
    Write-LogFileMessage "MSOnline Module Connected."

    Connect-ExchangeOnline | Out-File -FilePath $LogFile -Append
    Write-Host "Exchange Online Module Connected" -ForegroundColor Green
    Write-LogFileMessage "Exchange Online Module Connected"

    Connect-MicrosoftTeams | Out-File -FilePath $LogFile -Append
    Write-Host "Microsoft Teams Module Connected" -ForegroundColor Green
    Write-LogFileMessage "Microsoft Teams Module Connected"
    }

### FUNCTION - Creates Resource Account ###
function New-SharedCallingResourceAccount 
    {
    # Set Domain Name
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $SharedCallingDomainMsg = 'Domain Name'
    $SharedCallingDomainTitle   = 'Please provide the Domain Name to use (e.g. contoso.com):'
    $global:SharedCallingDomain = [Microsoft.VisualBasic.Interaction]::InputBox( $SharedCallingDomainTitle, $SharedCallingDomainMsg)

    # Write to Host and Log File
    Write-Host "Domain Name: $global:SharedCallingDomain"
    Write-LogFileMessage "Domain Name: $global:SharedCallingDomain"

    # Set Resource Account Name
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $SharedCallingAANameMsg = 'Auto Attendant Name'
    $SharedCallingAANameTitle   = 'Please provide the Auto Attendant Name (e.g. AA-SharedCalling-UK):'
    $global:SharedCallingAAName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAANameTitle, $SharedCallingAANameMsg)
    
    # Write to Host and Log File 
    Write-Host "Resource Account Name: $global:SharedCallingAAName"
    Write-LogFileMessage "Resource Account Name: $global:SharedCallingAAName"

    # Set Phone Number
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $SharedCallingAANumberMsg = 'Auto Attendant Number'
    $SharedCallingAANumberTitle   = 'Please provide the Auto Attendant Phone Number in E.164 format (e.g. +441632969000):'
    $global:SharedCallingAANumber = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAANumberTitle, $SharedCallingAANumberMsg)
    
    # Write to Host and Log File 
    Write-Host "Resource Account Number: $global:SharedCallingAANumber"
    Write-LogFileMessage "Resource Account Number: $global:SharedCallingAANumber"

    # Set Resource Account UPN
    $global:SharedCallingAAUPN = $global:SharedCallingAAName + "@" + $global:SharedCallingDomain
    # Write to Host and Log File
    Write-Host "Resource Account UPN: $global:SharedCallingAAUPN"
    Write-LogFileMessage "Resource Account UPN: $global:SharedCallingAAUPN"


    # Create Resource Account 
    New-CsOnlineApplicationInstance -UserPrincipalName "$global:SharedCallingAAUPN" -ApplicationId ce933385-9390-45d1-9512-c8d228074e07 -DisplayName "$global:SharedCallingAAName" | Out-File -FilePath $LogFile -Append
    
    # Assign Licensing Usage Location
    Set-ScriptSleep 120 
    Set-MsolUser -UserPrincipalName "$global:SharedCallingAAUPN" -UsageLocation GB
    
    # Assign Teams Phone Resource Account License
    Set-ScriptSleep 120 
    $TeamsResourceLicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'PHONESYSTEM_VIRTUALUSER'
    Set-MgUserLicense -UserId "$global:SharedCallingAAUPN" -addLicenses @{SkuId = $TeamsResourceLicenseSku.SkuId} -RemoveLicenses @() | Out-File -FilePath $LogFile -Append
    
    return $global:SharedCallingDomain,$global:SharedCallingAAName,$global:SharedCallingAAUPN,$global:SharedCallingAANumber
    }

### FUNCTION - Configure Shared Calling Emergency Address ###
function Set-SharedCallingEmergencyAddress
    {
        $EmergencyLocationSelection = [System.Windows.Forms.MessageBox]::Show("Use an existing Emergency Location?" , "Input Required" , 4, 64)

        if ($EmergencyLocationSelection -eq "Yes") {
            <# Action to perform if the answer is Yes #>
            $global:EmergencyLocation = Get-CsOnlineLisCivicAddress | Select-Object CompanyName,Description, HouseNumber, StreetName, City, Postcode,CountryOrRegion, DefaultLocationID | Out-GridView -OutputMode Single -Title "Please select an Emergency Location" 
            Write-Host "Existing Emergency Location selected."
            Write-LogFileMessage "Existing Emergency Location selected."
            Write-Host "Emergency Location:" $global:EmergencyLocation.DefaultLocationID
            Write-LogFileMessage "Emergency Location:" $global:EmergencyLocation.DefaultLocationID
            
        }elseif ($EmergencyLocationSelection -eq "No") {
            Write-Host "New Emergency Location selected."
            Write-LogFileMessage "New Emergency Location selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationDescMsg = 'Emergency Location Description'
            $EmergencyLocationDescTitle   = 'Please provide the Emergency Location Description (e.g. Microsoft-Campus-B2-TVP):'
            $EmergencyLocationDesc = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationDescTitle, $EmergencyLocationDescMsg)
    
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationCompanyMsg = 'Emergency Location Company'
            $EmergencyLocationCompanyTitle   = 'Please provide the Emergency Location Company (e.g. Microsoft):'
            $EmergencyLocationCompany = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationCompanyTitle, $EmergencyLocationCompanyMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationHouseMsg = 'Emergency Location House/Building'
            $EmergencyLocationHouseTitle   = 'Please provide the Emergency Location House/Building Name (e.g. Microsoft Campus Building 2):'
            $EmergencyLocationHouse = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationHouseTitle, $EmergencyLocationHouseMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationStreetMsg = 'Emergency Location Street'
            $EmergencyLocationStreetTitle   = 'Please provide the Emergency Location Street (e.g. Thames Valley Park):'
            $EmergencyLocationStreet = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationStreetTitle, $EmergencyLocationStreetMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationCityMsg = 'Emergency Location City'
            $EmergencyLocationCityTitle   = 'Please provide the Emergency Location City (e.g. Reading):'
            $EmergencyLocationCity = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationCityTitle, $EmergencyLocationCityMsg)
        
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationStateMsg = 'Emergency Location State/Province'
            $EmergencyLocationStateTitle   = 'Please provide the Emergency Location State/Province (e.g. Berkshire):'
            $EmergencyLocationState = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationStateTitle, $EmergencyLocationStateMsg)
            
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationCountryMsg = 'Emergency Location Country/Region'
            $EmergencyLocationCountryTitle   = 'Please provide the Emergency Location Country/Region (e.g. GB):'
            $EmergencyLocationCountry = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationCountryTitle, $EmergencyLocationCountryMsg)
    
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationPostCodeMsg = 'Emergency Location Postal Code'
            $EmergencyLocationPostCodeTitle   = 'Please provide the Emergency Location PostCode (e.g. RG6 1WG):'
            $EmergencyLocationPostCode = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationPostCodeTitle, $EmergencyLocationPostCodeMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationLatitudeMsg = 'Emergency Location Latitude'
            $EmergencyLocationLatitudeTitle   = 'Please provide the Emergency Location Latitude (e.g. 51.060692):'
            $EmergencyLocationLatitude = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationLatitudeTitle, $EmergencyLocationLatitudeMsg)
    
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationLongitudeMsg = 'Emergency Location Longitude'
            $EmergencyLocationLongitudeTitle   = 'Please provide the Emergency Location Longitude (e.g. -1.313154):'
            $EmergencyLocationLongitude = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationLongitudeTitle, $EmergencyLocationLongitudeMsg)


            New-CsOnlineLisCivicAddress -HouseNumber $EmergencyLocationHouse -StreetName $EmergencyLocationStreet -City $EmergencyLocationCity -StateorProvince $EmergencyLocationState -CountryOrRegion $EmergencyLocationCountry -PostalCode $EmergencyLocationPostCode -Description $EmergencyLocationDesc -CompanyName $EmergencyLocationCompany -Latitude $EmergencyLocationLatitude -Longitude $EmergencyLocationLongitude | Out-File -FilePath $LogFile -Append
            $global:EmergencyLocation = Get-CsOnlineLisLocation | Select-Object CompanyName,Description, HouseNumber, StreetName, City, Postcode,CountryOrRegion, DefaultLocationID | Out-GridView -OutputMode Single -Title "Please select the newly created Emergency Location"
            
            Write-Host "Emergency Location:" $global:EmergencyLocation.DefaultLocationID
            Write-LogFileMessage "Emergency Location:" $global:EmergencyLocation.DefaultLocationID
            
        }
        return $global:EmergencyLocation
    }
### FUNCTION - Create Auto Attendant ###
function New-SharedCallingAutoAttendant
    {
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $SharedCallingAAGreetingMsg = 'Auto Attendant Greeting'
        $SharedCallingAAGreetingTitle   = 'Please provide the Auto Attendant Greeting (e.g. Welcome to Contoso!):'
        $SharedCallingAAGreeting = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAAGreetingTitle, $SharedCallingAAGreetingMsg)

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $SharedCallingAAPromptMsg = 'Auto Attendant Prompt'
        $SharedCallingAAPromptTitle   = 'Please provide the Auto Attendant Prompt (e.g. If you know the extension you require, dial it now):'
        $SharedCallingAAPrompt = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAAPromptTitle, $SharedCallingAAPromptMsg)

        $greetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $SharedCallingAAGreeting
        $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $SharedCallingAAPrompt
        $defaultMenu = New-CsAutoAttendantMenu -Name "Default menu" -Prompts @($menuPrompt) -EnableDialByName -DirectorySearchMethod ByExtension
        $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default call flow" -Menu $defaultMenu -Greetings @($greetingPrompt)
        New-CsAutoAttendant -Name $global:SharedCallingAAName -DefaultCallFlow $defaultCallFlow -EnableVoiceResponse -LanguageId "en-GB" -TimeZoneId "GMT Standard Time" | Out-File -FilePath $LogFile -Append

        ### Assign Resource Account
        $applicationInstanceId = (Get-CsOnlineUser $global:SharedCallingAAUPN).Identity
        $autoAttendantId = (Get-CsAutoAttendant -NameFilter $global:SharedCallingAAName).Id
        New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceId) -ConfigurationId $autoAttendantId -ConfigurationType AutoAttendant | Out-File -FilePath $LogFile -Append

        Write-Host "Auto Attendant: $global:SharedCallingAAName created"
        Write-LogFileMessage "Auto Attendant: $global:SharedCallingAAName created"
    }

### FUNCTION - Shared CallingEmergency Call Routing Policy ###
function Set-SharedCallingEmergencyCallRoutingPolicy
    {
      $EmergencyCallRoutingPolicySelection = [System.Windows.Forms.MessageBox]::Show("Create a new Emergency Call Routing Policy?" , "Input Required" , 4, 64)

        if ($EmergencyCallRoutingPolicySelection -eq "No") {
          Write-Host "Existing Emergency Call Routing Policy selected."
          Write-LogFileMessage "Existing Emergency Call Routing Policy selected."

        }elseif ($EmergencyCallRoutingPolicySelection -eq "Yes") {
            
            Write-Host "New Emergency Call Routing Policy selected."
            Write-LogFileMessage "New Emergency Call Routing Policy selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingEmergencyRoutingPolicyNameMsg = 'Emergency Dial Routing Policy Name'
            $SharedCallingEmergencyRoutingPolicyNameTitle   = 'Please provide the Emergency Call Routing Policy Name (e.g. UK-ECRP):'
            $SharedCallingEmergencyRoutingPolicyName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingEmergencyRoutingPolicyNameTitle, $SharedCallingEmergencyRoutingPolicyNameMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingEmergencyDialStringMsg = 'Emergency Dial String'
            $SharedCallingEmergencyDialStringTitle   = 'Please provide the Emergency Dial String (e.g. 999):'
            $SharedCallingEmergencyDialString = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingEmergencyDialStringTitle, $SharedCallingEmergencyDialStringMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingEmergencyDialMaskMsg = 'Emergency Dial Mask'
            $SharedCallingEmergencyDialMaskTitle   = 'Please provide the Emergency Dial Mask (e.g. 999):'
            $SharedCallingEmergencyDialMask = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingEmergencyDialMaskTitle, $SharedCallingEmergencyDialMaskMsg)
            
            $en1 = New-CsTeamsEmergencyNumber -EmergencyDialString $SharedCallingEmergencyDialString -EmergencyDialMask $SharedCallingEmergencyDialMask
            New-CsTeamsEmergencyCallRoutingPolicy -Identity $SharedCallingEmergencyRoutingPolicyName -EmergencyNumbers @{add=$en1} -AllowEnhancedEmergencyServices:$true | Out-File -FilePath $LogFile -Append

            Write-Host "New emergency Call Routing Policy: $EmergencyCallRoutingPolicy"
            Write-LogFileMessage "New emergency Call Routing Policy: $EmergencyCallRoutingPolicy"
        }
      }
### FUNCTION - Shared Calling Voice Routing Policy ###
function Set-SharedCallingVoiceRoutingPolicy
        {
        $VoiceRoutingPolicySharedCallingSelection = [System.Windows.Forms.MessageBox]::Show("Create a new Shared Calling Voice Routing Policy?" , "Shared Calling Voice Routing Policy creation" , 4, 64)

        if ($VoiceRoutingPolicySharedCallingSelection -eq "No") {
          Write-Host "Existing Shared Calling Voice Routing Policy selected."
          Write-LogFileMessage "Existing Shared Calling Voice Routing Policy selected."

        }elseif ($VoiceRoutingPolicySharedCallingSelection -eq "Yes") {
            
            Write-Host "New Shared Calling Voice Routing Policy selected."
            Write-LogFileMessage "Shared Calling Voice Routing Policy selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingVRPNameMsg = 'Emergency Dial Routing Policy Name'
            $SharedCallingVRPNameTitle   = 'Please provide the Shared Calling Voice Routing Policy Name (e.g. UK-SharedCallingPolicy):'
            $SharedCallingVRPName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingVRPNameTitle, $SharedCallingVRPNameMsg)
                       
            New-CSOnlineVoiceRoutingPolicy -Identity $SharedCallingVRPName | Out-File -FilePath $LogFile -Append

            Write-Host "New Shared Calling Voice Routing Policy: $SharedCallingVRPName"
            Write-LogFileMessage "New Shared Calling Voice Routing Policy: $SharedCallingVRPName"
        }
      }

### FUNCTION - Shared Calling Caller ID Policy ###
function Set-SharedCallingCallerIDPolicy
        {
        $CallerIDSelection = [System.Windows.Forms.MessageBox]::Show("Create a new Caller ID Policy?" , "Caller ID creation" , 4, 64)

        if ($CallerIDSelection -eq "No") {
          Write-Host "Existing Caller ID Policy selected."
          Write-LogFileMessage "Existing Caller ID Policy selected."

        }elseif ($CallerIDSelection -eq "Yes") {
            
            Write-Host "New Caller ID Policy selected."
            Write-LogFileMessage "Shared Caller ID Policy selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallerIDNameMsg = 'Caller ID Policy Name'
            $SharedCallerIDNameTitle   = 'Please provide the Caller ID Policy Name (e.g. Shared Calling):'
            $SharedCallerIDName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallerIDNameTitle, $SharedCallerIDNameMsg)
                       
            $ObjId = (Get-CsOnlineApplicationInstance -Identity $global:SharedCallingAAUPN).ObjectId
            New-CsCallingLineIdentity -Identity $SharedCallerIDName -CallingIDSubstitute Resource -EnableUserOverride $false -ResourceAccount $ObjId | Out-File -FilePath $LogFile -Append

            Write-Host "New Shared Calling Voice Routing Policy: $SharedCallerIDName"
            Write-LogFileMessage "New Shared Calling Voice Routing Policy: $SharedCallerIDName"
        }
      }

### FUNCTION - Shared Calling Policy ###
function Set-SharedCallingPolicy
        {
        $SharedCallingPolicySelection = [System.Windows.Forms.MessageBox]::Show("Do you require emergency callback numbers?" , "Shared Calling Policy Creation" , 4, 64)

        if ($SharedCallingPolicySelection -eq "Yes") {
          Write-Host "Emergency callback numbers required."
          Write-LogFileMessage "Emergency callback numbers required."

          $SharedCallingRA = Get-CsOnlineUser -Identity $global:SharedCallingAAUPN

          [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
          $EmergencyNumber1Msg = 'Emergency callback number 1'
          $EmergencyNumber1Title   = 'Please provide the first Emergency callback number in E.164 format (e.g. +441632960999):'
          $EmergencyNumber1 = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyNumber1Title, $EmergencyNumber1Msg)

          [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
          $EmergencyNumber2Msg = 'Emergency callback number 2'
          $EmergencyNumber2Title   = 'Please provide the second Emergency callback number in E.164 format (e.g. +441632960999):'
          $EmergencyNumber2 = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyNumber2Title, $EmergencyNumber2Msg)

          New-CsTeamsSharedCallingRoutingPolicy -Identity $global:SharedCallingAAName -ResourceAccount $SharedCallingRA.Identity -EmergencyNumbers @{add=$EmergencyNumber1,$EmergencyNumber2} | Out-File -FilePath $LogFile -Append

          Write-Host "New Shared Calling Voice Routing Policy created: $global:SharedCallingAAName"
          Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $global:SharedCallingAAName"

        }elseif ($SharedCallingPolicySelection -eq "No") {
            
            Write-Host "Emergency callback numbers not required."
            Write-LogFileMessage "Emergency callback numbers not required."

            $SharedCallingRA = Get-CsOnlineUser -Identity $global:SharedCallingAAUPN
            New-CsTeamsSharedCallingRoutingPolicy -Identity $global:SharedCallingAAName -ResourceAccount $SharedCallingRA.Identity | Out-File -FilePath $LogFile -Append
            
            Write-Host "New Shared Calling Voice Routing Policy created: $global:SharedCallingAAName"
            Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $global:SharedCallingAAName"
            
        }

    }

### FUNCTION - Configure Shared Calling for Direct Routing numbers ###
function New-SharedCallingDirectRoutingConfig
    {
        Connect-PowerShell  

        New-SharedCallingResourceAccount
        
        Write-Host "Shared Calling Resource Accounts tasks completed." -ForegroundColor Green
        Write-Host "Shared Calling Resource account created: $global:SharedCallingAAName"
        Write-LogFileMessage "Shared Calling Resource Accounts tasks completed."
        Write-LogFileMessage "Shared Calling Resource account created: $global:SharedCallingAAName"

        Set-SharedCallingEmergencyAddress

        Write-Host "Shared Calling Emergency Address tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Emergency Address tasks completed."

        $RAVoiceRoutingPolicy = Get-CsOnlineVoiceRoutingPolicy | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Voice Routing Policy"
       
        Write-Host $RAVoiceRoutingPolicy.Identity "is your chosen Voice Routing Policy"
        Write-LogFileMessage $RAVoiceRoutingPolicy.Identity "is your chosen Voice Routing Policy"

        Grant-CsOnlineVoiceRoutingPolicy -PolicyName $RAVoiceRoutingPolicy.Identity -Identity $global:SharedCallingAAUPN | Out-File -FilePath $LogFile -Append

        Write-Host "Shared Calling Voice Routing Policy tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Routing Policy tasks completed."

        New-SharedCallingAutoAttendant

        Write-Host "Shared Calling Auto Attendant tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Auto Attendant tasks completed."
        
        Set-CsPhoneNumberAssignment -Identity $global:SharedCallingAAUPN -LocationID $global:EmergencyLocation.DefaultLocationID -PhoneNumber $global:SharedCallingAANumber -PhoneNumberType DirectRouting | Out-File -FilePath $LogFile -Append
        
        Write-Host "Shared Resource Account Phone Number ($global:SharedCallingAANumber) assigned to $global:SharedCallingAAUPN." -ForegroundColor Green
        Write-LogFileMessage "Shared Resource Account Phone Number ($global:SharedCallingAANumber) assigned to $global:SharedCallingAAUPN."

        Set-SharedCallingEmergencyCallRoutingPolicy
        Write-Host "Shared Emergency Call Routing Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Emergency Call Routing Policy Configuration tasks completed."
        
        Set-SharedCallingVoiceRoutingPolicy
        Write-Host "Shared Calling Voice Routing Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Routing Policy Configuration tasks completed."
        
        Set-SharedCallingCallerIDPolicy
        Write-Host "Shared Calling Caller ID Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Caller ID Policy Configuration tasks completed."
        
        Set-SharedCallingPolicy
        Write-Host "Shared Calling Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Policy Configuration tasks completed."
    }

### FUNCTION - Configure Shared Calling for Calling Plans numbers ###
function New-SharedCallingCallingPlansConfig
    {
        Connect-PowerShell  

        New-SharedCallingResourceAccount
        
        Write-Host "Shared Calling Resource Accounts tasks completed." -ForegroundColor Green
        Write-Host "Shared Calling Resource account created: $global:SharedCallingAAName"
        Write-LogFileMessage "Shared Calling Resource Accounts tasks completed."
        Write-LogFileMessage "Shared Calling Resource account created: $global:SharedCallingAAName"

        Set-SharedCallingEmergencyAddress

        Write-Host "Shared Calling Emergency Address tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Emergency Address tasks completed."

      ### Assign Pay-as-you go calling plan (Zone-1 countries i.e. UK)
        $PAYGCallingZone1LicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'Microsoft_Teams_Calling_Plan_pay_as_you_go_(country_zone_1)'
        Set-MgUserLicense -UserId $global:SharedCallingAAUPN -addLicenses @{SkuId = $PAYGCallingZone1LicenseSku.SkuId} -RemoveLicenses @() | Out-File -FilePath $LogFile -Append
        
      ### PAUSE - wait a few minutes for the above cmdlet configuration to be completed
        Set-ScriptSleep 120
      
      ### Assign communication credits       
        $CommunicationCreditsLicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'MCOPSTNC'
        Set-MgUserLicense -UserId $global:SharedCallingAAUPN -addLicenses @{SkuId = $CommunicationCreditsLicenseSku.SkuId} -RemoveLicenses @() | Out-File -FilePath $LogFile -Append

        Write-Host "Calling Plans licensing tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Calling Plans licensing tasks completed."

        New-SharedCallingAutoAttendant

        Write-Host "Shared Calling Auto Attendant tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Auto Attendant tasks completed."

        Set-CsPhoneNumberAssignment -Identity $global:SharedCallingAAUPN -LocationID $global:EmergencyLocation.DefaultLocationID -PhoneNumber $global:SharedCallingAANumber -PhoneNumberType CallingPlan | Out-File -FilePath $LogFile -Append
        Write-Host "Shared Resource Account Phone Number ($global:SharedCallingAANumber) assigned to $global:SharedCallingAAUPN." -ForegroundColor Green
        Write-LogFileMessage "Shared Resource Account Phone Number ($global:SharedCallingAANumber) assigned to $global:SharedCallingAAUPN."

        Set-SharedCallingEmergencyCallRoutingPolicy
        Write-Host "Shared Emergency Call Routing Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Emergency Call Routing Policy Configuration tasks completed."
        
        Set-SharedCallingVoiceRoutingPolicy
        Write-Host "Shared Calling Voice Routing Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Routing Policy Configuration tasks completed."
        
        Set-SharedCallingCallerIDPolicy
        Write-Host "Shared Calling Caller ID Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Caller ID Policy Configuration tasks completed."
        
        Set-SharedCallingPolicy
        Write-Host "Shared Calling Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Policy Configuration tasks completed."
    }
### FUNCTION - Configure Shared Calling for Operator Connect numbers ###
function New-SharedCallingOperatorConnectConfig
    {
        Connect-PowerShell
        
        New-SharedCallingResourceAccount
        
        Write-Host "Shared Calling Resource Accounts tasks completed." -ForegroundColor Green
        Write-Host "Shared Calling Resource account created: $global:SharedCallingAAName"
        Write-LogFileMessage "Shared Calling Resource Accounts tasks completed."
        Write-LogFileMessage "Shared Calling Resource account created: $global:SharedCallingAAName"

        Set-SharedCallingEmergencyAddress

        Write-Host "Shared Calling Emergency Address tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Emergency Address tasks completed."

        New-SharedCallingAutoAttendant

        Write-Host "Shared Calling Auto Attendant tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Auto Attendant tasks completed."

        Set-CsPhoneNumberAssignment -Identity $global:SharedCallingAAUPN -LocationID $global:EmergencyLocation.DefaultLocationID -PhoneNumber $global:SharedCallingAANumber -PhoneNumberType OperatorConnect | Out-File -FilePath $LogFile -Append
        
        Write-Host "Shared Resource Account Phone Number ($global:SharedCallingAANumber) assigned to $global:SharedCallingAAUPN." -ForegroundColor Green
        Write-LogFileMessage "Shared Resource Account Phone Number ($global:SharedCallingAANumber) assigned to $global:SharedCallingAAUPN."

        Set-SharedCallingEmergencyCallRoutingPolicy
        Write-Host "Shared Emergency Call Routing Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Emergency Call Routing Policy Configuration tasks completed."
        
        Set-SharedCallingVoiceRoutingPolicy
        Write-Host "Shared Calling Voice Routing Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Routing Policy Configuration tasks completed."
        
        Set-SharedCallingCallerIDPolicy
        Write-Host "Shared Calling Caller ID Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Caller ID Policy Configuration tasks completed."
        
        Set-SharedCallingPolicy
        Write-Host "Shared Calling Policy Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Policy Configuration tasks completed."

    }

################### END OF SHARED CALLING TASK FUNCTIONS ###################

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


# Installs Powershell Modules, if not already installed
Write-Host "Checking for PowerShell Modules" -ForegroundColor Gray -BackgroundColor Black
Write-LogFileMessage "Checking for PowerShell Modules"
#AzureADPreview
If (-not(Get-InstalledModule AzureADPreview -ErrorAction silentlycontinue)) {
    Write-Host "AzureADPreview module does not exist. Please run pre-requisite script." -ForegroundColor DarkRed
    Write-LogFileMessage "AzureADPreview module does not exist. Please run pre-requisite script."
    Install-AzureADPreview
  }
  Else {
    Write-Host "AzureADPreview module exists." -ForegroundColor Green
    Write-LogFileMessage "AzureADPreview module exists."
  }
  
#Microsoft.Graph
  If (-not(Get-InstalledModule Microsoft.Graph -ErrorAction silentlycontinue)) {
    Write-Host "Microsoft.Graph module does not exist. Please run pre-requisite script." -ForegroundColor DarkRed
    Write-LogFileMessage "Microsoft.Graph module does not exist. Please run pre-requisite script."
    Install-MicrosoftGraph
  }
  Else {
    Write-Host "Microsoft.Graph module exists." -ForegroundColor Green
    Write-LogFileMessage "Microsoft.Graph module exists."
  }
  
#MSOnline
  If (-not(Get-InstalledModule MSOnline -ErrorAction silentlycontinue)) {
    Write-Host "MSOnline module does not exist. Please run pre-requisite script." -ForegroundColor DarkRed
    Write-LogFileMessage "MSOnline module does not exist. Please run pre-requisite script."
    Install-MSOnline
  }
  Else {
    Write-Host "MSOnline module exists." -ForegroundColor Green
    Write-LogFileMessage "MSOnline module exists."
  }
  
#ExchangeOnlineManagement
  If (-not(Get-InstalledModule ExchangeOnlineManagement -ErrorAction silentlycontinue)) {
    Write-Host "ExchangeOnlineManagement module does not exist. Please run pre-requisite script." -ForegroundColor DarkRed
    Write-LogFileMessage "ExchangeOnlineManagement module does not exist. Please run pre-requisite script."
    Install-ExchangeOnlineManagement
  }
  Else {
    Write-Host "ExchangeOnlineManagement module exists." -ForegroundColor Green
    Write-LogFileMessage "ExchangeOnlineManagement module exists."
  }

#MicrosoftTeams
  If (-not(Get-InstalledModule MicrosoftTeams -ErrorAction silentlycontinue)) {
    Write-Host "MicrosoftTeams module does not exist. Please run pre-requisite script." -ForegroundColor DarkRed
    Write-LogFileMessage "MicrosoftTeams module does not exist. Please run pre-requisite script."
    Install-MicrosoftTeams
  }
  Else {
    Write-Host "MicrosoftTeams module exists." -ForegroundColor Green
    Write-LogFileMessage "MicrosoftTeams module exists."
  } 

#Shows Script Menu
ShowMenu –Title 'Script Options'
 $selection = Read-Host "Please choose your telephony configuration option." 
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
 Write-Host "Shared Calling configuration completed. Please remember to enable your users." -ForegroundColor Green
 Write-LogFileMessage "Shared Calling configuration completed. Please remember to enable your users."
################### END OF COMMANDS ###################

