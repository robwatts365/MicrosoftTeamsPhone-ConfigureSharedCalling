<# 
Configure Shared Calling
    Version: v1.0
    Date: 11/01/2024
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

################### START OF GLOBAL VARIABLES ###################
$SharedCallingDomain = "None"
$SharedCallingAAName = "None"
$SharedCallingAAUPN = $SharedCallingAAName + "@" + $SharedCallingDomain
$SharedCallingAANumber = "None"
$EmergencyLocation = "None"
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
### FUNCTION - Show Disclaimer ###
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

################### START OF SHARED CALLING TASK FUNCTIONS ###################

### FUNCTION - Connects to relevant PowerShell Modules ###
function Connect-PowerShell
    {
    Connect-AzureAD
    Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All"
    Connect-MSOLService
    Connect-ExchangeOnline
    Connect-MicrosoftTeams
    }

### FUNCTION - Creates Resource Account ###
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

    # Set Phone Number
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $SharedCallingAANumberTitle = 'Auto Attendant Number'
    $SharedCallingAANumberMsg   = 'Please provide the Auto Attendant Phone Number in E.164 format (e.g. +441632969000):'
    $SharedCallingAANumber = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAANumberTitle, $SharedCallingAANumberMsg)
    
    # Write to Host and Log File 
    Write-Host "Resource Account Number: $SharedCallingAANumber"
    Write-LogFileMessage "Resource Account Number: $SharedCallingAANumber"

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

### FUNCTION - Configure Shared Calling Emergency Address ###
function Set-SharedCallingEmergencyAddress
    {
        $EmergencyLocationSelection = [System.Windows.Forms.MessageBox]::Show("Use an existing Emergency Location?" , "Input Required" , 4, 64)

        if ($EmergencyLocationSelection -eq "Yes") {
            <# Action to perform if the answer is Yes #>
            $EmergencyLocation = Get-CsOnlineLisLocation | Select-Object CompanyName,Description, HouseNumber, StreetName, City, Postcode,CountryOrRegion, LocationID | Out-GridView -OutputMode Single -Title "Please select an Emergency Location"
            Write-Host "Existing Emergency Location selected"
            Write-LogFileMessage "Existing Emergency Location selected"
            Write-Host "Emergency Location: $EmergencyLocation.LocationID"
            Write-LogFileMessage "Emergency Location: $EmergencyLocation.LocationID"
        }elseif ($EmergencyLocationSelection -eq "No") {
            Write-Host "New Emergency Location selected"
            Write-LogFileMessage "New Emergency Location selected"

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationDescTitle = 'Emergency Location Description'
            $EmergencyLocationDescMsg   = 'Please provide the Emergency Location Description (e.g. Microsoft-Campus-B2-TVP):'
            $EmergencyLocationDesc = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationDescTitle, $EmergencyLocationDescMsg)
    
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationCompanyTitle = 'Emergency Location Company'
            $EmergencyLocationCompanyMsg   = 'Please provide the Emergency Location Company (e.g. Microsoft):'
            $EmergencyLocationCompany = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationCompanyTitle, $EmergencyLocationCompanyMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationHouseTitle = 'Emergency Location House/Building'
            $EmergencyLocationHouseMsg   = 'Please provide the Emergency Location House/Building Name (e.g. Microsoft Campus Building 2):'
            $EmergencyLocationHouse = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationHouseTitle, $EmergencyLocationHouseMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationStreetTitle = 'Emergency Location Street'
            $EmergencyLocationStreetMsg   = 'Please provide the Emergency Location Street (e.g. Thames Valley Park):'
            $EmergencyLocationStreet = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationStreetTitle, $EmergencyLocationStreetMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationCityTitle = 'Emergency Location City'
            $EmergencyLocationCityMsg   = 'Please provide the Emergency Location City (e.g. Reading):'
            $EmergencyLocationCity = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationCityTitle, $EmergencyLocationCityMsg)
        
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationStateTitle = 'Emergency Location State/Province'
            $EmergencyLocationStateMsg   = 'Please provide the Emergency Location State/Province (e.g. Berkshire):'
            $EmergencyLocationState = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationStateTitle, $EmergencyLocationStateMsg)
            
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationCountryTitle = 'Emergency Location Country/Region'
            $EmergencyLocationCountryMsg   = 'Please provide the Emergency Location Country/Region (e.g. GB):'
            $EmergencyLocationCountry = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationCountryTitle, $EmergencyLocationCountryMsg)
    
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationPostCodeTitle = 'Emergency Location Postal Code'
            $EmergencyLocationPostCodeMsg   = 'Please provide the Emergency Location PostCode (e.g. RG6 1WG):'
            $EmergencyLocationPostCode = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationPostCodeTitle, $EmergencyLocationPostCodeMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationLatitudeTitle = 'Emergency Location Latitude'
            $EmergencyLocationLatitudeMsg   = 'Please provide the Emergency Location Latitude (e.g. 51.060692):'
            $EmergencyLocationLatitude = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationLatitudeTitle, $EmergencyLocationLatitudeMsg)
    
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $EmergencyLocationLongitudeTitle = 'Emergency Location Longitude'
            $EmergencyLocationLongitudeMsg   = 'Please provide the Emergency Location Longitude (e.g. -1.313154):'
            $EmergencyLocationLongitude = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyLocationLongitudeTitle, $EmergencyLocationLongitudeMsg)


            New-CsOnlineLisCivicAddress -HouseNumber $EmergencyLocationHouse -StreetName $EmergencyLocationStreet -City $EmergencyLocationCity -StateorProvince $EmergencyLocationState -CountryOrRegion $EmergencyLocationCountry -PostalCode $EmergencyLocationPostCode -Description $EmergencyLocationDesc -CompanyName $EmergencyLocationCompany -Latitude $EmergencyLocationLatitude -Longitude $EmergencyLocationLongitude
            $EmergencyLocation = Get-CsOnlineLisLocation | Select-Object CompanyName,Description, HouseNumber, StreetName, City, Postcode,CountryOrRegion, LocationID | Out-GridView -OutputMode Single -Title "Please select the newly created Emergency Location"
            
            Write-Host "Emergency Location: $EmergencyLocation.LocationID"
            Write-LogFileMessage "Emergency Location: $EmergencyLocation.LocationID"
        }
    }
### FUNCTION - Create Auto Attendant ###
function New-SharedCallingAutoAttendant
    {
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $SharedCallingAAGreetingTitle = 'Auto Attendant Greeting'
        $SharedCallingAAGreetingMsg   = 'Please provide the Auto Attendant Greeting (e.g. Welcome to Contoso!):'
        $SharedCallingAAGreeting = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAAGreetingTitle, $SharedCallingAAGreetingMsg)

        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $SharedCallingAAPromptTitle = 'Auto Attendant Prompt'
        $SharedCallingAAPromptMsg   = 'Please provide the Auto Attendant Prompt (e.g. If you know the extension you require, dial it now):'
        $SharedCallingAAPrompt = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingAAPromptTitle, $SharedCallingAAPromptMsg)

        $greetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $SharedCallingAAGreeting
        $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $SharedCallingAAPrompt
        $defaultMenu = New-CsAutoAttendantMenu -Name "Default menu" -Prompts @($menuPrompt) -EnableDialByName -DirectorySearchMethod ByName
        $defaultCallFlow = New-CsAutoAttendantCallFlow -Name "Default call flow" -Menu $defaultMenu -Greetings @($greetingPrompt)
        New-CsAutoAttendant -Name $SharedCallingAAName -DefaultCallFlow $defaultCallFlow -EnableVoiceResponse -LanguageId "en-GB" -TimeZoneId "GMT Standard Time" 

        ### Assign Resource Account
        $applicationInstanceId = (Get-CsOnlineUser $SharedCallingAAUPN).Identity
        $autoAttendantId = (Get-CsAutoAttendant -NameFilter $SharedCallingAAName).Id
        New-CsOnlineApplicationInstanceAssociation -Identities @($applicationInstanceId) -ConfigurationId $autoAttendantId -ConfigurationType AutoAttendant

        Write-Host "Auto Attendant: $SharedCallingAAName created"
        Write-LogFileMessage "Auto Attendant: $SharedCallingAAName created"
    }

### FUNCTION - Shared Calling Voice Configuration ###
function Set-SharedCallingVoiceConfiguration
    {
      $EmergencyCallRoutingPolicySelection = [System.Windows.Forms.MessageBox]::Show("Create a new Emergency Call Routing Policy?" , "Input Required" , 4, 64)

        if ($EmergencyCallRoutingPolicySelection -eq "No") {
          Write-Host "Existing Emergency Call Routing Policy selected."
          Write-LogFileMessage "Existing Emergency Call Routing Policy selected."

        }elseif ($EmergencyPolicySelection -eq "Yes") {
            
            Write-Host "New Emergency Call Routing Policy selected."
            Write-LogFileMessage "New Emergency Call Routing Policy selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingEmergencyRoutingPolicyNameTitle = 'Emergency Dial Routing Policy Name'
            $SharedCallingEmergencyRoutingPolicyNameMsg   = 'Please provide the Emergency Call Routing Policy Name (e.g. UK-ECRP):'
            $SharedCallingEmergencyRoutingPolicyName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingEmergencyRoutingPolicyNameTitle, $SharedCallingEmergencyRoutingPolicyNameMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingEmergencyDialStringTitle = 'Emergency Dial String'
            $SharedCallingEmergencyDialStringMsg   = 'Please provide the Emergency Dial String (e.g. 999):'
            $SharedCallingEmergencyDialString = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingEmergencyDialStringTitle, $SharedCallingEmergencyDialStringMsg)

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingEmergencyDialMaskTitle = 'Emergency Dial Mask'
            $SharedCallingEmergencyDialMaskMsg   = 'Please provide the Emergency Dial Mask (e.g. 999):'
            $SharedCallingEmergencyDialMask = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingEmergencyDialMaskTitle, $SharedCallingEmergencyDialMaskMsg)
            
            $en1 = New-CsTeamsEmergencyNumber -EmergencyDialString $SharedCallingEmergencyDialString -EmergencyDialMask $SharedCallingEmergencyDialMask
            New-CsTeamsEmergencyCallRoutingPolicy -Identity $SharedCallingEmergencyRoutingPolicyName -EmergencyNumbers @{add=$en1} -AllowEnhancedEmergencyServices:$true

            Write-Host "New emergency Call Routing Policy: $EmergencyCallRoutingPolicy"
            Write-LogFileMessage "New emergency Call Routing Policy: $EmergencyCallRoutingPolicy"
        }

        $VoiceRoutingPolicySharedCallingSelection = [System.Windows.Forms.MessageBox]::Show("Create a new Shared Calling Voice Routing Policy?" , "Shared Calling Voice Routing Policy creation" , 4, 64)

        if ($VoiceRoutingPolicySharedCallingSelection -eq "No") {
          Write-Host "Existing Shared Calling Voice Routing Policy selected."
          Write-LogFileMessage "Existing Shared Calling Voice Routing Policy selected."

        }elseif ($VoiceRoutingPolicySharedCallingSelection -eq "Yes") {
            
            Write-Host "New Shared Calling Voice Routing Policy selected."
            Write-LogFileMessage "Shared Calling Voice Routing Policy selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallingVRPNameTitle = 'Emergency Dial Routing Policy Name'
            $SharedCallingVRPNameMsg   = 'Please provide the Emergency Call Routing Policy Name (e.g. UK-ECRP):'
            $SharedCallingVRPName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallingVRPNameTitle, $SharedCallingVRPNameMsg)
                       
            New-CSOnlineVoiceRoutingPolicy -Identity $SharedCallingVRPName

            Write-Host "New Shared Calling Voice Routing Policy: $SharedCallingVRPName"
            Write-LogFileMessage "New Shared Calling Voice Routing Policy: $SharedCallingVRPName"
        }

        $CallerIDSelection = [System.Windows.Forms.MessageBox]::Show("Create a new Caller ID Policy?" , "Caller ID creation" , 4, 64)

        if ($CallerIDSelection -eq "No") {
          Write-Host "Existing Caller ID Policy selected."
          Write-LogFileMessage "Existing Caller ID Policy selected."

        }elseif ($CallerIDSelection -eq "Yes") {
            
            Write-Host "New Caller ID Policy selected."
            Write-LogFileMessage "Shared Caller ID Policy selected."

            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $SharedCallerIDNameTitle = 'Caller ID Policy Name'
            $SharedCallerIDNameMsg   = 'Please provide the Caller ID Policy Name (e.g. Shared Calling):'
            $SharedCallerIDName = [Microsoft.VisualBasic.Interaction]::InputBox($SharedCallerIDNameTitle, $SharedCallerIDNameMsg)
                       
            $ObjId = (Get-CsOnlineApplicationInstance -Identity $SharedCallingAAUPN).ObjectId
            New-CsCallingLineIdentity -Identity $SharedCallerIDName -CallingIDSubstitute Resource -EnableUserOverride $false -ResourceAccount $ObjId

            Write-Host "New Shared Calling Voice Routing Policy: $SharedCallerIDName"
            Write-LogFileMessage "New Shared Calling Voice Routing Policy: $SharedCallerIDName"
        }

        $SharedCallingPolicySelection = [System.Windows.Forms.MessageBox]::Show("Do you require emergency callback numbers?" , "Shared Calling Policy Creation" , 4, 64)

        if ($SharedCallingPolicySelection -eq "Yes") {
          Write-Host "Emergency callback numbers required."
          Write-LogFileMessage "Emergency callback numbers required."

          [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
          $EmergencyNumber1Title = 'Emergency callback number 1'
          $EmergencyNumber1Msg   = 'Please provide the first Emergency callback number in E.164 format (e.g. +441632960999):'
          $EmergencyNumber1 = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyNumber1Title, $EmergencyNumber1Msg)

          [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
          $EmergencyNumber2Title = 'Emergency callback number 2'
          $EmergencyNumber2Msg   = 'Please provide the second Emergency callback number in E.164 format (e.g. +441632960999):'
          $EmergencyNumber2 = [Microsoft.VisualBasic.Interaction]::InputBox($EmergencyNumber2Title, $EmergencyNumber2Msg)

          New-CsTeamsSharedCallingRoutingPolicy -Identity $SharedCallingAAName -ResourceAccount $SharedCallingRA.Identity -EmergencyNumbers @{add=$EmergencyNumber1,$EmergencyNumber2}

          Write-Host "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"
          Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"

        }elseif ($SharedCallingPolicySelection -eq "No") {
            
            Write-Host "Emergency callback numbers not required."
            Write-LogFileMessage "Emergency callback numbers not required."

            $SharedCallingRA = Get-CsOnlineUser -Identity $SharedCallingAAUPN
            New-CsTeamsSharedCallingRoutingPolicy -Identity $SharedCallingAAName -ResourceAccount $SharedCallingRA.Identity
            
            Write-Host "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"
            Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"
            
        }

    }

### FUNCTION - Configure Shared Calling for Direct Routing numbers ###
function New-SharedCallingDirectRoutingConfig
    {
        New-SharedCallingResourceAccount
        
        Write-Host "Shared Calling Resource Accounts tasks completed." -ForegroundColor Green
        Write-Host "Shared Calling Resource account created: $SharedCallingAAName"
        Write-LogFileMessage "Shared Calling Resource Accounts tasks completed."
        Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"

        Set-SharedCallingEmergencyAddress

        Write-Host "Shared Calling Emergency Address tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Emergency Address tasks completed."

        $RAVoiceRoutingPolicy = Get-CsOnlineVoiceRoutingPolicy | Select-Object Identity, Description | Out-GridView -OutputMode Single -Title "Please select a Voice Routing Policy"
       
        Write-Host $RAVoiceRoutingPolicy.Identity "is your chosen Voice Routing Policy"
        Write-LogFileMessage $RAVoiceRoutingPolicy.Identity "is your chosen Voice Routing Policy"

        Grant-CsOnlineVoiceRoutingPolicy -PolicyName $RAVoiceRoutingPolicy -Identity $SharedCallingAAUPN

        Write-Host "Shared Calling Voice Routing Policy tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Routing Policy tasks completed."

        New-SharedCallingAutoAttendant

        Write-Host "Shared Calling Auto Attendant tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Auto Attendant tasks completed."
        
        Set-CsPhoneNumberAssignment -Identity $SharedCallingAAUPN -LocationID $LocationID.Identity -PhoneNumber $SharedCallingAANumber -PhoneNumberType DirectRouting
        
        Write-Host "Shared Resource Account Phone Number ($SharedCallingAANumber) assigned to $SharedCallingAAUPN." -ForegroundColor Green
        Write-LogFileMessage "Shared Resource Account Phone Number ($SharedCallingAANumber) assigned to $SharedCallingAAUPN."

        Set-SharedCallingVoiceConfiguration

        Write-Host "Shared Calling Voice Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Configuration tasks completed."
    }

### FUNCTION - Configure Shared Calling for Calling Plans numbers ###
function New-SharedCallingCallingPlansConfig
    {
        New-SharedCallingResourceAccount
        
        Write-Host "Shared Calling Resource Accounts tasks completed." -ForegroundColor Green
        Write-Host "Shared Calling Resource account created: $SharedCallingAAName"
        Write-LogFileMessage "Shared Calling Resource Accounts tasks completed."
        Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"

        Set-SharedCallingEmergencyAddress

        Write-Host "Shared Calling Emergency Address tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Emergency Address tasks completed."

      ### Assign Pay-as-you go calling plan (Zone-1 countries i.e. UK)
        $PAYGCallingZone1LicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'Microsoft_Teams_Calling_Plan_pay_as_you_go_(country_zone_1)'
        Set-MgUserLicense -UserId $SharedCallingAAUPN -addLicenses @{SkuId = $PAYGCallingZone1LicenseSku.SkuId} -RemoveLicenses @()
        
      ### PAUSE - wait a few minutes for the above cmdlet configuration to be completed
        Set-ScriptSleep 180
      
      ### Assign communication credits       
        $CommunicationCreditsLicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'MCOPSTNC'
        Set-MgUserLicense -UserId $SharedCallingAAUPN -addLicenses @{SkuId = $CommunicationCreditsLicenseSku.SkuId} -RemoveLicenses @()

        Write-Host "Calling Plans licensing tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Calling Plans licensing tasks completed."

        New-SharedCallingAutoAttendant

        Write-Host "Shared Calling Auto Attendant tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Auto Attendant tasks completed."

        Set-CsPhoneNumberAssignment -Identity $SharedCallingAAUPN -LocationID $LocationID.Identity -PhoneNumber $SharedCallingAANumber -PhoneNumberType CallingPlan
        Write-Host "Shared Resource Account Phone Number ($SharedCallingAANumber) assigned to $SharedCallingAAUPN." -ForegroundColor Green
        Write-LogFileMessage "Shared Resource Account Phone Number ($SharedCallingAANumber) assigned to $SharedCallingAAUPN."

        Set-SharedCallingVoiceConfiguration

        Write-Host "Shared Calling Voice Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Configuration tasks completed."
    }
### FUNCTION - Configure Shared Calling for Operator Connect numbers ###
function New-SharedCallingOperatorConnectConfig
    {
        New-SharedCallingResourceAccount
        
        Write-Host "Shared Calling Resource Accounts tasks completed." -ForegroundColor Green
        Write-Host "Shared Calling Resource account created: $SharedCallingAAName"
        Write-LogFileMessage "Shared Calling Resource Accounts tasks completed."
        Write-LogFileMessage "New Shared Calling Voice Routing Policy created: $SharedCallingAAName"

        Set-SharedCallingEmergencyAddress

        Write-Host "Shared Calling Emergency Address tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Emergency Address tasks completed."

        New-SharedCallingAutoAttendant

        Write-Host "Shared Calling Auto Attendant tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Auto Attendant tasks completed."

        Set-CsPhoneNumberAssignment -Identity $SharedCallingAAUPN -LocationID $LocationID.Identity -PhoneNumber $SharedCallingAANumber -PhoneNumberType OperatorConnect
        
        Write-Host "Shared Resource Account Phone Number ($SharedCallingAANumber) assigned to $SharedCallingAAUPN." -ForegroundColor Green
        Write-LogFileMessage "Shared Resource Account Phone Number ($SharedCallingAANumber) assigned to $SharedCallingAAUPN."

        Set-SharedCallingVoiceConfiguration

        Write-Host "Shared Calling Voice Configuration tasks completed." -ForegroundColor Green
        Write-LogFileMessage "Shared Calling Voice Configuration tasks completed."

    }

################### END OF SHARED CALLING TASK FUNCTIONS ###################

################### START OF COMMANDS ###################

ShowDisclaimer

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
    Write-Host "AzureADPreview module does not exist. Please run pre-requisite script."
    Write-LogFileMessage "AzureADPreview module does not exist. Please run pre-requisite script."
    Install-AzureADPreview
  }
  Else {
    Write-Host "AzureADPreview module exists. Please run pre-requisite script."
    Write-LogFileMessage "AzureADPreview module exists. Please run pre-requisite script."
  }
  
#Microsoft.Graph
  If (-not(Get-InstalledModule Microsoft.Graph -ErrorAction silentlycontinue)) {
    Write-Host "Microsoft.Graph module does not exist. Please run pre-requisite script."
    Write-LogFileMessage "Microsoft.Graph module does not exist. Please run pre-requisite script."
    Install-MicrosoftGraph
  }
  Else {
    Write-Host "Microsoft.Graph module exists. Please run pre-requisite script."
    Write-LogFileMessage "Microsoft.Graph module exists. Please run pre-requisite script."
  }
  
#MSOnline
  If (-not(Get-InstalledModule MSOnline -ErrorAction silentlycontinue)) {
    Write-Host "MSOnline module does not exist. Please run pre-requisite script."
    Write-LogFileMessage "MSOnline module does not exist. Please run pre-requisite script."
    Install-MSOnline
  }
  Else {
    Write-Host "MSOnline module exists. Please run pre-requisite script."
    Write-LogFileMessage "MSOnline module exists. Please run pre-requisite script."
  }
  
#ExchangeOnlineManagement
  If (-not(Get-InstalledModule ExchangeOnlineManagement -ErrorAction silentlycontinue)) {
    Write-Host "ExchangeOnlineManagement module does not exist. Please run pre-requisite script."
    Write-LogFileMessage "ExchangeOnlineManagement module does not exist. Please run pre-requisite script."
    Install-ExchangeOnlineManagement
  }
  Else {
    Write-Host "ExchangeOnlineManagement module exists. Please run pre-requisite script."
    Write-LogFileMessage "ExchangeOnlineManagement module exists. Please run pre-requisite script."
  }

#MicrosoftTeams
  If (-not(Get-InstalledModule MicrosoftTeams -ErrorAction silentlycontinue)) {
    Write-Host "MicrosoftTeams module does not exist. Please run pre-requisite script."
    Write-LogFileMessage "MicrosoftTeams module does not exist. Please run pre-requisite script."
    Install-MicrosoftTeams
  }
  Else {
    Write-Host "MicrosoftTeams module exists. Please run pre-requisite script."
    Write-LogFileMessage "MicrosoftTeams module exists. Please run pre-requisite script."
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

