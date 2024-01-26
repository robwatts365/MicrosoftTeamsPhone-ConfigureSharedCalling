# Deployment Guide

## Pre-requisites
* PowerShell modules:
  * AzureADPreview
  * Microsoft.Graph
  * MSOnline
  * ExchangeOnlineManagment
  * MicrosoftTeams
 > [!NOTE]
  >  These will be installed  for you by the [pre-requisites script](1-ConfigureSharedCalling-PreReqs.ps1)

*	Microsoft Teams Phone Resource Account licensing
*	Shared telephone phone number for inbound and outbound calling (Must be enabled for Voice app usage)
*	Microsoft Phone System licensing for users (E5 or other licences with Phone System Standard)
*	Optional - telephone number(s) for emergency callback

### For Direct Routing deployments:
* Your SBC is deployed and working correctly.
* Teams voice routing configuration has been completed.
### For Calling Plans deployments:
* Pay-as-you go (PAYG) calling plan is purchased and ready to be assigned.
* Communication credits licence is purchased and has funds.
### For Operator Connect deployments:
* You have an enabled Operator Connect carrier.

1. If required, run the Configure-SharedCalling-PreReqs PowerShell script - [1-ConfigureSharedCalling-PreReqs.ps1](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/blob/main/1-ConfigureSharedCalling-PreReqs.ps1)  
This PowerShell file requires **Run as Administrator**, if the script is not run as Administrator, it will prompt you accordinfly.  
  
![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/6d484f43-f135-467e-9484-28981d4712e9)

This script will check for pre-requisite PowerShell modules and install these where necessary.  
    
![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/7a76bb52-57fd-41f1-8875-6e5c0b53def3)  
  
Once you have all required PowerShell modules installed, you may move on to the next step.

## 2. Configuration
1. Run the Configure-SharedCalling PowerShell script - [2-Configure-SharedCalling.ps1](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/blob/main/2-Configure-SharedCalling.ps1)  
2. The script will start by asking you to select a location to save the log file. Create an appropriate name for your log file and press "Save"
3. Next the script will check whether the required PowerShell modules exist
4. Finally you'll be asked for your chosen telephony configuration option. Select the appropriate option by Typing the corresponding number.
   ![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/264e75ae-f337-412e-b56a-e431106aac33)
5. Once selected, the PowerShell modules will be connected. 
 > [!TIP]
  >  Be patient! There's five modules to be connected, you should be prompted for Modern Authentication four times.
6. Once connected, the configuration can begin. You'll first be asked for the domain name you'd like to use. e.g. **contoso.com**
   ![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/78e3d743-a422-4e7d-bd96-134da0d46f81)
7. Next you'll be asked for the Auto Attendant name. This is used throughout the configuration process.
   ![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/cae382df-ff7c-457e-9a22-e5aba0802ef9)
8. You'll then be asked for the telephone number you wish to use.
   ![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/79366712-9af6-4b9a-8f37-8916824dc675)
> [!IMPORTANT]
  >  This needs to be in E.164 format. Please check your telepghone number before pressing Ok to continue.
9. The Auto Attendant Resource Account will now be created and appropriately licensed.
> This process does take about 4 mins, as the script will sleep for 2 minutes between tasks, to allow for changes to be synced.

> [!NOTE]
> Dependant on the telephony configuration selection the flow may be different, but the same principles apply for each. 
> You'll be given the option to use an existing policy or create a new one. 
> If you select an existing policy, you'll recieve a pop out asking you select the policy
> If you select new, you'll be prompted for the information required.

**Here's an example of what you may see:**  

   * Request pop out  
    ![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/8960f163-b0b9-46cd-ad4a-66d771cbe1a5)
<br></br>
   * Gridview to select a current policy  
    ![image](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/assets/65971102/9f3a7839-95c5-4f7a-98df-fc616f829dcc)

## 3. Enabling Users  

1. 

| Page | Deployment Guide |
| :--- | :--- |
| Author | Rob Watts ([@robwatts365](https://github.com/robwatts365)) |
| **Version** | 1.1 |
| **Date** | 26/01/2024 |
