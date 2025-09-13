# Configure Shared Calling for Microsoft Teams Phone

| [Home](README.md) | [About](about.md) | [Considerations](considerations.md) | [Deployment Guide](deployment.md) | [Support](support.md) | 
| --- | --- | --- | --- | --- |

## Background
This project was started to help Microsoft customers deploy Shared Calling for Microsoft Teams Phone. The project consists of interactive powershell scripts that guide IT admins through the process of setting up Shared Calling features for Microsoft Teams Phone. Here you can find the source code, documentation, and instructions on how to use the scripts. It is recommended that you download the latest release of the project directly from this GitHub repository. Find the latest release [here](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/releases).

The project aims to make it easier and faster for IT admins to configure and deploy shared calling for Microsoft Teams Phone, which is a feature that allows multiple users to share a phone number and make or receive calls on behalf of a group or a department. The scripts automate the steps required to create and assign shared calling policies, create and configure resource accounts, assign phone numbers and licenses, and enable or disable shared calling for users.

The project is open source and welcomes contributions from the community. 

## Documentation
For further guidance about Shared Calling in Teams, please see this [configuration guide](https://aka.ms/TeamsSharedCallingConfigGuide) by [@ariprotheroe](https://github.com/ariprotheroe) and the [Microsot Learn page](https://learn.microsoft.com/en-us/microsoftteams/shared-calling-setup).

## Support
If you encounter any bugs or issues while using the scripts, please raise them as issues [here](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/issues) in this GitHub repository. You can also submit pull requests with your suggestions or improvements to the code. The project team appreciates your feedback and support. 

Thank you for using this project and I hope you find it useful and helpful. 😊

 > [!TIP]
> As a starting place, it's best to review the [Considerations for deploying Shared Calling](considerations.md).

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

## Deployment Guide
Please find step-by-step guidance on the [deployment guide](deployment.md) page.  
If you think something is missing, please raise them as issues [here](https://github.com/robwatts365/MicrosoftTeamsPhone-ConfigureSharedCalling/issues).

## Page info

| Name | README.md |
| :--- | :--- |
| Author | Rob Watts ([@robwatts365](https://github.com/robwatts365)) |
| **Version** | 1.1 |
| **Date** | 26/01/2024 |
