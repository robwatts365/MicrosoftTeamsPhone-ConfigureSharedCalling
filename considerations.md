# Considerations

Before you begin there are a few things to consider:
1. How many shared telephone numbers will be required (i.e. one for the entire organization or by country or by location?).
2. Consider emergency calling specific to your country and/or location.
3. Consider how you will be assigning the Teams policy configuration to the users.  Either using the global configuration (tenant-wide to all users) or user assigned policy configuration.  This guide has been written assuming per-user assigned policies.
4. When using an operator connect carrier, ensure with the carrier that shared calling is supported. Alternatively, you can use acquire a new number with Microsoft and use a single pay-as-you-go calling plan and communication credits.
5. When using Direct Routing, ensure that there is adequate call/session capacity for Shared Calling users on both the SBC as well as SIP channels with the carrier.
6. How will inbound calls be connected to users via the Auto Attendant: 

    **a.** **Dial by Name** i.e. name lookup callers can say (speech recognition) or enter (keypad). either the full or partial name of the person they are trying to reach.

    They can also reach anyone in Active Directory by saying the full or partial name of the person they are trying to locate. Using voice inputs can recognize names in various formats, including FirstName, LastName, FirstName + LastName, or LastName + FirstName.

    Searches the entire organization's directory first before applying any Dial Scope Include or Exclude lists that have been configured.  If the initial search against the entire directory returns more than 100 users, the Dial Scope lists will not be applied, the search will fail, and the caller will be told that too many names were found.

    People can also use the keypad.  '0' (zero) key is used to indicate a space between the first name and last or last name and first. When they are entering the name, they should be asked to terminate their keypad entry with the # key. For example, "After you enter the name of the person you are trying to reach, press #."

    If there are multiple names that are found, the person calling will be given a list of names to select from.

    **b.** **Dial by Extension** (Number lookup)

    Users you want to make available for Dial By Extension need to have an extension specified as part of one of the following phone attributes defined in Active Directory (and synchronized via Azure AD Connect) or Azure Active Directory

-     TelephoneNumber (AD and Azure AD)
-     HomePhone (AD)
-     Mobile (AD and Azure AD)
-     OtherTelephone (AD)

    The required format to enter the extension in the user phone number field can be one of the following formats:

    | Format | Example |
    | --- | --- |
    | +[phone number];ext=[extension] | +441632960990;ext=60990 |
    | +[phone number]x[extension] | +441632960990x60990 |
    | x[extension] | x60990 |

7. User scoping for the auto attendant lookup (i.e. the Dial Scope who is include and who is excluded). A Microsoft 365 group is required.

8. Consider the calling charges for shared calling users. Calling tariff rates for PAYG calling plan can be found [here](https://www.microsoft.com/en-gb/microsoft-teams/microsoft-teams-phone).  Consider monthly reporting on usage to identify the top users. Then assign these users telephony connectivity with an inclusive call minutes bundle (i.e. a full calling plan or operator connect).

9. Consider how Shared Calling can be used as part of a wider telephony transformation strategy for telephony cost reduction.  

    | Telephony Usage | User Licencing with Microsoft Calling Plan | User Licencing with Operator Connect, Direct Routing, or Teams Phone Mobile | Telephone Number |
    | ------------- | ------------- | ------------- | ------------- |
    | Mobile | N/A | Teams Phone + add-on with operator | Yes (Mobile)
    | High Usage | Teams Phone + calling plan with include minutes | Teams Phone | Yes |
    | Medium/Low Usage (or require DDI) | Teams Phone + PAYG calling plan | Teams Phone | Yes |
    | Low/No Usage | Teams Phone | Teams Phone | No (Shared) 

| Page | Considerations |
| :--- | :--- |
| Author | Rob Watts ([@robwatts365](https://github.com/robwatts365)) |
| **Version** | 1.1 |
| **Date** | 26/01/2024 |