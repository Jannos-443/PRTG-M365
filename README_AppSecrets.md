<!-- ABOUT THE PROJECT -->
### About The Project
Project Owner: Jannos-443

PRTG Powershell Script to monitor Microsoft 365 App Secret expiration

Free and open source: [MIT License](https://github.com/Jannos-443/PRTG-M365/blob/main/LICENSE)

**Features**
* Monitor Office365 Application Secret expiration
* Certs expiration
* Passwords expiration

## HOW TO

1. Create AzureAD App
   - Open Azure AD
   - Register new APP
   - Overview >> Get Application ID 
   - Set API Permissions >> MS Graph >> Application >> Application.Read.All
   - Certificates & secrets >> new Secret
or follow this Guide: [Paessler M365 Credentials](https://kb.paessler.com/en/topic/88462-how-do-i-obtain-credentials-and-set-permissions-for-the-microsoft-365-sensors)

2. Place `PRTG-M365-AppSecrets.ps1` under `C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\EXEXML`

3. Create new Sensor 
   - EXE/Script Advanced = `PRTG-M365-AppSecrets.ps1`
   - Parameter = `-ApplicationID 'Test-APPID' -TenatDomainName 'contoso.onmicrosoft.com' -AccessSecret 'Test-AppSecret'`

