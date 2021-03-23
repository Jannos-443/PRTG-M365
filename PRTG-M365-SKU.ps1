<#   
    .SYNOPSIS
    Monitors Microsoft 365 License usage using Microsoft Graph API

    .DESCRIPTION
    Using MS Graph this Script shows the Microsoft 365 License usage
    You can display all licenses, include only some or exclude some.

    Copy this script to the PRTG probe EXEXML scripts folder (${env:ProgramFiles(x86)}\PRTG Network Monitor\Custom Sensors\EXEXML)
    and create a "EXE/Script Advanced. Choose this script from the dropdown and set at least:

    + Parameters: TenatDomainName, ApplicationID, AccessSecret
    + Scanning Interval: minimum 15 minutes

    .PARAMETER TenatDomainName
    your Microsoft 365 TenantName for Example contoso.onmicrosoft.com

    .PARAMETER ApplicationID
    Provide the ApplicationID

    .PARAMETER AccessSecret
    Provide the Application Secret

    .PARAMETER exclude
    true = exclude $SKUPattern SKUs
    false (default) = only include $SKUPattern SKUs

    .PARAMETER SKUPattern
    Regular expression to select the SKUs

    exclude = $false (default)
        Example: '^(Enterprisepack)$' includes only Enterprisepack (Office 365 E3)
        Example2: '^(Enterprisepack|EMS)$' includes only Enterprisepack (Office 365 E3) and EMS (Enterprise Mobility + Security)

    exclude = $true
        Example: '^(Enterprisepack)$' includes all but Enterprisepack (Office 365 E3)
        Example2: '^(Enterprisepack|EMS)$' includes all but Enterprisepack (Office 365 E3) and EMS (Enterprise Mobility + Security)
    
    Regular Expression: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1
    License Names: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced

    "PRTG-M365-SKU.ps1" -ApplicationID 'Test-APPID' -TenatDomainName 'contoso.onmicrosoft.com' -AccessSecret 'Test-AppSecret' -SKUPattern '(Enterprisepack|EMS|ATP_ENTERPRISE)'

    Microsoft 365 Permission:
        1. Open Azure AD
        2. Register new APP
        3. Overview >> Get Application ID 
        4. Set API Permissions >> MS Graph >> Application >> Organization.Read.All
        5. Certificates & secrets >> new Secret >> unlimited

    Author:  Jannos-443
    https://github.com/Jannos-443/PRTG-M365-SKU
#>
param(
    [Parameter(Mandatory)] [string] $TenatDomainName = '',
    [Parameter(Mandatory)] [string] $ApplicationID = '',
    [Parameter(Mandatory)] [string] $AccessSecret = '',
    [string] $SKUPattern = '',
    [boolean] $exclude = $false
)


#Catch all unhandled Errors
$ErrorActionPreference = "Stop"
trap
    {
    Write-Error $_.ToString()
    Write-Error $_.ScriptStackTrace
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>$($_.ToString() - $($_.ScriptStackTrace))</text>"
    Write-Output "</prtg>"
    Exit
    }


# Get MS Graph Token
try 
    {
    #Check if Token is expired
    $renew = $false

    if($ConnectGraph)
        {
        if((get-date).AddMinutes(2) -ge $tokenexpire)
            {
            Write-Host "Token expired or close to expire, going to renew Token"
            $renew = $true
            }

        else
            {
            Write-Host "Token found and still valid"
            }


        }

    else
        {
        $renew = $true
        Write-Host "Token not found, going to renew Token"
        }



    if($renew)
        {
        #Request Token
        $Body = @{    
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        client_Id     = $ApplicationID
        Client_Secret = $AccessSecret
        } 
  
        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenatDomainName/oauth2/v2.0/token" -Method POST -Body $Body
        $token = $ConnectGraph.access_token
        $tokenexpire = (Get-Date).AddSeconds($ConnectGraph.expires_in)

        Write-Host "sucessfully got new MS Graph Token"
        }
    } 
    
catch 
    {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error getting MS Graph Token ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
    }



#Get List of Azure Licenses
try 
    {
    #Get All SKUs
    $Headers = @{Authorization = "$($ConnectGraph.token_type) $($ConnectGraph.access_token)"}
   
    $GraphUrl = "https://graph.microsoft.com/v1.0/subscribedSkus" #User URL
    
    $SKUs = Invoke-RestMethod -Headers $Headers -Uri $GraphUrl -Method Get
      
    $AllSKUs = $SKUs.value
    while($SKUs.'@odata.nextLink')
        {
        $graphURL = $SKUs.'@odata.nextLink'
        $SKUs = Invoke-RestMethod -Headers $Headers -Uri $GraphUrl -Method Get
        $AllSKUs = $AllSKUs + $SKUs.value

        }
    } 

catch 
    {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Could not MS Graph $($GraphUrl). Error: $($_.Exception.Message)</text>"
    Write-Output "</prtg>"
    Exit
    }


#Filter LICs
#Remove all SKUs with Zero Licenses
$AllSKUs = $AllSKUs | where {$_.consumedUnits -gt 0}

#Use Exclude
if($exclude)
    {
    if ($SKUPattern -ne "") 
        {
        $AllSKUs = $AllSKUs | where {$_.SKupartnumber -notmatch $SKUPattern}  
        }
    }

#Use Include
else
    {
    if ($SKUPattern -ne "") 
        {
            $AllSKUs = $AllSKUs | where {$_.SKupartnumber -match $SKUPattern}  
        }
    }

$xmlOutput = '<prtg>'

foreach($LIC in $AllSKUs)
    {
    $xmlOutput = $xmlOutput + "
        <result>
        <channel>$($LIC.SkuPartNumber) - Free Licenses</channel>
        <value>$($LIC.PrepaidUnits.Enabled - $LIC.ConsumedUnits)</value>
        <unit>Count</unit>
        </result>

        <result>
        <channel>$($LIC.SkuPartNumber) - Total Licenses</channel>
        <value>$($LIC.PrepaidUnits.Enabled)</value>
        <unit>Count</unit>
        </result>
        "
    }


$xmlOutput = $xmlOutput + "</prtg>"

$xmlOutput