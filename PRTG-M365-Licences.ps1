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

    without "-exclude"
        Example: '^(Enterprisepack)$' includes only Enterprisepack (Office 365 E3)
        Example2: '^(Enterprisepack|EMS)$' includes only Enterprisepack (Office 365 E3) and EMS (Enterprise Mobility + Security)

    with "-exclude"
        Example: '^(Enterprisepack)$' includes all but Enterprisepack (Office 365 E3)
        Example2: '^(Enterprisepack|EMS)$' includes all but Enterprisepack (Office 365 E3) and EMS (Enterprise Mobility + Security)

    Regular Expression: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1
    License Names: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced

    "PRTG-M365-Licences.ps1" -ApplicationID 'Test-APPID' -TenatDomainName 'contoso.onmicrosoft.com' -AccessSecret 'Test-AppSecret' -SKUPattern '^(Enterprisepack|EMS|ATP_ENTERPRISE)$'

    Microsoft 365 Permission:
        1. Open Azure AD
        2. Register new APP
        3. Overview >> Get Application ID
        4. Set API Permissions >> MS Graph >> Application >>
           - Organization.Read.All
           - Group.Read.All
        5. Certificates & secrets >> new Secret

    Author:  Jannos-443
    https://github.com/Jannos-443/PRTG-M365
#>
param(
    [string] $TenatDomainName = '',
    [string] $ApplicationID = '',
    [string] $AccessSecret = '',
    [string] $SKUPattern = '',
    [Switch] $exclude,
    [Switch] $Hide_LicenceCount,
    [Switch] $Hide_GroupBasedLicence,
    [Switch] $Hide_LastDirSync
)

#Catch all unhandled Errors
$ErrorActionPreference = "Stop"
trap {
    $Output = "line:$($_.InvocationInfo.ScriptLineNumber.ToString()) char:$($_.InvocationInfo.OffsetInLine.ToString()) --- message: $($_.Exception.Message.ToString()) --- line: $($_.InvocationInfo.Line.ToString()) "
    $Output = $Output.Replace("<", "")
    $Output = $Output.Replace(">", "")
    $Output = $Output.Replace("#", "")
    Write-Output "<prtg>"
    Write-Output "<error>1</error>"
    Write-Output "<text>$Output</text>"
    Write-Output "</prtg>"
    Exit
}

#region set TLS to 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#endregion

if (($TenatDomainName -eq "") -or ($Null -eq $TenatDomainName)) {
    Throw "TenantDomainName Variable is empty"
}

if (($ApplicationID -eq "") -or ($Null -eq $ApplicationID)) {
    Throw "ApplicationID Variable is empty"
}

if (($AccessSecret -eq "") -or ($Null -eq $AccessSecret)) {
    Throw "AccessSecret Variable is empty"
}

# Get MS Graph Token
try {
    #Check if Token is expired
    $renew = $false

    if ($ConnectGraph) {
        if ((get-date).AddMinutes(2) -ge $tokenexpire) {
            Write-Host "Token expired or close to expire, going to renew Token"
            $renew = $true
        }

        else {
            Write-Host "Token found and still valid"
        }


    }

    else {
        $renew = $true
        Write-Host "Token not found, going to renew Token"
    }



    if ($renew) {
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

catch {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error getting MS Graph Token ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
}

$xmlOutput = '<prtg>'

#Function Graph API Call
Function GraphCall($URL) {
    #MS Graph Request
    try {
        $Headers = @{Authorization = "$($ConnectGraph.token_type) $($ConnectGraph.access_token)" }
        $GraphUrl = $URL
        $Result_Part = Invoke-RestMethod -Headers $Headers -Uri $GraphUrl -Method Get
        $Result = $Result_Part.value
        while ($Result_Part.'@odata.nextLink') {
            $graphURL = $Result_Part.'@odata.nextLink'
            $Result_Part = Invoke-RestMethod -Headers $Headers -Uri $GraphUrl -Method Get
            $Result = $Result + $Result_Part.value
        }
    }

    catch {
        Write-Output "<prtg>"
        Write-Output " <error>1</error>"
        Write-Output " <text>Could not MS Graph $($GraphUrl). Error: $($_.Exception.Message)</text>"
        Write-Output "</prtg>"
        Exit
    }
    return $Result
}

if ($Hide_LastDirSync -eq $false) {
    $Result = GraphCall -URL "https://graph.microsoft.com/v1.0/organization"

    if ($Result.onPremisesSyncEnabled) {
        $DirSyncEnabled = $true
    }


    if ($DirSyncEnabled) {
        [DateTime]$DirSyncTime = $result.onPremisesLastSyncDateTime
        $LastSyncTime = (((Get-Date).ToUniversalTime()) - ($DirSyncTime.ToUniversalTime())).TotalSeconds
        $LastSyncTime = [System.Math]::Round($LastSyncTime, 0)

        $xmlOutput = $xmlOutput + "
            <result>
            <channel>LastDirSync</channel>
            <value>$($LastSyncTime)</value>
            <unit>TimeSeconds</unit>
            <limitmode>1</limitmode>
            <LimitMaxError>7200</LimitMaxError>
            </result>"
    }
}

if ($Hide_GroupBasedLicence -eq $false) {
    #Get groups with licence Error
    $Result = GraphCall -URL "https://graph.microsoft.com/v1.0/groups?`$filter=hasMembersWithLicenseErrors+eq+true"

    #Get license error from group
    $LicenceErrorCount = 0
    $LicenceText = "Groups with Licence Errors: "
    foreach ($group in $Result) {
        $Errors = $null
        $Errors = GraphCall -URL "https://graph.microsoft.com/v1.0/groups/$($group.id)/membersWithLicenseErrors?`$select=licenseAssignmentStates"
        $LicenceErrorCount += ($Errors | Measure-Object).Count
        $ErrorText = ($Errors.licenseAssignmentStates | Where-Object { $_.assignedByGroup -eq $group.id })[0].error
        $LicenceText += "group: $($group.displayName) error: $($ErrorText)"
    }
    if ($LicenceErrorCount -gt 0) {
        $xmlOutput = $xmlOutput + "<text>$($LicenceText)</text>"
    }

    $xmlOutput = $xmlOutput + "
    <result>
    <channel>GroupBasedLicenceError</channel>
    <value>$($LicenceErrorCount)</value>
    <unit>Count</unit>
    <limitmode>1</limitmode>
    <LimitMaxError>0</LimitMaxError>
    </result>"
}

if ($Hide_LicenceCount -eq $false) {
    $Result = GraphCall -URL "https://graph.microsoft.com/v1.0/subscribedSkus"

    #Filter LICs
    #Remove all SKUs with Zero Licenses
    $Result = $Result | Where-Object { $_.consumedUnits -gt 0 }

    #Use Exclude
    if ($exclude) {
        if ($SKUPattern -ne "") {
            $Result = $Result | Where-Object { $_.SKupartnumber -notmatch $SKUPattern }
        }
    }

    #Use Include
    else {
        if ($SKUPattern -ne "") {
            $Result = $Result | Where-Object { $_.SKupartnumber -match $SKUPattern }
        }
    }

    foreach ($LIC in $Result) {
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
}

$xmlOutput = $xmlOutput + "</prtg>"

$xmlOutput