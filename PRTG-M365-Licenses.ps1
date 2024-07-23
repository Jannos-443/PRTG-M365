<#
    .SYNOPSIS
    Monitors Microsoft 365 License usage using Microsoft Graph API

    .DESCRIPTION
    Using MS Graph this Script shows the Microsoft 365 License usage
    You can display all licenses, include only some or exclude some.

    Copy this script to the PRTG probe EXEXML scripts folder (${env:ProgramFiles(x86)}\PRTG Network Monitor\Custom Sensors\EXEXML)
    and create a "EXE/Script Advanced. Choose this script from the dropdown and set at least:

    + Parameters: TenantID, ApplicationID, AccessSecret
    + Scanning Interval: minimum 15 minutes

    .PARAMETER TenantID
    Provide the TenantID or TenantName (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx or contoso.onmicrosoft.com)

    .PARAMETER ApplicationID
    Provide the ApplicationID

    .PARAMETER AccessSecret
    Provide the Application Secret

    .PARAMETER FriendlyName
    Use this switch if you want to display the friendlynames and not the SKU Names.
    With this Switch you are able to use "-IncludeName" and "-ExcludeName" to include and exclude based on the friendlynames

    At the moment the friendlynames are translatet by an csv downloaded from microsoft, there is no REST API option :(

    .PARAMETER IncludeName
    -FriendlyName required
    See IncludeSKU but you can use the friendlynames

    .PARAMETER ExcludeName
    -FriendlyName required
    See IncludeSKU but you can use the friendlynames

    .PARAMETER ExcludeSKU
    See IncludeSKU

    .PARAMETER IncludeSKU
    Use regular expression to include/exclude based on the stringnames (See Licence Name Link)

    Licence Names: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    Regular Expression: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1

    Examples:
    -IncludeSKU '^(Enterprisepack)$' includes only Enterprisepack (Office 365 E3)
    -IncludeSKU '^(Enterprisepack|EMS)$' includes only Enterprisepack (Office 365 E3) and EMS (Enterprise Mobility + Security)
    -ExcludeSKU '^(Enterprisepack)$' includes all but Enterprisepack (Office 365 E3)
    -ExcludeSKU '^(Enterprisepack|EMS)$' includes all but Enterprisepack (Office 365 E3) and EMS (Enterprise Mobility + Security)

    .NOTES
    Version 1.01

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced

    "PRTG-M365-Licenses.ps1" -ApplicationID 'Test-APPID' -TenantID 'contoso.onmicrosoft.com' -AccessSecret 'Test-AppSecret' -IncludeSKU '^(Enterprisepack|EMS|ATP_ENTERPRISE)$'

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
    [string] $TenantID = '',
    [string] $ApplicationID = '',
    [string] $AccessSecret = '',
    [string] $IncludeSKU = '',
    [string] $ExcludeSKU = '',
    [string] $IncludeName = '',
    [string] $ExcludeName = '',
    [Switch] $FriendlyName,
    [Switch] $Hide_LicenseCount,
    [Switch] $Hide_GroupBasedLicense,
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

if (($TenantID -eq "") -or ($Null -eq $TenantID)) {
    Throw "TenantID Variable is empty"
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
            #Write-Host "Token expired or close to expire, going to renew Token"
            $renew = $true
        }

        else {
            #Write-Host "Token found and still valid"
        }


    }

    else {
        $renew = $true
        #Write-Host "Token not found, going to renew Token"
    }



    if ($renew) {
        #Request Token
        $Body = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $ApplicationID
            Client_Secret = $AccessSecret
        }

        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token" -Method POST -Body $Body
        $token = $ConnectGraph.access_token
        $tokenexpire = (Get-Date).AddSeconds($ConnectGraph.expires_in)

        #Write-Host "sucessfully got new MS Graph Token"
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

        $xmlOutput += "
            <result>
            <channel>LastDirSync</channel>
            <value>$($LastSyncTime)</value>
            <unit>TimeSeconds</unit>
            <limitmode>1</limitmode>
            <LimitMaxError>7200</LimitMaxError>
            </result>"
    }
}

if ($Hide_GroupBasedLicense -eq $false) {
    #Get groups with License Error
    $Result = GraphCall -URL "https://graph.microsoft.com/v1.0/groups?`$filter=hasMembersWithLicenseErrors+eq+true"

    #Get license error from group
    $LicenseErrorCount = 0
    $LicenseText = "Groups with License Errors: "
    foreach ($group in $Result) {
        
        $MembersWithErrors = $null
        $MembersWithErrors = GraphCall -URL "https://graph.microsoft.com/v1.0/groups/$($group.id)/membersWithLicenseErrors"
        $LicenseText += "group: $($group.displayName) errors: $($MembersWithErrors.Count); "
        $LicenseErrorCount += 1
    }
    if ($LicenseErrorCount -gt 0) {
        $xmlOutput += "<text>$($LicenseText)</text>"
    }

    $xmlOutput += "
    <result>
    <channel>GroupBasedLicenseError</channel>
    <value>$($LicenseErrorCount)</value>
    <unit>Count</unit>
    <limitmode>1</limitmode>
    <LimitMaxError>0</LimitMaxError>
    </result>"
}

if ($Hide_LicenseCount -eq $false) {
    $Result = GraphCall -URL "https://graph.microsoft.com/v1.0/subscribedSkus"

    #Filter LICs
    #Remove all SKUs with Zero Licenses
    $Result = $Result | Where-Object { $_.consumedUnits -gt 0 }

    #friendlynames/Productnames are not available via API, but it´s possible to download a csv File to translate them.
    #https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    if ($FriendlyName) {
        $TempFile = New-TemporaryFile -Verbose
        Invoke-WebRequest -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" -OutFile $TempFile.FullName
        $LicCSV = Import-Csv -Path $TempFile.FullName -Delimiter ","
        Start-Sleep -Seconds 1
        Remove-Item -Path $TempFile.FullName -Verbose
    }

    $LicList = New-Object System.Collections.ArrayList

    foreach ($R in $Result) {
        $NewLic = "" | Select-Object "skuId", "skuPartNumber", "FriendlyName", "ConsumedUnits", "PrepaidUnitsEnabled"
        $NewLic.skuId = $R.skuId
        $NewLic.skuPartNumber = $R.Skupartnumber
        $NewLic.ConsumedUnits = $R.ConsumedUnits
        $NewLic.PrepaidUnitsEnabled = ($R.PrepaidUnits.enabled + $R.PrepaidUnits.warning + $R.PrepaidUnits.suspended)
        if ($FriendlyName) {
            $NewLic.FriendlyName = $LicCSV | Where-Object { ($_.GUID.Trim()) -eq $R.skuId } | Select-Object -First 1 -ExpandProperty Product_Display_Name
        }
        $null = $LicList.Add($NewLic)
    }

    #Include SKU´s
    if ($IncludeSKU -ne "") {
        $LicList = $LicList | Where-Object { $_.skuPartNumber -match $IncludeSKU }
    }

    #Exclude SKU´s
    if ($ExcludeSKU -ne "") {
        $LicList = $LicList | Where-Object { $_.skuPartNumber -notmatch $ExcludeSKU }
    }

    if ($FriendlyName) {
        #Include Names
        if ($IncludeName -ne "") {
            $LicList = $LicList | Where-Object { $_.FriendlyName -match $IncludeName }
        }

        #Exclude Names
        if ($ExcludeName -ne "") {
            $LicList = $LicList | Where-Object { $_.FriendlyName -notmatch $ExcludeName }
        }
    }

    foreach ($LIC in $LicList) {
        if (($FriendlyName) -and ($null -ne $LIC.FriendlyName)) {
            $xmlOutput += "
            <result>
            <channel>$($LIC.FriendlyName) - Free Licenses</channel>
            <value>$($LIC.PrepaidUnitsEnabled - $LIC.ConsumedUnits)</value>
            <unit>Count</unit>
            </result>

            <result>
            <channel>$($LIC.FriendlyName) - Total Licenses</channel>
            <value>$($LIC.PrepaidUnitsEnabled)</value>
            <unit>Count</unit>
            </result>
            "
        }
        else {
            $xmlOutput += "
            <result>
            <channel>$($LIC.SkuPartNumber) - Free Licenses</channel>
            <value>$($LIC.PrepaidUnitsEnabled - $LIC.ConsumedUnits)</value>
            <unit>Count</unit>
            </result>

            <result>
            <channel>$($LIC.SkuPartNumber) - Total Licenses</channel>
            <value>$($LIC.PrepaidUnitsEnabled)</value>
            <unit>Count</unit>
            </result>
            "
        }
    }
}

$xmlOutput += "</prtg>"

$xmlOutput