<#
    .SYNOPSIS
    Monitors Microsoft 365 App Secret expiration

    .DESCRIPTION
    Using MS Graph this Script shows the Microsoft 365 App Secret expiration

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

    .PARAMETER Exclude/Include
    Regular expression to exclude secrets on DisplayName or AppName

    Example: ^(PRTG-APP)$

    Example2: ^(PRTG-.*|TestApp123)$ excludes PRTG-* and TestApp123

    #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced

    "PRTG-M365-AppSecrets.ps1" -ApplicationID 'Test-APPID' -TenantID 'contoso.onmicrosoft.com' -AccessSecret 'Test-AppSecret'

    Microsoft 365 Permission:
        1. Open Azure AD
        2. Register new APP
        3. Overview >> Get Application ID
        4. Set API Permissions >> MS Graph >> Application >> Application.Read.All
        5. Certificates & secrets >> new Secret

    Author:  Jannos-443
    https://github.com/Jannos-443/PRTG-M365
#>
param(
    [string] $TenantID = '',
    [string] $ApplicationID = '',
    [string] $AccessSecret = '',
    [string] $IncludeSecretName = '',
    [string] $ExcludeSecretName = '',
    [string] $IncludeAppName = '',
    [string] $ExcludeAppName = ''
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

        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Method POST -Body $Body
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

$Result = GraphCall -URL " https://graph.microsoft.com/v1.0/applications"

$NextExpiration = 2000

$SecretList = New-Object System.Collections.ArrayList

foreach ($SingleResult in $Result) {
    foreach ($passwordCredential in $SingleResult.passwordCredentials) {
        [datetime]$ExpireTime = $passwordCredential.endDateTime
        $object = [PSCustomObject]@{
            AppDisplayname    = $SingleResult.displayName
            SecretDisplayname = $passwordCredential.displayName
            Enddatetime       = $ExpireTime
            DaysLeft          = ($ExpireTime - (Get-Date)).days
        }
        $null = $SecretList.Add($object)
    }

    foreach ($keyCredential in $SingleResult.keyCredentials) {
        [datetime]$ExpireTime = $keyCredential.endDateTime
        $object = [PSCustomObject]@{
            AppDisplayname    = $SingleResult.displayName
            SecretDisplayname = $keyCredential.displayName
            Enddatetime       = $ExpireTime
            DaysLeft          = ($ExpireTime - (Get-Date)).days
        }
        $null = $SecretList.Add($object)
    }
}

#Region Filter
#APP
if ($ExcludeAppName -ne "") {
    $SecretList = $SecretList | Where-Object { $_.AppDisplayname -notmatch $ExcludeAppName }
}

if ($IncludeAppName -ne "") {
    $SecretList = $SecretList | Where-Object { $_.AppDisplayname -match $IncludeAppName }
}
#SECRET
if ($ExcludeSecretName -ne "") {
    $SecretList = $SecretList | Where-Object { $_.SecretDisplayname -notmatch $ExcludeSecretName }
}

if ($IncludeSecretName -ne "") {
    $SecretList = $SecretList | Where-Object { $_.SecretDisplayname -match $IncludeSecretName }
}
#End Region Filter

$ListCount = ($SecretList | Measure-Object).Count
if ($ListCount -eq 0) {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>No Secrets or Certs found! Check Permissions</text>"
    Write-Output "</prtg>"
    Exit
}

$SecretList = $SecretList | Sort-Object Enddatetime

$Top5 = $SecretList | Select-Object -First 5
$OutputText = "Next to expire: "

foreach ($Top in $Top5) {
    $OutputText += "App `"$($Top.AppDisplayname)`" Secret `"$($Top.SecretDisplayname)`" expires in $($Top.DaysLeft)d; "
}

#Next Expiration
$NextExpiration = ($SecretList | Select-Object -First 1).DaysLeft

$xmlOutput += "<result>
<channel>Next Cert Expiration</channel>
<value>$($NextExpiration)</value>
<unit>Custom</unit>
<CustomUnit>Days</CustomUnit>
<LimitMode>1</LimitMode>
<LimitMinWarning>30</LimitMinWarning>
<LimitMinError>10</LimitMinError>
</result>"

$Less90Days = ($SecretList | Where-Object { $_.DaysLeft -le 90 } | Measure-Object).count
$Less180Days = ($SecretList | Where-Object { $_.DaysLeft -le 180 } | Measure-Object).count

$xmlOutput += "<result>
<channel>less than 90 days left</channel>
<value>$($Less90Days)</value>
<unit>Count</unit>
</result>
<result>
<channel>less than 180 days left</channel>
<value>$($Less180Days)</value>
<unit>Count</unit>
</result>"

$OutputText = $OutputText.Replace("<", "")
$OutputText = $OutputText.Replace(">", "")
$OutputText = $OutputText.Replace("#", "")
$xmlOutput = $xmlOutput + "<text>$($OutputText)</text>"

$xmlOutput = $xmlOutput + "</prtg>"

$xmlOutput