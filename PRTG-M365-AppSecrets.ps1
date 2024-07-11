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

    .PARAMETER ProxyAddress
    Provide a proxy server address if this required to make connections to M365
    Example: http://proxy.example.com:3128

    .PARAMETER ProxyUser
    Provide a proxy authentication user if ProxyAddress is used

    .PARAMETER ProxyPassword
    Provide a proxy authentication password if ProxyAddress is used

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
    [string] $ExcludeAppName = '',
    [string] $ProxyAddress = '',
    [string] $ProxyUser = '',
    [string] $ProxyPassword = ''
)

# Remove ProxyAddress var if it only contains an empty string or else the Invoke-RestMethod will fail if no proxy address has been provided
if ($ProxyAddress -eq "") {
    Remove-Variable ProxyAddress -ErrorAction SilentlyContinue
}

if (($ProxyAddress -ne "") -and ($ProxyUser -ne "") -and ($ProxyPassword -ne "")) {
    try {
        $SecProxyPassword = ConvertTo-SecureString $ProxyPassword -AsPlainText -Force
        $ProxyCreds = New-Object System.Management.Automation.PSCredential ($ProxyUser, $SecProxyPassword)
    }
    catch {
        Write-Output "<prtg>"
        Write-Output " <error>1</error>"
        Write-Output " <text>Error Parsing Proxy Credentials ($($_.Exception.Message))</text>"
        Write-Output "</prtg>"
        Exit
    }
}
else {
    Remove-Variable ProxyCreds -ErrorAction SilentlyContinue
}

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
        } else {
            Write-Host "Token found and still valid"
        }
    } else {
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

        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Method POST -Body $Body -Proxy $ProxyAddress -ProxyCredential $ProxyCreds
        $token = $ConnectGraph.access_token
        $tokenexpire = (Get-Date).AddSeconds($ConnectGraph.expires_in)

        Write-Host "successfully got new MS Graph Token"
    }
}
catch {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error getting MS Graph Token ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
}

$xmlOutput = '{
    "prtg": {'


#Function Graph API Call
Function GraphCall($URL) {
    #MS Graph Request
    try {
        $Headers = @{Authorization = "$($ConnectGraph.token_type) $($ConnectGraph.access_token)" }
        $GraphUrl = $URL
        $Result_Part = Invoke-RestMethod -Headers $Headers -Uri $GraphUrl -Method Get -Proxy $ProxyAddress -ProxyCredential $ProxyCreds
        $Result = $Result_Part.value
        while ($Result_Part.'@odata.nextLink') {
            $graphURL = $Result_Part.'@odata.nextLink'
            $Result_Part = Invoke-RestMethod -Headers $Headers -Uri $graphURL -Method Get -Proxy $ProxyAddress -ProxyCredential $ProxyCreds
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

$Result = GraphCall -URL "https://graph.microsoft.com/v1.0/applications"

$NextExpiration = 2000

$SecretList = New-Object System.Collections.ArrayList

# Added handling of app secrets that will return either a displayname or a customKeyIdentifier. If one is set the other one is null. As the customKeyIdentifier is a base64 string it will be encoded as UTF8.
foreach ($SingleResult in $Result) {
    foreach ($passwordCredential in $SingleResult.passwordCredentials) {
        [datetime]$ExpireTime = $passwordCredential.endDateTime
        if ($passwordCredential.displayName -ne $null) {
            $SecretDisplayName = $passwordCredential.displayName
        } else {
            $SecretDisplayName = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($passwordCredential.customKeyIdentifier))
        }
        $object = [PSCustomObject]@{
            AppDisplayname    = $SingleResult.displayName
            SecretDisplayname = $SecretDisplayName
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

#Also monitor SAML Signing certs
$Result2 = GraphCall -URL "https://graph.microsoft.com/v1.0/serviceprincipals"

foreach ($SingleResult in $Result2) {
    if ($SingleResult.signInAudience -eq "AzureADMyOrg") {
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

# Ignore secrets with the value "CWAP_AuthSecret". This is created by default with Azure AD app proxy and working as designed. It rotates keys and needs the last 3 passwords even if expired. https://learn.microsoft.com/en-us/entra/identity/app-proxy/application-proxy-faq
$SecretList = $SecretList | Where-Object {$_.SecretDisplayname -ne "CWAP_AuthSecret"}

# Ignore secrets with empty value ""
$SecretList = $SecretList | Where-Object {$_.SecretDisplayname -ne $null}

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

$xmlOutput += "
    `"result`": [
        {
            `"channel`": `"Next Cert Expiration`",
            `"value`": `"$($NextExpiration)`",
            `"unit`": `"custom`",
            `"customunit`": `"Days`",
            `"LimitMode`": `"1`",
            `"LimitMinWarning`": `"30`",
            `"LimitMinError`": `"10`"  
        },"

$Less90Days = ($SecretList | Where-Object { $_.DaysLeft -le 90 } | Measure-Object).count
$Less180Days = ($SecretList | Where-Object { $_.DaysLeft -le 180 } | Measure-Object).count

$xmlOutput += "
        {
            `"channel`": `"less than 90 days left`",
            `"value`": `"$($Less90Days)`",
            `"unit`": `"Count`"  
        },
        {
            `"channel`": `"less than 180 days left`",
            `"value`": `"$($Less180Days)`",
            `"unit`": `"Count`"  
        }"

$OutputText = $OutputText.Replace("<", "")
$OutputText = $OutputText.Replace(">", "")
$OutputText = $OutputText.Replace("#", "")
$OutputText = $OutputText.Replace("`'", "")
$xmlOutput += "

    ],
    `"text`": `"$($OutputText.Replace('`"',"`'"))`"
    }
}"

$xmlOutput
