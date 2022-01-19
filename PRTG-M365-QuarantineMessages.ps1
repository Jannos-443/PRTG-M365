param(
    [string] $CertificateFilePath = '',
    [string] $CertificateThumbPrint = '',
    [string] $CertificatePassword = '',
    [string] $ApplicationID = '',
    [string] $TenatDomainName = '',
    [int]$DaysAgo = 30                      #Days in the past to search for Quarantine Messages
)

#Catch all unhandled Errors
trap{
    if($connected)
        {
        try
            {
            $null = Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
        catch
            {
            }
        }
    $Output = "line:$($_.InvocationInfo.ScriptLineNumber.ToString()) char:$($_.InvocationInfo.OffsetInLine.ToString()) --- message: $($_.Exception.Message.ToString()) --- line: $($_.InvocationInfo.Line.ToString()) "
    $Output = $Output.Replace("<","")
    $Output = $Output.Replace(">","")
    $Output = $Output.Replace("#","")
    try
        {
        $Output = $Output.Substring(0,2000)
        }
    catch
        {
        }
    Write-Output "<prtg>"
    Write-Output "<error>1</error>"
    Write-Output "<text>$Output</text>"
    Write-Output "</prtg>"
    Exit
}
<#
#https://stackoverflow.com/questions/19055924/how-to-launch-64-bit-powershell-from-32-bit-cmd-exe
#############################################################################
#If Powershell is running the 32-bit version on a 64-bit machine, we 
#need to force powershell to run in 64-bit mode .
#############################################################################
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") 
    {
    if ($myInvocation.Line) 
        {
        [string]$output = &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
        }
    else
        {
        [string]$output = &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" $args
        }

    #Remove any text after </prtg>
    try{
        $output = $output.Substring(0,$output.LastIndexOf("</prtg>")+7)
        }

    catch
        {
        }

    Write-Output $output
    exit
    }

#############################################################################
#End
#############################################################################
#>
if(($TenatDomainName -notlike "*.onmicrosoft.com") -or ($null -eq $TenatDomainName))
    {
    Throw "TenantDomainName Variable does not end *.onmicrosoft.com"
    }

if(($ApplicationID -eq "") -or ($Null -eq $ApplicationID))
    {
    Throw "ApplicationID Variable is empty"
    }

#Check Authentication Requirements
if("" -eq $CertificateThumbPrint)
    {
    if(("" -eq $CertificatePassword) -or ("" -eq $CertificateFilePath))
        {
        Throw "`$CertificateFilePath+`$CertificatePassword or `$CertificateThumbPrint is needed"
        }
    }
else 
    {
    if(("" -ne $CertificatePassword) -or ("" -ne $CertificateFilePath))
        {
        Throw "Use ether `$CertificateFilePath or `$CertificateThumbPrint. Not both"
        }
    else
        {
        Test-Path (-not $CertificateFilePath)
            {
            Throw "Cert not found under Path $($CertificateFilePath)"
            }
        }
    }

# Error if there's anything going on
$ErrorActionPreference = "Stop"


# Import ExchangeOnlineManagement module
try {
    Import-Module "ExchangeOnlineManagement" -ErrorAction Stop
    #Import-Module "C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.3\ExchangeOnlineManagement.psm1" -ErrorAction Stop     
} catch {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error Loading ExchangeOnlineManagement Powershell Module ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
}

# Disconnect hanging Session if available
$connected = $false
try
    {
    $null = Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
catch
    {
    }

# Connect to EXO
try {
    if($CertificateThumbPrint)
        {
        $null = Connect-ExchangeOnline -CertificateThumbPrint $CertificateThumbPrint -AppId $ApplicationID -Organization $TenatDomainName
        }
    
    if($CertificateFilePath)
        {
        $null = Connect-ExchangeOnline -CertificateFilePath $CertificateFilePath -CertificatePassword (ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force) -AppId $ApplicationID -Organization $TenatDomainName
        }
    $connected = $true
    Write-Output "connected successfully"
    Start-Sleep -Seconds 1
    }
 
catch
    {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Could not connect to EXO. Error: $($_.Exception.Message)</text>"
    Write-Output "</prtg>"
    Exit
    }

# Get Quarantine Messages last 30 Days
$QuaraMails = Get-QuarantineMessage -StartReceivedDate (Get-Date).AddDays(-$daysago)

# Count
$QuaraMailsCount = 0
$QuaraMailsText = ""

# hardcoded list that applies to all hosts
$IgnoreScript = '^(TestServer123:E:\\)$' 
# \ has to be Escaped with another \

ForEach ($QuaraMail in $QuaraMails)
    {
    $QuaraMailsCount ++
    $dateoutput = (Get-Date -Date $QuaraMail.ReceivedTime -Format "dd.MM.yy-HH:mm").ToString()
    $QuaraMailsText += "Time: $($dateoutput) Direction: $($QuaraMail.Direction) Sender: $($QuaraMail.SenderAddress) Type: $($QuaraMail.Type)###"
    }
$QuaraMailsText = $QuaraMailsText.Insert(0,"QuarantineMessages: $($QuaraMailsCount) - ")

try
    {
    $QuaraMailsText = $QuaraMailsText.Substring(0,2000)
    }
catch
    {
    }

# Disconnect from EXO
Disconnect-ExchangeOnline -Confirm:$false

$connected = $false

# Results
$xmlOutput = '<prtg>'
if ($QuaraMailsCount -gt 0)
    {
    $xmlOutput = $xmlOutput + "<text>$($QuaraMailsText)</text>"
    }


$xmlOutput = $xmlOutput + "<result>
        <channel>QuaraMailsCount</channel>
        <value>$QuaraMailsCount</value>
        <unit>Count</unit>
        <limitmode>1</limitmode>
        <LimitMaxError>0</LimitMaxError>
        </result>"   
        


$xmlOutput = $xmlOutput + "</prtg>"

#region: Output

Write-Output $xmlOutput

#endregion