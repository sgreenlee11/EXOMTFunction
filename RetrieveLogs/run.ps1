<#
    
    This is a sample Azure Storage Queue Trigger function to retrieve content 
    blobs from the Office 365 Management and Activity API
    The sample scripts are not supported under any Microsoft standard support 
    program or service. The sample scripts are provided AS IS without warranty  
    of any kind. Microsoft further disclaims all implied warranties including,  
    without limitation, any implied warranties of merchantability or of fitness for 
    a particular purpose. The entire risk arising out of the use or performance of  
    the sample scripts and documentation remains with you. In no event shall 
    Microsoft, its authors, or anyone else involved in the creation, production, or 
    delivery of the scripts be liable for any damages whatsoever (including, 
    without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use 
    of or inability to use the sample scripts or documentation, even if Microsoft 
    has been advised of the possibility of such damages.
#>
# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

#Specify Username with EXO Message trace permissions
$username = "admin@M365x002534.onmicrosoft.com"

#Retrieve Account password from Credential Vault
# Our Key Vault Credential that we want to retreive URI - Update with customer
$vaultSecretURI = "https://sgo365key.vault.azure.net/secrets/ExoPassword/4cff04fe564b4a668a83262210b2156a"
$vaultSecretURI = $vaultSecretURI + "?api-version=7.0"

#Values for local token service

$apiVersion = "2017-09-01"
$resourceURI = "https://vault.azure.net"
$tokenAuthURI = $env:MSI_ENDPOINT + "?resource=$resourceURI&api-version=$apiVersion"
$tokenResponse = Invoke-RestMethod -Method Get -Headers @{"Secret" = "$env:MSI_SECRET" } -Uri $tokenAuthURI

# Use Key Vault AuthN Token to create Request Header
$requestHeader = @{ Authorization = "Bearer $($tokenresponse.access_token)" }
# Call the Vault and Retrieve Creds
$password = Invoke-RestMethod -Method GET -Uri $vaultSecretURI -ContentType 'application/json' -Headers $requestHeader

#Create Credential Object
$securepassword = ConvertTo-SecureString $password.value -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($username, $securepassword)

#Connect to EXO PowerShell - Not working as the module does not yet support PS Core.
#Import-Module ExchangeOnlineManagement
#Connect-ExchangeOnline -Credential $cred

$session = @{ }
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber

$index = 1
If ((Test-Path .\RetrieveLogs\StartTime.xml) -eq $true) {
    $StartDate = Import-Clixml .\RetrieveLogs\StartTime.xml
}
else {
    $StartDate = (Get-Date).AddHours(-1)
}
$EndDate = (Get-Date).addminutes(-30).ToString()


[int]$msg_count = 0
Do {
    $messageTrace = Get-MessageTrace -PageSize 5000 -StartDate $StartDate -EndDate $EndDate -Page $index #| Select MessageTraceID,Received,*Address,*IP,Subject,Status,Size,MessageID  #| Sort-Object Received
    $index ++
    $messageTrace.count
    $msg_count = $msg_count + $messageTrace.count
  
    if ($messageTrace.count -gt 0) {
        $Up_date = $MesageTrace | Select-Object Received -Last 1
        Push-OutputBinding -Name outputEventHubMessage -Value $MessageTrace
    }


} 
while ($messageTrace.count -gt 0)

if ($up_date -ne $null) {
    $up_date | Export-Clixml .\RetrieveLogs\StartTime.xml
}
else {
    (Get-Date).AddHours(-1) | Export-Clixml .\RetrieveLogs\StartTime.xml
}
