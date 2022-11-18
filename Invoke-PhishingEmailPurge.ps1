## Name: Invoke-PhishingEmailPurge
## Purpose: Purging phishing emails from office 365 mailbox's and blocking them in mimecast
## By: CHris Tejeda
## Date: 5.6.2019
## GitHub Profile: https://github.com/ctejeda

Function Invoke-PhishingEmailPureg {

    	 param 
    (
        [Parameter(Mandatory=$true)]
        [string]$PhishingEmail,
        [Parameter(Mandatory=$true)]
        [string]$Logfile
    )


Function Connectto-365 {

$365AdminUser = Read-Host "Enter Your Office 365 Admin Email"
$pass = Read-Host "Enter 365 Password" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass)
$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$UserCredential = New-Object System.Management.Automation.PsCredential("$365AdminUser", (ConvertTo-SecureString "$password" -AsPlainText -Force))
$SessionOption = New-PSSessionOption -IdleTimeout 120000 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -SessionOption $SessionOption
Import-PSSession $Session -DisableNameChecking -AllowClobber 


}

Function Invoke-Mimecast {

 param 
    (
        [Parameter(Mandatory=$true)]
        [string]$emailAddress
        
    )

#Setup required variables
$baseUrl = "https://us-api.mimecast.com"
$uri = "/api/directory/add-group-member"	
$url = $baseUrl + $uri
$accessKey = "Accesskeyhere"
$secretKey = "seckeyhere"
$appId = "appIDhere"
$appKey = "APpkeyHere"


## AccessKey
#$SecureAccessKey = Get-Content "path\to\accesskeyhere\Mimecast_AccessKey.txt" | ConvertTo-SecureString 
#$BSTR_AccessKey = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureAccessKey)
#$accessKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR_AccessKey)
#$global:accessKey = $accessKey

## Secret Key
#$Secure_Sec_Key = Get-Content "path\to\accesskeyhere\Mimecast_SecKEY.txt" | ConvertTo-SecureString 
#$BSTR_Sec_Key = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure_Sec_Key)
#$secretKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR_Sec_Key)
#$global:secretKey = $secretKey


## App IP 


 
#Generate request header values
$hdrDate = (Get-Date).ToUniversalTime().ToString("ddd, dd MMM yyyy HH:mm:ss UTC")
$requestId = [guid]::NewGuid().guid
 
#Create the HMAC SHA1 of the Base64 decoded secret key for the Authorization header
$sha = New-Object System.Security.Cryptography.HMACSHA1
$sha.key = [Convert]::FromBase64String($secretKey)
$sig = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($hdrDate + ":" + $requestId + ":" + $uri + ":" + $appKey))
$sig = [Convert]::ToBase64String($sig)
 
#Create Headers
$headers = @{"Authorization" = "MC " + $accessKey + ":" + $sig;
                "x-mc-date" = $hdrDate;
                "x-mc-app-id" = $appId;
                "x-mc-req-id" = $requestId;
                "Content-Type" = "application/json"}
 


#Create post body
$postBody = "{
                    ""data"": [
                        {
                            ""id"": ""IDhere"",
                            ""emailAddress"": ""$emailAddress""
                         
                        }
                    ]
                }"



#Send Request
$response = Invoke-RestMethod -Method Post -Headers $headers -Body $postBody -Uri $url 
#Print the response
$fail = $response | select -ExpandProperty fail | select -ExpandProperty key
$response




}

Function Invoke-Logger {
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$true)]
        [string]$LogFile,
        [Parameter(Mandatory=$false)]
        [Switch]$ShowOutput,
        [Parameter(Mandatory=$false)]
        [Switch]$SaveToCSV,
        [Parameter(Mandatory=$false)]
        [Switch]$ShowError,
        [Parameter(Mandatory=$false)]
        [Switch]$ShowWarning
    )
        $date = Get-Date -UFormat "%m/%d/%Y %H:%M:%S"
        Add-Content $LogFile -Value "$date - $Message"
        if ($ShowOutput)
        {$ShowOutput = Write-Host $Message -ForegroundColor Green }
        if ($ShowError)
        {$ShowError = Write-Host $Message -ForegroundColor Red }
        if ($ShowWarning)
        {$ShowWarning = Write-Host $Message -ForegroundColor Yellow }
        if ($SaveToCSV){$array = @(); $array += [pscustomobject] @{"computer" = $computer; "Message" = "$Message"; "Date" = "$date" }; $array | Export-Csv -Path "$env:USERPROFILE\$LogFile.csv" -NoTypeInformation }
        $ErrorActionPreference='stop'
        
}
Invoke-Logger -Message "Connecting to Office 365" -ShowOutput -LogFile $Logfile
Connectto-365
try {


Invoke-Logger -Message "Starting the Purging proccess for phising email $PhishingEmail" -LogFile $logfile -ShowOutput

try {
Invoke-Logger -Message "Adding email $PhishingEmail to the Blcoked Senders list on Mimecast" -LogFile $logfile -ShowOutput
$PhishingEmail.ToString()
Invoke-Mimecast -emailAddress $PhishingEmail 
}

catch {
Invoke-Logger "The following error occured while trying to add $PhishingEmail to Blocked Senders group in Mimecast: $_" -LogFile $Logfile -ShowError

}
Get-MessageTrace -SenderAddress "$PhishingEmail" | foreach {

$user = $_.RecipientAddress 

$PhishingEmail.ToString()
$SearchQuery = @"
from:"$PhishingEmail"
"@

Invoke-Logger -Message "Purging $user mailbox with the following search email Query: $SearchQuery" -LogFile $logfile -ShowOutput
Search-Mailbox -Identity $user -SearchQuery ($($SearchQuery)) -DeleteContent -Force -Verbose


}

} 

catch {Invoke-Logger -Message "The Following error has occured: $_" -LogFile $logfile}

}
