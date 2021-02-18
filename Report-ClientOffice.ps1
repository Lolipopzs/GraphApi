using namespace Microsoft.Identity.Client
    
    function Get-AuthTokenMSAL
    {
    <#
    .SYNOPSIS
     Gets an OAuth token for use with the Microsoft Graph API
     .DESCRIPTION
     Gets an OAuth token for use with the Microsoft Graph API using MSAL
    .EXAMPLE
     Get-AuthToken -TenantName "contoso.onmicrosoft.com" -clientId "74f0e6c8-0a8e-4a9c-9e0e-4c8223013eb9" -redirecturi "urn:ietf:wg:oauth:2.0:oob" -resourceAppIdURI "https://graph.microsoft.com"
    .PARAMETER TentantName
    Tenant name in the format <tenantname>.onmicrosoft.com
    .PARAMETER clientID
    The clientID or AppID of the native app created in AzureAD to grant access to the reporting API
    .Parameter redirecturi
    The replyURL of the native app created in AzureAD to grant access to the reporting API
    .Parameter resourceAppIDURI
    protocol and hostname for the endpoint you are accessing. For the Graph API enter "https://graph.microsoft.com" 
    #>
 
            # For testing purpose ... not clean --> add autodetection of msal
            $MSAL = "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Identity.Client.4.17.1\lib\net461\Microsoft.Identity.Client.dll"
           
            Try
                {
                    [System.Reflection.Assembly]::LoadFrom($MSAL) | Out-Null
                   
                }
            Catch
                {
                    
                    Write-Warning "Unable to load MSAL assemblies."
                    Throw $error[0]
                }
           
           #Build the logon URL with the tenant name
           $authority = "https://login.microsoftonline.com/$TenantName"
           $TenantId = (Invoke-WebRequest https://login.microsoftonline.com/$TenantName/v2.0/.well-known/openid-configuration | ConvertFrom-Json).token_endpoint.split('/')[3]
           Write-Verbose "Logon Authority: $authority"
           
           # to adapt to the usecase
                      
                
           $redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
                          
           # build the app using MSAL
           $pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ApplicationID).WithTenantId($TenantId).WithRedirectUri($redirectUri)
           $AuthResult = $pcaConfig.Build().AcquireTokenByUsernamePassword($Scopes,$Ucreds.UserName, $Ucreds.Password)
           $AuthResultExec = $AuthResult.ExecuteAsync()
           Start-Sleep -Seconds 3
           if ($AuthResultExec.Status -eq "RanToCompletion"){
                 $token =  $AuthResultExec.Result
                 $AccessToken = $token.AccessToken
                 write-host "Authent succesfull" -ForegroundColor Green
                 write-host "Your access token is : " $AccessToken
           }
           Else {
               Write-host "An auth error occured :"
               write-host $AuthResultExec.Exception
               $AccessToken = "ERROR"  
           }
           return $token 
    }


$TenantName = "contoso.onmicrosoft.com"

$CLID = "client_id"

$scopeList = @("https://graph.microsoft.com/Reports.Read.All")
$Scopes = New-Object System.Collections.Generic.List[string]

ForEach ($scope in $scopeList){ 
$Scopes.Add($Scope)
}
	

$folder = "E:\Scripts\Scheduling"
$keypath = "$folder\aes.key"
$UserName = Get-Content "$folder\MyU.O365" | ConvertTo-SecureString -Key (Get-Content $keypath)
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($UserName)
[string]$UserNameString = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
$pass = Get-Content "$folder\MyP.O365" | ConvertTo-SecureString -Key (Get-Content $keypath)
$creds = New-Object System.Management.Automation.PsCredential($UserNameString,$pass)
$Username = $creds.UserName
$Password = $creds.Password	
	
$Credential = new-object -typeName System.Management.Automation.PSCredential -ArgumentList $userName, $Password  

$Tok = Get-AuthTokenMSAL -Ucreds $Credential -ApplicationID $CLID -resourceAppIdURI "https://graph.microsoft.com" -TenantName $TenantName -Scopes $scopes
$authHeader = @{
    'Authorization'= $Tok.AccessToken  
    'Content-Type'= 'application/json'  
 }

$outfile = ".\ClientOfficeVersion.csv"

if (Test-Path -Path $outfile -ErrorAction SilentlyContinue)
{Remove-Item $outfile}

$tenantId="tenant_id" #replace with your tenant ID

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Data = Invoke-RestMethod -Headers $authHeader -Uri "https://graph.microsoft.com/v1.0/reports/getEmailAppUsageVersionsUserCounts(period='D7')" -Method Get
