$tenantDomain = 'f891245c-c75d-4431-8f28-f888c6dc51fb' 
$ApplicationClientId = '0ccea774-8049-0000-0000-d8f302477acb' 
$ApplicationClientSecretKey = 'KK0*******************************************Zlb8M'
$CompanyName = '31b55c63-7cb7-ef11-b8f6-6045bdc89e7b' #or 31b55c63-7cb7-ef11-b8f6-6045bdc89e7b or 'CRONUS USA, Inc.'

[uri]$url =  'https://api.businesscentral.dynamics.com/' 

$BCAPIType = 'V2' # 'V2' or 'ODATA'
$isCompaniesSectionNeeded = $true

#$DataEntity = 'companies'; $isCompaniesSectionNeeded = $false
#$DataEntity = 'Power_BI_Customer_List'
#$DataEntity = 'Power_BI_Sales_List'
#$DataEntity = 'salesInvoices'
#$DataEntity = 'salesShipments'
$DataEntity = 'salesOrders'


[System.UriBuilder] $ListRecordsURL = $url

switch -Exact ($BCAPIType) {
    'V2' { 
        #Write-Host "   Will use API V2" -ForegroundColor Yellow
        $urlPathPart = '/v2.0/'+ $tenantDomain +'/Production/api/v2.0'  #for Business Central API V2    
        $CompanyURLPart = 'companies(' + $CompanyName +')' 

        if($isCompaniesSectionNeeded) {
            $ListRecordsURL.Path = "$urlPathPart/$CompanyURLPart/$DataEntity"
        } else {
            $ListRecordsURL.Path = "$urlPathPart/$DataEntity"
        }
        #$ListRecordsURL.Query = '$top=10'
    }
    'ODATA' { 
        #Write-Host "   Will use OData API" -ForegroundColor Yellow
        $urlPathPart = '/v2.0/'+ $tenantDomain +'/Production/ODataV4'  #for Business Central OData API    '/ODataV4/Company('mycompany')/salesDocumentLines'
         
        $CompanyURLPart = 'Company(' +"'"+ $CompanyName +"'"+ ')' 
        if($isCompaniesSectionNeeded) {
            $ListRecordsURL.Path = "$urlPathPart/$CompanyURLPart/$DataEntity"
        } else {
            $ListRecordsURL.Path = "$urlPathPart/$DataEntity"
        }
        #$ListRecordsURL.Query = '$top=10'
            }
    Default { 
        throw Write-Error "$_ is not supported as BC API Type"; 
    }
}

 
#Write-Host "Authorization..." -ForegroundColor Yellow
Add-Type -AssemblyName System.Web
[string]$absoluteURL = $url.AbsoluteUri.Remove($url.AbsoluteUri.Length-1,1)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Body = @{
    "client_id" = $ApplicationClientId
    "client_secret" = $ApplicationClientSecretKey
    "grant_type" = 'client_credentials'
    "scope" = "$absoluteURL/.default"
}
 
#Write-Host "   URL" $absoluteURL -ForegroundColor Yellow
#Write-Host "   Body" -ForegroundColor Yellow
#$Body
Write-Host "   Get Authorization token..." -ForegroundColor Yellow
 
$login = $null
$login = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantDomain/oauth2/v2.0/token" -Body $Body -ContentType 'application/x-www-form-urlencoded' #-Verbose
 
$Bearer = $null
[string]$Bearer = $login.access_token
 
Write-Host "   Getting data..." -ForegroundColor Yellow
 
$headers = @{
    "Accept" = "application/xml"
    "Accept-Charset" = "UTF-8"
    "Authorization" = "Bearer $Bearer"
    "Host" = "$($url.Host)"
    "Accept-Language" = "en-US"
    #"Data-Access-Intent" = "ReadOnly"
    "Prefer" = "odata.maxpagesize=20000"
}
 
$resultREST=$null
[string]$RequestURL = $ListRecordsURL.Uri.AbsoluteUri

Write-Host "   URL" $RequestURL -ForegroundColor Yellow
$resultREST = Invoke-RestMethod -Method Get -Uri $RequestURL -Headers $headers -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3 -RetryIntervalSec 3 -Verbose

$resultREST.value | Select -Last 10 | Format-Table
