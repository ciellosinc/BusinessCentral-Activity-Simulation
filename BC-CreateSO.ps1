$startTimeGlobal = [System.Diagnostics.Stopwatch]::StartNew()

1..2000 | Foreach-Object -ThrottleLimit 5 -Parallel {

$AzureApplication = @();
$AzureApplication += [PSCustomObject]@{
    Name     = 'BC Create SO 01'
    TenantId = 'f891245c-c75d-4431-8f28-f888c6dc51fb';
    ClientId = '28f6b0bb-9919-441e-8763-baf74fc8484c';
    SecretKey = 'bNy***************************GddjT';
}
$AzureApplication += [PSCustomObject]@{
    Name     = 'BC Create SO 02'
    TenantId = 'f891245c-c75d-4431-8f28-f888c6dc51fb';
    ClientId = '8391a7a6-5d0f-4b99-a565-8c717da00f42';
    SecretKey = 'Ysb8****************************bwV';
}
$AzureApplication += [PSCustomObject]@{
    Name     = 'BC Create SO 03'
    TenantId = 'f891245c-c75d-4431-8f28-f888c6dc51fb';
    ClientId = 'c174b496-065c-41ae-b372-645a3c0184c1';
    SecretKey = '5rU*****************************azv';
}
$AzureApplication += [PSCustomObject]@{
    Name     = 'BC Create SO 04'
    TenantId = 'f891245c-c75d-4431-8f28-f888c6dc51fb';
    ClientId = '2e7361f3-3ae3-479e-b943-d21f97bc93ff';
    SecretKey = 'qBO*******************************3';
}
$AzureApplication += [PSCustomObject]@{
    Name     = 'BC Create SO 05'
    TenantId = 'f891245c-c75d-4431-8f28-f888c6dc51fb';
    ClientId = '7f732d39-83c0-420b-9ce1-0bf90e5dde87';
    SecretKey = 'qV******************************cCI';
}

$CompanyName = '31b55c63-7cb7-ef11-b8f6-6045bdc89e7b' #or 31b55c63-7cb7-ef11-b8f6-6045bdc89e7b or 'CRONUS USA, Inc.'
$EnvironmentName = "Production"

$ExportDataRootFolder = 'D:\Data\SO Creation\'
$null = New-Item -Path $ExportDataRootFolder -ItemType Directory -Force -ErrorAction SilentlyContinue

# https://api.businesscentral.dynamics.com/v2.0/f891245c-c75d-4431-8f28-f888c6dc51fb/Production/ODataV4/Company('CRONUS%20USA%2C%20Inc.')/G_LBudgetEntries
# https://api.businesscentral.dynamics.com/v2.0/f891245c-c75d-4431-8f28-f888c6dc51fb/Production/api/v2.0/
[uri]$url =  'https://api.businesscentral.dynamics.com/'

$BCAPIType = 'V2' # 'V2' or 'ODATA'
$isCompaniesSectionNeeded = $true
$isDeltaIncrementalExportEnabled = $true

$DataEntity = 'salesOrders'

##### FUNCTIONS - START

function Get-BCURL {
    param (
        [string] $Entity,
        [ValidateSet('V2', 'ODATA')]
        [string] $APIType = "V2",
        [string] $Tenant,
        [string] $Company,
        [string] $Environment = "Production",
        [string] $Query = "",
        [string] $EntityId = "",
        [string] $Function = "",
        [bool] $isCompaniesSectionNeeded = $true,
        [bool] $isDeltaIncrementalExportEnabled = $true
    )

    if($Expand.Length -gt 0 -and $Function.Length -gt 0)
    {
        throw Write-Error "Expand and Function options are not supported at the same time."
    }

    [System.UriBuilder] $ListRecordsURL = 'https://api.businesscentral.dynamics.com/'

    switch -Exact ($APIType) {
        'V2' {
            $urlPathPart = '/v2.0/' + $Tenant + '/' + $Environment + '/api/v2.0'
            $CompanyURLPart = 'companies(' + $Company +')'
        }
        'ODATA' {
            $urlPathPart = '/v2.0/' + $Tenant + '/' + $Environment + '/ODataV4'
            $CompanyURLPart = 'Company(' + "'" + $Company + "'" + ')'
        }
        Default {
            throw Write-Error "$APIType is not supported as BC API Type"
        }
    }

    if ($isCompaniesSectionNeeded) {
        $ListRecordsURL.Path = "$urlPathPart/$CompanyURLPart/$Entity"
    } else {
        $ListRecordsURL.Path = "$urlPathPart/$Entity"
    }

    if($EntityId.Length -gt 0)
    {
        $tmpUrl = $ListRecordsURL.Path
        $ListRecordsURL.Path = $tmpUrl + '(' + $EntityId + ')'
    }

    if($Query.Length -gt 0)
    {
        $ListRecordsURL.Query = $Query
    }

    if($Function.Length -gt 0)
    {
        $tmpUrl = $ListRecordsURL.Path
        $ListRecordsURL.Path = $tmpUrl + '/' + $Function
    }
    #Write-Host " BC URL was built: " $ListRecordsURL
    return $ListRecordsURL
}

function Get-BCAuthToken {
    param (
        [string] $Tenant,
        [string] $ClientId,
        [string] $ClientSecret,
        $url = 'https://api.businesscentral.dynamics.com/'
    )
    
    # Authorization
    Add-Type -AssemblyName System.Web
    [string]$absoluteURL = $url.Remove($url.Length-1,1)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $Body = @{
        "client_id" = $ClientId
        "client_secret" = $ClientSecret
        "grant_type" = 'client_credentials'
        "scope" = "$absoluteURL/.default"
    }   
    #Write-Host "   ..App Id:" $ClientId -ForegroundColor Yellow
    $login = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body $Body -ContentType 'application/x-www-form-urlencoded'
     
    [string]$Bearer = $login.access_token
    return $Bearer
}

function Get-BCHeaders {
    param (
        [string] $BearerToken,
        [System.UriBuilder] $URI = 'https://api.businesscentral.dynamics.com/',
        [string] $ContentType = "application/xml"
    )
    
    $headers = @{
        "Accept" = $ContentType
        "Accept-Charset" = "UTF-8"
        "Authorization" = "Bearer $BearerToken"
        "Host" = "$($URI.Host)"
        "Accept-Language" = "en-US"
        #"Data-Access-Intent" = "ReadOnly"
        "Prefer" = "odata.maxpagesize=20000"
    }
    #Write-Host "--Headers host:" $headers.Host
    return $headers
}

function Get-BCCustomerList
{
    param (
        [string] $Tenant,
        [string] $Company,
        [string] $Environment,
        [string] $ClientId,
        [string] $ClientSecret
    )
    [System.UriBuilder] $URL = Get-BCURL -Entity "customers" -Tenant $Tenant -Company $Company -Environment $Environment
    $token = Get-BCAuthToken -Tenant $Tenant -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = Get-BCHeaders -BearerToken $token -URI $URL
    $headers.Add("Data-Access-Intent","ReadOnly");
    #Write-Host "  Headers for Customers:" $headers

    [string]$RequestURL = $URL.Uri.AbsoluteUri
    $resultREST = Invoke-RestMethod -Method Get -Uri $RequestURL -Headers $headers -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3
    return $resultREST.value 
}

function Get-BCRandomCustomer
{
    param (
        [string] $Tenant,
        [string] $Company,
        [string] $Environment,
        [string] $ClientId,
        [string] $ClientSecret
    )

    return Get-BCCustomerList -Tenant $Tenant -Company $Company -ClientId $ClientId -ClientSecret $ClientSecret -Environment $Environment | Get-Random
}

function Get-BCItemList
{
    param (
        [string] $Tenant,
        [string] $Company,
        [string] $ClientId,
        [string] $Environment,
        [string] $ClientSecret
    )
    [System.UriBuilder] $URL = Get-BCURL -Entity "items" -Tenant $Tenant -Company $Company -Environment $Environment
    $token = Get-BCAuthToken -Tenant $Tenant -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = Get-BCHeaders -BearerToken $token -URI $URL
    [string]$RequestURL = $URL.Uri.AbsoluteUri
    
    $headers.Add("Data-Access-Intent","ReadOnly");
    #Write-Host "  Headers for Items:" $headers

    $resultREST = Invoke-RestMethod -Method Get -Uri $RequestURL -Headers $headers -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3

    [System.Array]$items = $resultREST.value | where {  ($_.number -notlike 'WRB-*') -and ($_.number -notlike 'SP-BOM*') -and ($_.number -notlike 'SP-SCM*') } #Remove not usable items
    return [System.Array]$items
}

function Get-BCRandomItem
{
    param (
        [string] $Tenant,
        [string] $Company,
        [string] $ClientId,
        [string] $Environment,
        [string] $ClientSecret
    )
    
    return (Get-BCItemList  -Tenant $Tenant -Company $Company -ClientId $ClientId -ClientSecret $ClientSecret -Environment $Environment | Get-Random)
}

function Invoke-BCSalesOrderCreation {
    param (
        [string] $Tenant,
        [string] $Company,
        [string] $Environment,
        [string] $ClientId, 
        [string] $ClientSecret,
        [PSCustomObject] $Customer,
        [System.Array] $ItemsList
    )

    ## Get bearer token and headers

    $BearerToken = Get-BCAuthToken -Tenant $Tenant -ClientId $ClientId -ClientSecret $ClientSecret
    $headers = Get-BCHeaders -BearerToken $BearerToken -URI 'https://api.businesscentral.dynamics.com/' -ContentType 'application/json; charset=utf-8' 

    #Write-Host "   Randomly selected customer number is: $($Customer.number)" -ForegroundColor Cyan

    $curDate = Get-Date -UFormat "%Y-%m-%d"
    $lineList = @()
    $row = 1
    $sequence = 10000
    do{
        $randomItem = $ItemsList | Get-Random
        
        #Write-Host "     Randomly selected item id is: $($randomItem.number)" -ForegroundColor Cyan
        
        $obj = @{
            sequence= $sequence
            lineType= "Item"
            lineObjectNumber= $randomItem.number
            unitOfMeasureCode= $randomItem.baseUnitOfMeasureCode
            quantity= 1
            unitPrice=$randomItem.unitPrice
        }

        if($lineList -contains $obj)
        {
            continue;
        }

        $lineList +=  $obj
        $sequence += 1
        $row++;
    }
    while($row -le $(1..10|Get-Random)) #########################

    $lineList
    # Request Body
    $body = @{
        orderDate = $curDate
        postingDate = $curDate
        customerNumber = "$($Customer.number)"
        currencyCode = "USD"
        salesOrderLines = $lineList
    } | ConvertTo-Json

    ## Generate SO url with expanded salesOrderLines
    [System.UriBuilder] $ListRecordsURL = Get-BCURL -Entity $DataEntity -Tenant $Tenant -Company $Company -Environment $Environment -Query '$expand=salesOrderLines'
    [string]$RequestURLSO = $ListRecordsURL.Uri.AbsoluteUri
    #Write-Host "Constructed URL: $RequestURLSO"
    #Write-Host "Headers:" $headers
    #Write-Host "Header host:" $headers.Host
    #Write-Host "Body:" $body

    #Write-Host "Creating SO for Customer:" $Customer.number "With items:" $lineList.lineObjectNumber -ForegroundColor Cyan
    # Create the Sales Order
    $startTimeSO = [System.Diagnostics.Stopwatch]::StartNew()
    $createdSO = Invoke-RestMethod -Uri $RequestURLSO -Method Post -Headers $headers -Body $body -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3 -RetryIntervalSec 1
    $startTimeSO.Stop();

    #Write-Host "   SalesOrder with id $($createdSO.number) is created! Duration:" $startTimeSO.ElapsedMilliseconds "ms" -ForegroundColor Cyan
    # ship SO
    ## Generate SO shipment$invoicing url
    [System.UriBuilder] $ListRecordsURL = Get-BCURL -Entity $DataEntity -Tenant $Tenant -Company $Company -Environment $Environment -EntityId $createdSO.id -Function "Microsoft.NAV.shipandinvoice"
    [string]$RequestURLShipAndInvoice = $ListRecordsURL.Uri.AbsoluteUri
    #Write-Host "Constructed URL: $RequestURLShipAndInvoice"

    $startTimeSOShipAndInvoice = [System.Diagnostics.Stopwatch]::StartNew()
    Invoke-RestMethod -Uri $RequestURLShipAndInvoice -Method  Post -Headers $headers -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3 -RetryIntervalSec 1
    $startTimeSOShipAndInvoice.Stop();

    #Write-Host "   SalesOrder with id $($createdSO.number) was shiped and invoiced! Duration:" $startTimeSOShipAndInvoice.ElapsedMilliseconds "ms" -ForegroundColor Cyan
    #Write-Host "   SalesOrder " $createdSO.number "was created, shiped and invoiced! Duration:" $startTimeSOShipAndInvoice.ElapsedMilliseconds "ms" -ForegroundColor Cyan
    # Output the response
    return $createdSO
}

##### FUNCTIONS - END

### BODY - START

$AzureApp = $AzureApplication | Get-Random

$customerList = Get-BCCustomerList -Tenant $AzureApp.TenantId -Company $CompanyName -ClientId $AzureApp.ClientId -ClientSecret $AzureApp.SecretKey -Environment $EnvironmentName
$itemList = Get-BCItemList -Tenant $AzureApp.TenantId -Company $CompanyName -ClientId $AzureApp.ClientId -ClientSecret $AzureApp.SecretKey -Environment $EnvironmentName

$startTime = [System.Diagnostics.Stopwatch]::StartNew()

$customer = $customerList | Get-Random
$createdSalesOrder = Invoke-BCSalesOrderCreation -Tenant $AzureApp.TenantId -Company $CompanyName -ClientId $AzureApp.ClientId -ClientSecret $AzureApp.SecretKey -Environment $EnvironmentName -Customer $customer -ItemsList $itemList 
#Write-Host $createdSalesOrder.id

$startTime.Stop();
#Write-Host "Total duration $($startTime.ElapsedMilliseconds) ms" -ForegroundColor Yellow

$GlobalDurationInMS = $using:startTimeGlobal
Write-Host "($($_.ToString('0000'))) Timer:" $GlobalDurationInMS "SalesOrder" $createdSalesOrder.number "for customer" $createdSalesOrder.customerNumber "was created, shipped and invoiced! Duration:" $startTime.ElapsedMilliseconds "ms" "Items:[" $createdSalesOrder.SalesOrderLines.LineObjectNumber "] App User:" $AzureApp.Name -ForegroundColor Cyan

$ExportMeasure = [PSCustomObject]@{
    No  = $_.ToString('0000');
    Timer = $GlobalDurationInMS;
    CurrentTime = $(Get-Date -Format "HH:mm:ss:fffff");
    UNIXtimeInMS = $([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds());
    SalesOrder = $createdSalesOrder.number;
    Customer = $createdSalesOrder.customerNumber;
    DurationMS = $startTime.ElapsedMilliseconds;
    ItemsCnt = $createdSalesOrder.SalesOrderLines.Count;
    AppUser = $AzureApp.Name
}

$ExportMeasure | Export-Csv -Path $($ExportDataRootFolder+'\measureSOcreation.csv') -Append

### BODY - END

} #For each in parallel - END

$startTimeGlobal.Stop();
