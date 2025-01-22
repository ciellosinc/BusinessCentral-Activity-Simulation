$tenantDomain = 'f891245c-c75d-4431-8f28-f888c6dc51fb' 
$ApplicationClientId = '0ccea774-8049-423e-86a2-d8f302477acb' 
$ApplicationClientSecretKey = 'KK0*******************************************b8M'
$CompanyName = '31b55c63-7cb7-ef11-b8f6-6045bdc89e7b' #or 31b55c63-7cb7-ef11-b8f6-6045bdc89e7b or 'CRONUS USA, Inc.'

$ExportDataRootFolder = 'D:\Data\'
$null = New-Item -Path $ExportDataRootFolder -ItemType Directory -Force -ErrorAction SilentlyContinue

#Almost infinittive loop to repeat the Data Export every 5 minutes
$total = 300
foreach ($k in 1..$total) {
    $progPercent = (($k/$total) * 100)
    Write-Progress -Activity "Data Extraction" -Status "Iteration No $k" -PercentComplete $progPercent


# Loop through all Data Entities
$DataEntities = @('companies','customers','vendors','items','unitsOfMeasure','salesInvoices','salesShipments','salesOrders','purchaseInvoices','purchaseOrders','purchaseReceipts','itemLedgerEntries','generalLedgerEntries','accounts');
foreach ($DataEntity in $DataEntities) {
    Write-Host "   Working on" $DataEntity -ForegroundColor Magenta


# https://api.businesscentral.dynamics.com/v2.0/f891245c-c75d-4431-8f28-f888c6dc51fb/Production/ODataV4/Company('CRONUS%20USA%2C%20Inc.')/G_LBudgetEntries
# https://api.businesscentral.dynamics.com/v2.0/f891245c-c75d-4431-8f28-f888c6dc51fb/Production/api/v2.0/
[uri]$url =  'https://api.businesscentral.dynamics.com/' 

$BCAPIType = 'V2' # 'V2' or 'ODATA'
$isCompaniesSectionNeeded = $true
$isDeltaIncrementalExportEnabled = $true

#$DataEntity = 'Power_BI_Customer_List'
#$DataEntity = 'Power_BI_Sales_List'
#$DataEntity = 'salesInvoices'
#$DataEntity = 'salesShipments'
#$DataEntity = 'salesOrders'

#$DataEntity = 'purchaseInvoices'
#$DataEntity = 'purchaseOrders'
#$DataEntity = 'purchaseReceipts'

#$DataEntity = 'itemLedgerEntries'
#$DataEntity = 'generalLedgerEntries'
#$DataEntity = 'accounts'

#$DataEntity = 'customers'
#$DataEntity = 'vendors'
#$DataEntity = 'items'
#$DataEntity = 'unitsOfMeasure'
#$DataEntity = 'companies'; $isCompaniesSectionNeeded = $false; $isDeltaIncrementalExportEnabled = $false

switch ($DataEntity) {
    'salesInvoices'     { $DataEntityForLines = 'salesInvoiceLines'; $isLinesMustBeExtractedToo = $true}
    'salesShipments'    { $DataEntityForLines = 'salesShipmentLines'; $isLinesMustBeExtractedToo = $true}
    'salesOrders'       { $DataEntityForLines = 'salesOrderLines'; $isLinesMustBeExtractedToo = $true}
    'purchaseInvoices'  { $DataEntityForLines = 'purchaseInvoiceLines'; $isLinesMustBeExtractedToo = $true}
    'purchaseOrders'    { $DataEntityForLines = 'purchaseOrderLines'; $isLinesMustBeExtractedToo = $true}
    'purchaseReceipts'  { $DataEntityForLines = 'purchaseReceiptLines'; $isLinesMustBeExtractedToo = $true}
    'companies'         { $DataEntityForLines = ''; $isCompaniesSectionNeeded = $false; $isDeltaIncrementalExportEnabled = $false}
    Default             { $DataEntityForLines = ''; $isLinesMustBeExtractedToo = $false }
}

$measurementArray = @();

[System.UriBuilder] $ListRecordsURL = $url
$startTime = [System.Diagnostics.Stopwatch]::StartNew()

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
    "Data-Access-Intent" = "ReadOnly"
    "Prefer" = "odata.maxpagesize=20000"
}
 
#    "Prefer" = "odata.include-annotations=""OData.Community.Display.V1.FormattedValue"""

#$ListRecordsURL.Query = '$top=10'

[string]$lastDateTime = '';
[string]$currentDateTime = Get-Date -Format "yyyy'-'MM'-'dd'T'HH':'mm':'ssZ"
#[string]$MeasureDataExtractionFileName = Join-Path -Path $ExportDataRootFolder -ChildPath $('!!measurements_'+$DataEntity+$currentDateTime+'.csv').Replace(':','');
[string]$MeasureDataExtractionFileName = Join-Path -Path $ExportDataRootFolder -ChildPath $('!!measurements_.csv')

if ($isDeltaIncrementalExportEnabled)
{
    $LastDateTimeFileName = Join-Path -Path $ExportDataRootFolder -ChildPath $('!lastDateTime_' + $DataEntity + '.txt')
    If (Test-Path -Path $LastDateTimeFileName)
    {
        # Read the last datetime execution from the file
        $lastDateTime = Get-Content -Path $LastDateTimeFileName 
    } else {
        # if file doesn't exist then geenrate Zero date time
        $lastDateTime = '1999-01-01T00:00:00Z'
    }
    $ListRecordsURL.Query = '$filter=lastModifiedDateTime gt ' + $lastDateTime
}


#Write-Host "   URL" $ListRecordsURL -ForegroundColor Yellow
#Write-Host "   Headers" -ForegroundColor Yellow
#$headers
#Write-Host "   Trying to get data..." -ForegroundColor Yellow
 
 
$resultREST=$null
[string]$RequestURL = $ListRecordsURL.Uri.AbsoluteUri
[Int64]$i = 0
#Write-Host "   Trying to get data from $RequestURL" -ForegroundColor Yellow

Do {
    Write-Host "   URL" $RequestURL -ForegroundColor Yellow
    $startTimeDataEntity = [System.Diagnostics.Stopwatch]::StartNew()
    $resultREST = Invoke-RestMethod -Method Get -Uri $RequestURL -Headers $headers -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3 -RetryIntervalSec 3
    

    $measurementArray += [PSCustomObject]@{
        Entity = $DataEntity;
        DateTimeStamp = $currentDateTime;
        Count = $resultREST.value.Count;
        Kind = 'HTTP Only';
        Duration = $startTimeDataEntity.ElapsedMilliseconds.ToString();
    }
 
    Write-Host "Results [$i]..." -ForegroundColor Yellow
    #$resultREST.value | ConvertTo-Json
    $resultREST.value[0] | ConvertTo-Json

    # Export results if any to flat file
    if ($resultREST.value.Length -gt 0)
    {
        #Export to file
        $tempCSVFileName = $DataEntity +'_' + $lastDateTime + '_' + $i.ToString('000000')      + '.csv'
        $CSVFileName = Join-Path -Path $ExportDataRootFolder -ChildPath $tempCSVFileName.Replace(':','')
        $resultREST.value | Export-Csv -Path $CSVFileName -Force

        $measurementArray += [PSCustomObject]@{
            Entity = $DataEntity;
            DateTimeStamp = $currentDateTime;
            Count = $resultREST.value.Count;
            Kind = 'HTTP and CSV';
            Duration = $startTimeDataEntity.ElapsedMilliseconds.ToString();
        }
    

        # Export Lines per each document
        if ($isLinesMustBeExtractedToo -and ($DataEntityForLines.Length -gt 1))
        {
            #$sync = [System.Collections.Hashtable]::Synchronized($origin) # https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/write-progress-across-multiple-threads?view=powershell-7.4
            $job = $resultREST.value | Foreach-Object <#-AsJob#> -ThrottleLimit 10 -Parallel {
                #Action that will run in Parallel. Reference the current object via $PSItem and bring in outside variables with $USING:varname

                $document = $_
                $DataEntity = $using:DataEntity
                $DataEntityForLines = $using:DataEntityForLines
                $RequestURL = $using:RequestURL
                $lastDateTime = $using:lastDateTime
                $i = $using:i
                $ExportDataRootFolder = $USING:ExportDataRootFolder
                $measurementLinesArray = @()

                $startTimeLine = [System.Diagnostics.Stopwatch]::StartNew()

                #Replace Data Entity to the DataEntityLines and remove$filter on lastModifiedDateTime field. We are going to extract all lines.
                [string]$tempDocLinesText = $DataEntity + '('+ $document.Id +')' +'/' + $DataEntityForLines
                $position = $RequestURL.IndexOf($DataEntity);
                [string]$RequestURLines = $RequestURL.Substring(0, $position) + $tempDocLinesText
                
                #Extract all document lines
                $resultRESTLines = Invoke-RestMethod -Method Get -Uri $RequestURLines -Headers $using:headers -ContentType 'application/json; charset=utf-8' -MaximumRetryCount 3 -RetryIntervalSec 3
                #$resultRESTLines.value[0]

                $measurementLinesArray += [PSCustomObject]@{
                    Entity = $DataEntityForLines;
                    DateTimeStamp = $using:currentDateTime;
                    Count = $resultRESTLines.value.Count;
                    Kind = 'HTTP Only';
                    Duration = $startTimeLine.ElapsedMilliseconds.ToString();
                }

                # Save extracted lines to a flat file
                $tempCSVFileNameLines = $DataEntity +'_' + $lastDateTime + '_' + $i.ToString('000000') + '_'+ $DataEntityForLines +'_' +$document.Id   + '.csv'
                $CSVFileNameLines = Join-Path -Path $ExportDataRootFolder -ChildPath $tempCSVFileNameLines.Replace(':','')
                if ($resultRESTLines.value.Length -ge 1)
                {   #Export when it is anything to export
                    $resultRESTLines.value | Export-Csv -Path $CSVFileNameLines -Force
                }

                $startTimeLine.Stop();
                Write-Host 'Extract Lines per' $DataEntity ':' $document.Id '(' $document.number ')' ' ..duration'  $($startTimeLine.ElapsedMilliseconds) 'ms' ' Lines extracted:' $resultRESTLines.value.Count

                $measurementLinesArray += [PSCustomObject]@{
                    Entity = $DataEntityForLines;
                    DateTimeStamp = $using:currentDateTime;
                    Count = $resultRESTLines.value.Count;
                    Kind = 'HTTP and CSV';
                    Duration = $startTimeLine.ElapsedMilliseconds.ToString();
                }
            }

            #while ($job.State -eq 'Running') {
            #    Start-Sleep -Milliseconds 100
            #    $job.Progress
            #}
            #Write-Progress -Activity 'Extract Lines per Document' -Completed
        }
    }

    $RequestURL = $resultREST.'@odata.nextLink'
    $i = $i + 1;

    $startTimeDataEntity.Stop();

    $measurementArray += [PSCustomObject]@{
        Entity = $DataEntity;
        DateTimeStamp = $currentDateTime;
        Count = $resultREST.value.Count;
        Kind = 'Full extract';
        Duration = $startTimeDataEntity.ElapsedMilliseconds.ToString();
    }

} until ($null -eq $resultREST.'@odata.nextLink')

if ($isDeltaIncrementalExportEnabled)
{
    #If export was successfull then save the current datetime for the next execution 
    $currentDateTime | Out-File -FilePath $LastDateTimeFileName
}

$startTime.Stop();
Write-Host "Total duration $($startTime.ElapsedMilliseconds) ms" -ForegroundColor Yellow

$measurementArray | Export-Csv -Append -Path $MeasureDataExtractionFileName -Force

Start-Sleep -Milliseconds 500 #Add delay between each Data entity

} #For each DataEntity loop end


Write-Host "Waiting for 5 minutes" -ForegroundColor Magenta
Start-Sleep -Seconds 60 # 5 minutes * 60 seconds
} #End of Almost infinittive loop to repeat the Data Export every 5 minutes
Write-Progress -Activity "Data Extraction" -Completed
