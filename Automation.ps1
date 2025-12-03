<#Author: Kyler B.
Date: July 7, 2025
Program was orginially developed in Windows PowerShell ISE console.
This automation program takes an Excel file from Dell and processes asset 
information for import into the IT asset management system.#>
#Requires -Version 5.1
#Requires -Modules ImportExcel

Import-Module ImportExcel

#Make sure to get rid of any credentials stored in the variable
#Clear any confidential variables

#Timestamp
$timestamp = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"


# Define a dynamic file path
$OutputSuccess = "AssetImport_Success_Log_$($timestamp).xlsx"
$OutputFail = "AssetImport_Error_Log_$($timestamp).xlsx"
$Path_Success = "C:\Users\KB0192\Downloads\Testing\$OutputSuccess"
$Path_Fail = "C:\Users\KB0192\Downloads\Testing\$OutputFail"

#Process has begun
Write-Host "Asset Automation process has begun."

#########################Authentication
Write-Host "Setting up API connection and Credentials..." -ForegroundColor Yellow

################ITA API Setup
#ITA API Details
$itmsApiUrl = "https://milliken-amc-stg.ivanticloud.com/api/odata/businessobject/cis"
$incidentApiUrl = "https://milliken-amc-stg.ivanticloud.com/api/odata/businessobject/incidents"
$journalApiUrl = "https://milliken-amc-stg.ivanticloud.com/api/odata/businessobject/journal__notess"
$apiKey = "4E57B85DE161465C8FAD7F0E9BF307CA"

$ITA_Headers = @{
    "Authorization" = "rest_api_key=$apiKey"
    "content-type"  = "application/json"
}

################SharePoint Setup
$clientId = "#####################################"
$tenantId = "#####################################"
$clientSecret = "#####################################"
$scope = "https://graph.microsoft.com/.default"

$body_sp = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = $scope
}

$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body_sp
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }



#####################Helper Functions
<#Gets the subtype for asset#>
function Get-Country {
    param ($country)

    if ($country -eq "RMC") {return "US"}
    return "Europe"
   #Placeholders for now.
    
}

<#Gets the subtype for asset#>
function Get-SubType {
    param ($model)

    if ($model -eq "Latitude 7450") {return "Standard Laptop"}
    return "Standard Laptop"
   #return "Mobile Device" Saying Mobile Device is not an acceptable field.Not in the validation list.
    
}

<#Get the acquistion for the asset#>
function Get-Acquisition {
    param ($Inputvalue)
    if($Inputvalue -eq "US"){ return "Leased"}
    #For Non-US
    return "Purchased" 
}

<#Function that calculates the start of the Lease#>
function Get-Lease {
    param ($lease)

    <#Some calculation will be performed#>

}

<#Function that calculates the start of the warranty.#>
function Get-Warranty {
    param ($warrant)
 
    <#Some calculation#>

}

<#Function for determining the location of the asset#>
function Get-Location {
    param ($inputValue)

    if ($inputValue -eq "US") {return "RMC Roger Milliken Center"}

    return "RMC Roger Milliken Center" #THIS IS A PLACEHOLDER USED FOR TESTING
    #Has to be a location that is validated by Ivanti
    
}

<#Function for determining storage space#>
function Get-Storage {
    param ($inputVal)

    if ($inputVal -eq "US") {return "RMC Intake" }
    return "RMC Intake" #Has to be a validated input by Ivanti
    
}

<#Function for getting missing fields#>
function Get-MissingFields {
    param ($asset, $requiredFields)
    $missing = @()
    foreach ($field in $requiredFields) {
        if (-not $asset.$field) {
            $missing += $field
        }
    }
    return $missing
}

<#Function that gets the RECID for the asset if it
already exists#>
function Get-RecID{
    param($serialNumber)
    Write-Host "Retrieving RecID for Serial Number: $serialNumber"

     $url = "https://milliken-amc-stg.ivanticloud.com/api/odata/businessobject/cis?$" + "filter" + "=SerialNumber eq '" + $serialNumber + "'"

    try{

        $response = Invoke-RestMethod -Uri $url -Method "GET" -Headers $ITA_Headers
        if($response.value.Count -gt 0){
            return $response.value[0].RecID
        }
    }catch{
        Write-Host "Error retrieving RecId for asset: $serialNumber"
    }
    return $null
}

<#Adds journal entry for each asset 
 whether the being newly added or updated.#>
function Add-JournalEntry {
    
    param ([string]$RecordID)
    try {
        $JournalData = @{
            ParentLink_Category = "CI"
            ParentLink_RecID = "$RecordID"
            Subject = "Asset Automation Program"
            Source = "Other"
            Category = "Memo"
            NotesBody = "Asset was added by Kyler's Asset Automation Program. From file $($file.name)"
        }
        
        Write-Host "Creating Journal Entry..." -ForegroundColor Yellow 

        $JournalJson = $JournalData | ConvertTo-Json -Depth 7

        Write-Host "Journal Entry added successfully" -ForegroundColor Green
        
        $journalResponse = Invoke-RestMethod -Uri $journalapiurl -Method POST  -Headers $ITA_Headers -Body $JournalJson 
        return $journalResponse

    
    } catch {
        Write-Error "Failed to insert journal entry: $_" 
    }
}

#####################Core Logic Functions
<#Defines function ITAFormat that will take the 
asset and map data from the email to ITA format#>
function ITAFormat {
    param ($asset)

    $formatted = @{
        SerialNumber           = $asset.'Dell Service Tag' #May need to be changed based upon the region
        Status                 = "En route"
        CIType                 = "Computer"
        ivnt_AssetSubtype      = Get-SubType -model $asset.Model
        Model                  = $asset.Model
        ivnt_Location          = Get-Location -country $asset.'Ship Zip' 
        ivnt_AcquisitionMethod = Get-Acquisition -Inputvalue $asset.'Ship Country'
        Name                   = $asset.'Dell Service Tag'
    }

    $requiredFields = @("SerialNumber","Model","Status","ivnt_Location","ivnt_AcquisitionMethod","Name")
    $missingFields = Get-MissingFields -asset $formatted -requiredFields $requiredFields

    return @{
        Formatted     = $formatted
        MissingFields = $missingFields
    }
}

<#Process of adding or updating assets
to ITA. If a complete duplicate of the asset will not be processed.#>
function AddOrUpdateAssetToITA {
    param ($formattedAsset)

    $recID = Get-RecID -serialNumber $formattedAsset['SerialNumber']
    if ($recID) {
        $getUrl = "$itmsApiUrl('$recID')"
        try {
            <# Sends a GET request to retrieve the current asset 
            data from the specified URL using the provided headers.#>

            $currentAsset = Invoke-RestMethod -Uri $getUrl -Method "GET" -Headers $ITA_Headers
            
            # Initializes a flag to track whether an update is needed.
            $needsUpdate = $false
            
            # Initializes an empty hashtable to store fields that need to be updated.
            $updatePayload = @{}

            
            # Iterates through each key in the $formattedAsset hashtable.
            foreach ($key in $formattedAsset.Keys) {

                # Retrieves the new value for the current key
                $newValue = $formattedAsset[$key]

                # Skips the current iteration if the new value is null, empty, or whitespace.
                if ([string]::IsNullOrWhiteSpace($newValue)) {
                    continue
                }

                
                # Compares the current asset's value with the new value.
                <# If they differ, adds the key and new value to the update payload and 
                sets the update flag.#>

                if ($currentAsset.$key -ne $newValue) {
                    $updatePayload[$key] = $newValue
                    $needsUpdate = $true
                }
            }

            # If no updates are needed, logs a message and exits the function early.
            if (-not $needsUpdate) {
                Write-Host "No update needed for asset: $($formattedAsset['SerialNumber'])"
                return $true
            }
            
            # Constructs the PATCH URL using the record ID.
            $patchUrl = "$itmsApiUrl('$recID')"

            
            # Sends a PATCH request to update the asset with the new values in the payload.
            # Converts the payload to JSON format with a depth of 5 to handle nested structures.
            Invoke-RestMethod -Uri $patchUrl -Method "PATCH" -Headers $ITA_Headers -Body ($updatePayload | ConvertTo-Json -Depth 5)
            
            # Logs a message indicating the asset was updated.
            Write-Host "Asset updated: $($formattedAsset['SerialNumber'])" -ForegroundColor Green  
            
            #Insert journal entry after update
            Add-JournalEntry -RecordID $recID

            return $true

        } catch {
            Write-Host "Error updating asset: $($formattedAsset['SerialNumber'])"
            return $false
        }

    } else {
        try {
            $cleanedAsset = @{}
            foreach ($key in $formattedAsset.Keys) {
                if (-not [string]::IsNullOrWhiteSpace($formattedAsset[$key])) {
                    $cleanedAsset[$key] = $formattedAsset[$key]
                }
            }

            Invoke-RestMethod -Uri $itmsApiUrl -Method "POST" -Headers $ITA_Headers `
                -Body ($cleanedAsset | ConvertTo-Json -Depth 7)
            Write-Host "Asset added: $($formattedAsset['SerialNumber'])"

            #Retrieve new RecID after adding assert for it to be updated
            $newRecID = Get-RecID -serialNumber $formattedAsset['SerialNumber']
            if($newRecID){
                Add-JournalEntry -RecordID $newRecID
            }
            return $true
        } catch {
            Write-Host "Error adding asset: $($formattedAsset['SerialNumber'])"
            return $false
        }
    }
}

#####################Main Logic Functions
Write-Host "Beginning processing of each asset..."

function ProcessAssets{
    param($assetsArray)

    
    # Process and collect formatted assets goes through an array
    foreach ($asset in $assetsArray) {
        Write-Host "----------------------------------------"
        $result = ITAFormat -asset $asset
        $formatted = $result.Formatted
        $missingFields = $result.MissingFields

        #Missing fields check 
        if ($missingFields.Count -gt 0) {
            $global:missingFieldAssets += [PSCustomObject]@{
                SerialNumber  = $asset.'Dell Service Tag'
                MissingFields = ($missingFields -join ", ")
            }

            # Also add to failed assets for logging
            $global:failedAssets += [PSCustomObject]@{
                SerialNumber  = $asset.'Dell Service Tag'
                Status        = "Missing Required Fields"
                MissingFields = ($missingFields -join ", ")
            }

            continue
        }


        $response = AddOrUpdateAssetToITA -formattedAsset $formatted


        if ($response) {
            $global:successAssets += [PSCustomObject]$formatted
        } else {
            $global:failedAssets += [PSCustomObject]$formatted
        }
    }
}

####################SharePoint


###Site ID
try{
    Write-Host "Attempting to obtain site ID..." -ForegroundColor Yellow
    # Step: Get Site ID
    $siteUrl = "https://graph.microsoft.com/v1.0/sites/milliken.sharepoint.com:/sites/MobileInfrastructureandSupportTeam/"
    $siteInfo = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $siteUrl -Method Get
    $siteId = $siteInfo.id
    Write-Host "Site ID Found: $($siteId)" -ForegroundColor Green

}catch{
    
    Write-Host "Error trying to get site ID: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails) {Write-Host "Details: $($_.ErrorDetails.Message)" -ForegroundColor DarkRed}

}

###Drive ID 
try{
    Write-Host "Attempting to obtain drive ID..." -ForegroundColor Yellow
    $drives = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
    $driveId = $drives.value[1].id # Adjust if you know the specific library name
    Write-Host "Drive ID Found: $($driveId)" -ForegroundColor Green

    $headers = @{Authorization = "Bearer $accessToken"}
}catch{
    Write-Host "Error obtaining drive ID: $_"
    Write-Host "Drive ID: $driveId" -ForegroundColor Cyan

}

# === Step: Get Drive Items
$URI = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/root/children"
$drive_items = Invoke-RestMethod -Headers $headers -Uri $URI

$URI = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$NeedsProcessing/children"
$files = Invoke-RestMethod -Headers $headers -Uri $URI
Write-Host ""


##########SharePoint File Processing

Write-Host "Connecting to SharePoint..."
$siteUrl = "https://graph.microsoft.com/v1.0/sites/milliken.sharepoint.com:/sites/MobileInfrastructureandSupportTeam/"
$siteInfo = Invoke-RestMethod -Headers $headers -Uri $siteUrl
$siteId = $siteInfo.id

$drives = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
$driveId = $drives.value[1].id

$NeedsProcessing = "01V36QITTLC5X2JG7RDJCLGBLHF6TKJ4WQ"
$files = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$NeedsProcessing/children"

#Storage variables for assets.
$global:successAssets = @()
$global:failedAssets = @()
$global:missingFieldAssets = @()

$processedFiles = @()

foreach ($file in $files.value) {
    Write-Host "Found File: $($file.name)" -ForegroundColor Green
    Write-Host "File ID: $($file.id)" -ForegroundColor Yellow
    
    $processedFiles += $($file.name -join ', ')

    $file_id = $file.id

    
    # Get all worksheets in the file
    $worksheetsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$file_id/workbook/worksheets"
    $worksheets = Invoke-RestMethod -Headers $headers -Uri $worksheetsUri

    # Check if any worksheets were returned
    if (-not $worksheets.value) {
        Write-Warning "No worksheets found in $($file.name)"
        continue
    }

    # Use the first worksheet's name
    $sheetName = $worksheets.value[0].name

    # Get the used range from that worksheet
    $rangeUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$file_id/workbook/worksheets('$sheetName')/usedRange(valuesOnly=true)"
    $range = Invoke-RestMethod -Headers $headers -Uri $rangeUri


    if (-not $range.values) {
        Write-Warning "No data found in $($file.name)"
        continue
    }

    #Convert to Structured Array
    #Strips HTML tags in headers if they exist
    $headersRow = $range.values[0] | ForEach-Object { ($_ -replace '<.*?>', '').Trim() }
    $dataRows = $range.values[1..($range.values.Count - 1)]

    $assetsArray = @()
    foreach ($row in $dataRows) {
        $asset = @{}
        for ($i = 0; $i -lt $headersRow.Count; $i++) {
            $asset[$headersRow[$i]] = $row[$i]
        }
        $assetsArray += [PSCustomObject]$asset
    }

    #Process This File's Data
    Write-Host "`n--- Data from $($file.name) ---" -ForegroundColor Cyan
    $assetsArray | Format-Table -AutoSize

   
    #Calling process function to begin processing
    ProcessAssets -assetsArray $assetsArray
}

############## After processing all files/Got Excel Files
Write-Host "`n--- Exporting logs to Excel ---" -ForegroundColor Cyan

# FolderID for Processed Folder
$ProcessedAssets = "01V36QITWKE23L4VLEYVEYACKH7HFC4NXM"

#ID for Log Folders
$ErrorLog = "01V36QITT273DFDMPCSZAKAFYWW7EJBJO4"
$SuccessLog = "01V36QITWXJNJQ23UA2RH23T4WWG64HP2I"


#Export logs to Excel

$global:successAssets | Export-Excel -Path $Path_Success -WorksheetName "Success" -AutoSize

#If there are any failed assets then export the file
if ($global:failedAssets.Count -gt 0) {
    $global:failedAssets | Export-Excel -Path $Path_Fail -WorksheetName "Failed" -AutoSize
}

#Append processed file names to the Success Excel file
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($Path_Success)
$worksheet = $workbook.Sheets.Item(1)

$row = $worksheet.UsedRange.Rows.Count + 2
$worksheet.Cells.Item($row, 1).Value2 = "Files Processed:"
$row++

foreach ($fileName in $processedFiles) {
    $worksheet.Cells.Item($row, 1).Value2 = $fileName
    $row++
}

$workbook.Save()
$workbook.Close($false)

<#A potential extension of the below might be to 
only append the files where there was an error adding/updating that 
asset.
#>

#Append processed file names to the Failed Excel file (if it exists)
if ($global:failedAssets.Count -gt 0) {
    $workbook = $excel.Workbooks.Open($Path_Fail)
    $worksheet = $workbook.Sheets.Item(1)

    $row = $worksheet.UsedRange.Rows.Count + 2
    $worksheet.Cells.Item($row, 1).Value2 = "Files Processed:"
    $row++

    foreach ($fileName in $processedFiles) {
        $worksheet.Cells.Item($row, 1).Value2 = $fileName
        $row++
    }

    $workbook.Save()
    $workbook.Close($false)
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null


# Upload logs to SharePoint
Write-Host "`n--- Uploading logs to SharePoint ---" -ForegroundColor Cyan

#Credentials for uploading the logs to SharePoint
$uploadHeaders = @{
    Authorization = "Bearer $accessToken"
    "Content-Type" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

####Components that make the Success Log Upload
$successFileContent = [System.IO.File]::ReadAllBytes($Path_Success)

# Upload to the same folder or a subfolder
$uploadUriSuccess = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/${ProcessedAssets}:/SuccessLogs/${OutputSuccess}:/content"

try {
    Invoke-RestMethod -Uri $uploadUriSuccess -Headers $uploadHeaders -Method PUT -Body $successFileContent
    Write-Host "Uploaded success log to SharePoint: $OutputSuccess" -ForegroundColor Green
} catch {
    Write-Host "Failed to upload success log: $($_.Exception.Message)" -ForegroundColor Red
}

####Components for the Error Log(But Conditioanl)
#Only uploads the Error Log if failed assets exist and there is an actual file path to the log
if($global:failedAssets.Count -gt 0 -and (Test-Path $Path_Fail)) {
    $failFileContent = [System.IO.File]::ReadAllBytes($Path_Fail)

    $uploadUriFail = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/${ProcessedAssets}:/ErrorLogs/${OutputFail}:/content"
     try {
        Invoke-RestMethod -Uri $uploadUriFail -Headers $uploadHeaders -Method PUT -Body $failFileContent
        Write-Host "Uploaded error log to SharePoint: $OutputFail" -ForegroundColor Green
    } catch {
        Write-Host "Failed to upload error log: $($_.Exception.Message)" -ForegroundColor Red
    }
}
#########################Email Notification Setup

$Sender = "ivanti.in@milliken.com"
$MailServer = "mlknsmtp.milliken.com"
$Recipient = "Kyler.Bailey@Milliken.com"
$CC = "Kyler.Bailey@Milliken.com"
$Subject = "Asset Automation Program Completed Summary"
$Subject_Incident = "Incident Ticket Has Been Created"
$Encoding = [System.Text.Encoding]::UTF8

#Serial number summaries
$serialNumber = ($global:successAssets | ForEach-Object { $_.SerialNumber }) -join "`r`n"
$serialNumber_Failed = ($global:failedAssets | ForEach-Object { $_.SerialNumber }) -join "`r`n"

#Email body
$Body = @"
The asset automation process has been completed.
These files were processed: $($processedFiles)

***************************************
Successfully Imported Assets:
File: $outputPathSuccess
Serial Numbers:
$serialNumber

***************************************
Failed Assets:
File: $outputPathFail
Serial Numbers (If Serial Number Missing that Asset had no Serial Number):
$serialNumber_Failed

If any assets failed, an incident ticket has been generated.
Best regards,
Kyler's Automation Program
"@


#########################Send Summary Email

Write-Host "Sending summary email..." -ForegroundColor Yellow
Send-MailMessage -SmtpServer $MailServer -From $Sender -To $Recipient -Cc $CC -Subject $Subject -Body $Body -Encoding $Encoding
Write-Host "Summary email sent." -ForegroundColor Green


#########################Incident Ticket Email (If Needed)

if ($global:failedAssets.Count -gt 0) {
    $Body_Incident = @"
An incident ticket has been generated for the assets that failed to upload.
Please review the attached error log or ticket system for more details.
Best regards,
Kyler's Automation Program
"@

    Send-MailMessage -SmtpServer $MailServer -From $Sender -To $Recipient -Cc $CC -Subject $Subject_Incident -Body $Body_Incident -Encoding $Encoding
    Write-Host "Incident email sent." -ForegroundColor Green
}

#####################Incident Ticket Created

#Incident Description for the body of the ticker
$Incident_Description = @"
Error Log of Assets`nFileName for Error Log: $($file.name)`nThe following assets were missing required fields:`n$missingFieldDetails

"@

#Incident Ticket Body (Json formatted)
$incident = @{

CreatedBy= "Kyler.Bailey@Milliken.com"
Subject= "Asset Automation Error Log"
Symptom= "$($Incident_Description)" 
Impact= "Only I have Issue"
Urgency= "Not Urgent"
Priority= "3"
Source= "Phone"
Status= "Logged"
OwnerTeam= "Client Support"
ProfileFullName= "Kyler Bailey"
ProfileLink_Category= "Employee"
ProfileLink_RecID= "7060D3F88A6743ABAB792A7792CD7967"
ProfileLink = "7060D3F88A6743ABAB792A7792CD7967"
}

$incidentJson = $incident | ConvertTo-Json -Depth 7

$Body_Incident = @"

An Incident Ticket has been generate for the assets that had missing required fields during 
upload.

Pleasae review the ticket for details.

Best regards.

"@


Write-Host "MissingFieldAssets count: $($global:missingFieldAssets.Count)" -ForegroundColor Cyan

# Incident Ticket and Email Notification for Missing Fields
if ($global:missingFieldAssets.Count -gt 0) {
    $missingFieldDetails = ""
    foreach ($entry in $global:missingFieldAssets) {
        $missingFieldDetails += "SerialNumber: $($entry.SerialNumber) - Missing Fields: $($entry.MissingFields)`n"
    }
    try{
        Write-Host "Missing Field(s) Data being Sent:" -ForegroundColor Yellow
     
        #Incident Ticket is sent to ITA

        $Variable = Invoke-RestMethod -Uri $incidentApiUrl -Method Post -Headers $ITA_Headers -Body $incidentJson
        Write-Host "Incident Ticket has been sent." -ForegroundColor Green
        
        #Incident Email is sent
        Send-MailMessage -SmtpServer $MailServer -From $Sender -To $Recipient -Cc $CC -Subject $Subject_Incident -Body $Body_Incident -Encoding $Encoding
        Write-Host "Email has been sent." -ForegroundColor Green
    
    }catch{

        Write-Host "Error occurred here are the details:$_" -ForegroundColor Red
    
    }
}
#########################Final Log
Write-Host "Export complete." -ForegroundColor Green
Write-Host "Successfully processed assets saved to SuccessLogs via SharePoint" -ForegroundColor Magenta

#Only outputing an error path if there are any failed assets
if($global:failedAssets.Count -gt 0){
    Write-Host "Failed assets saved to ErrorLogs via SharePoint" -ForegroundColor Magenta
}


Write-Host "Kyler's Asset Automation Process Complete." -ForegroundColor Green