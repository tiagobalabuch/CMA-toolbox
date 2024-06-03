<#
.SYNOPSIS
    Creates an CSV files for SQL Server assessment at scale.
.DESCRIPTION
    This PowerShell script creates CSV files for SQL Server assessment at scale.
.PARAMETER FolderPath
    Specifies the path to the target directory.
.PARAMETER JsonFile
    Specifies the JSON file that contains the assessment.
.EXAMPLE
    PS> ./ReportCMA.ps1 -FolderPath "C:\temp\CMA-Report\" -JsonFile "SQLAssessment.json"
    ...
.NOTES
    Author: Tiago Balabuch
    Date: 04/06/2024
    Version: 1.0
    GitHub: https://github.com/tiagobalabuch/CMA_toolbox
    Copyright: (c) 2024 by Tiago Balabuch, licensed under MIT
    License: MIT https://opensource.org/licenses/MIT
#>
param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the folder path.")]
    [string]$FolderPath,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the JSON file path.")]
    [string]$JsonFile
)

function Handle-Error {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    # Extract error details
    $ErrorMessage = $ErrorRecord.Exception.Message
    $ErrorSource = $ErrorRecord.InvocationInfo.MyCommand
    $ErrorLine = $ErrorRecord.InvocationInfo.ScriptLineNumber
    $ErrorPosition = $ErrorRecord.InvocationInfo.OffsetInLine
    $ErrorStackTrace = $ErrorRecord.Exception.StackTrace

    # Format the error message
    $ErrorDetails = @"
Date:        $(Get-Date)
Error:       $ErrorMessage
Source:      $ErrorSource
Line:        $ErrorLine
Position:    $ErrorPosition
Stack Trace: $ErrorStackTrace
"@

    # Log the error details to a file (or take any other appropriate action)
    $ErrorLogFileName = "ErrorLogReportCMALogs.log"
    $ErrorLogFilePath = [IO.Path]::Combine($FolderPath, $ErrorLogFileName)

    Add-Content -Path $ErrorLogFilePath -Value $ErrorDetails

    # Display the error details to the user
    Write-Error ($errorDetails)
}

function Add-CsvFileServerAssessments {
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host ("=======================================================================================") -ForegroundColor White
        Write-Host ("Creating CSV file and writing data for 'ServerAssessments'") -ForegroundColor White
       
        # Create a list to hold the data
        $data = @()

        $Status = $JsonData.Status
        foreach ($Server in $JsonData.Servers) {
            foreach ($Assessment in $Server.ServerAssessments) {
                $row = [PSCustomObject]@{
                    Status                           = $Status
                    Timestamp                        = $Assessment.Timestamp.ToString()
                    ServerName                       = $Assessment.ServerName
                    FeatureId                        = $Assessment.FeatureId
                    IssueCategory                    = $Assessment.IssueCategory
                    MoreInformation                  = $Assessment.MoreInformation
                    Description                      = $Assessment.RuleMetadata.Description
                    Id                               = $Assessment.RuleMetadata.Id
                    HelpLink                         = $Assessment.RuleMetadata.HelpLink
                    Level                            = $Assessment.RuleMetadata.Level.ToString()
                    Message                          = $Assessment.RuleMetadata.Message
                    RuleScope                        = $Assessment.RuleScope
                    AppliesToMigrationTargetPlatform = $Assessment.AppliesToMigrationTargetPlatform
                }
                $data += $row
            }
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "ServerAssessment"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating the CSV file 'ServerAssessments'..") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
}

function Add-CsvFileServerProperties {
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host ("=======================================================================================") -ForegroundColor White
        Write-Host ("Creating CSV file and writing data for 'ServerProperties'") -ForegroundColor White

        $data = @()

        foreach ($Server in $JsonData.Servers) {
            $Properties = $Server.Properties
            $row = [PSCustomObject]@{
                "Server Name"         = $Properties.ServerName
                "FQDN"                = $Properties.FQDN
                "Server Version"      = $Properties.ServerVersion
                "Server Edition"      = $Properties.ServerEdition
                "Hosting Platform"    = $Properties.ServerHostPlatform
                "Server Level"        = $Properties.ServerLevel
                "Core Count"          = $Properties.ServerCoreCount.ToString()
                "Collation"           = $Properties.ServerCollation
                "Hyperthread Ratio"   = $Properties.HyperthreadRatio.ToString()
                "Logical CPU"         = $Properties.LogicalCpuCount.ToString()
                "Physical CPU"        = $Properties.PhysicalCpuCount.ToString()
                "Max Memory In Use"   = $Properties.MaxServerMemoryInUse.ToString()
                "Number of Databases" = $Properties.NumberOfUserDatabases.ToString()
                "Total Database Size" = $Properties.SumOfUserDatabasesSize.ToString()
            }
            $data += $row
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "ServerProperties"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White

    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating the CSV file 'ServerProperties'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath

    }
  
}


function Add-CsvFileServerImpactedObjects {
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Creating CSV file and writing data for 'ServerImpactedObjects'" -ForegroundColor White

        # Create a list to hold the data
        $data = @()

        foreach ($Server in $JsonData.Servers) {
            foreach ($ServerAssessment in $Server.ServerAssessments) {
                $Col_ServerName = $ServerAssessment.ServerName
                $Col_FeatureID = $ServerAssessment.FeatureId
                $Col_AppliesToTargetPlatform = $ServerAssessment.AppliesToMigrationTargetPlatform

                foreach ($ImpactedObject in $ServerAssessment.ImpactedObjects) {
                    $row = [PSCustomObject]@{
                        "Server Name"                = $Col_ServerName
                        "Feature ID"                 = $Col_FeatureID
                        "Object Name"                = $ImpactedObject.Name
                        "Object Type"                = $ImpactedObject.ObjectType
                        "Details"                    = $ImpactedObject.ImpactDetail
                        "Database Object Type"       = $ImpactedObject.DatabaseObjectType
                        "Applies To Target Platform" = $Col_AppliesToTargetPlatform
                    }
                    $data += $row
                }
            }
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "ServerImpactedObjects"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating to the CSV file 'ServerImpactedObjects'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }

}
function Add-CsvFileServerTargetReadinesses { 
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Creating CSV file and writing data for 'ServerTargetReadinesses'" -ForegroundColor White

        # Create a list to hold the data
        $data = @()

        foreach ($Server in $JsonData.Servers) {
            $Col_ServerName = $Server.Properties.ServerName

            $row = [PSCustomObject]@{
                "Server Name"                          = $Col_ServerName
                "Applies To Target Platform"           = $Server.TargetReadinesses.AzureSqlDatabase.AppliesToMigrationTargetPlatform
                "Databases List Ready For Migration"   = ($Server.TargetReadinesses.AzureSqlDatabase.DatabasesListReadyForMigration -join ',')
                "Num Of Databases Ready For Migration" = $Server.TargetReadinesses.AzureSqlDatabase.NumberOfDatabasesReadyForMigration.ToString()
                "Total Number Of Databases"            = $Server.TargetReadinesses.AzureSqlDatabase.TotalNumberOfDatabases.ToString()
                "Recommendation Status"                = $Server.TargetReadinesses.AzureSqlDatabase.RecommendationStatus
            }
            $data += $row

            $row = [PSCustomObject]@{
                "Server Name"                          = $Col_ServerName
                "Applies To Target Platform"           = $Server.TargetReadinesses.AzureSqlManagedInstance.AppliesToMigrationTargetPlatform
                "Databases List Ready For Migration"   = ($Server.TargetReadinesses.AzureSqlManagedInstance.DatabasesListReadyForMigration -join ',')
                "Num Of Databases Ready For Migration" = $Server.TargetReadinesses.AzureSqlManagedInstance.NumberOfDatabasesReadyForMigration.ToString()
                "Total Number Of Databases"            = $Server.TargetReadinesses.AzureSqlManagedInstance.TotalNumberOfDatabases.ToString()
                "Recommendation Status"                = $Server.TargetReadinesses.AzureSqlManagedInstance.RecommendationStatus
            }
            $data += $row
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "ServerTargetPlatformReadinesses"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while writing to the CSV file 'ServerTargetReadinesses'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath

    }

}

function Add-CsvFileDatabaseAssessments {
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Creating CSV file and writing data for 'DatabaseAssessments'" -ForegroundColor White

        # Create a list to hold the data
        $data = @()

        foreach ($Server in $JsonData.Servers) {
            foreach ($DatabaseAssessment in $Server.Databases.DatabaseAssessments) {
                $row = [PSCustomObject]@{
                    "Timestamp"                  = $DatabaseAssessment.Timestamp.ToString()
                    "Server Name"                = $DatabaseAssessment.ServerName
                    "Database Name"              = $DatabaseAssessment.DatabaseName
                    "Feature ID"                 = $DatabaseAssessment.FeatureId
                    "Issue Category"             = $DatabaseAssessment.IssueCategory
                    "More Information"           = $DatabaseAssessment.MoreInformation
                    "Description"                = $DatabaseAssessment.RuleMetadata.Description
                    "ID"                         = $DatabaseAssessment.RuleMetadata.Id
                    "Help Link"                  = $DatabaseAssessment.RuleMetadata.HelpLink
                    "Level"                      = $DatabaseAssessment.RuleMetadata.Level.ToString()
                    "Message"                    = $DatabaseAssessment.RuleMetadata.Message
                    "Rule Scope"                 = $DatabaseAssessment.RuleScope
                    "Applies to Target Platform" = $DatabaseAssessment.AppliesToMigrationTargetPlatform
                }
                $data += $row
            }
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "DatabaseAssessments"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating to the CSV file 'DatabaseAssessments'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
}

function Add-CsvFileDatabaseImpactedObjects {
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Creating CSV file and writing data for 'DatabaseImpactedObjects'" -ForegroundColor White

        $data = @()

        foreach ($Server in $JsonData.Servers) {
            foreach ($database in $Server.Databases.DatabaseAssessments) {
                $Col_ServerName = $database.ServerName
                $Col_DatabaseName = $database.DatabaseName
                $Col_FeatureID = $database.FeatureId
                $Col_AppliesToTargetPlatform = $database.AppliesToMigrationTargetPlatform

                foreach ($impactedObject in $database.ImpactedObjects) {
                    $row = [PSCustomObject]@{
                        "Server Name"                = $Col_ServerName
                        "Database Name"              = $Col_DatabaseName
                        "Feature ID"                 = $Col_FeatureID
                        "Object Name"                = $impactedObject.Name
                        "Object Type"                = $impactedObject.ObjectType
                        "Details"                    = $impactedObject.ImpactDetail
                        "Database Object Type"       = $impactedObject.DatabaseObjectType
                        "Applies To Target Platform" = $Col_AppliesToTargetPlatform
                    }
                    $data += $row
                }
            }
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "DatabaseImpactedObjects"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while writing to the CSV file 'DatabaseImpactedObjects'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
}

function Add-CsvFileDatabaseTargetReadinesses {
    param (
        [Parameter(Mandatory = $true)]
        [string] $JsonFile,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Creating CSV file and writing data for 'DatabaseTargetReadinesses'" -ForegroundColor White

        $data = @()

        foreach ($Server in $JsonData.Servers) {
            $Col_ServerName = $Server.Properties.ServerName

            foreach ($database in $Server.Databases) {
                # SQL DB
                $row = [PSCustomObject]@{
                    "Server Name"                = $Col_ServerName
                    "Database Name"              = $database.Properties.Name
                    "Applies To Target Platform" = $database.TargetReadinesses.AzureSqlDatabase.AppliesToMigrationTargetPlatform
                    "State"                      = $database.TargetReadinesses.AzureSqlDatabase.State
                    "Recommendation Status"      = $database.TargetReadinesses.AzureSqlDatabase.RecommendationStatus 
                    "Num Of Blocker Issues"      = $database.TargetReadinesses.AzureSqlDatabase.NumOfBlockerIssues.ToString()
                }
                $data += $row

                # SQL MI


                $row = [PSCustomObject]@{
                    "Server Name"                = $Col_ServerName
                    "Database Name"              = $database.Properties.Name
                    "Applies To Target Platform" = $database.TargetReadinesses.AzureSqlManagedInstance.AppliesToMigrationTargetPlatform
                    "State"                      = $database.TargetReadinesses.AzureSqlManagedInstance.State
                    "Recommendation Status"      = $database.TargetReadinesses.AzureSqlManagedInstance.RecommendationStatus
                    "Num Of Blocker Issues"      = $database.TargetReadinesses.AzureSqlManagedInstance.NumOfBlockerIssues.ToString()
                }
                $data += $row
            }
        }

        $CsvFilePathReturn = rename-fileToCsv -FolderPath $FolderPath -JsonFile $JsonFile -SectionName "DatabaseTargetReadinesses"
        # Export data to CSV
        $data | Export-Csv -Path $CsvFilePathReturn -NoTypeInformation
        Write-Host ("CSV file saved at '$CsvFilePathReturn'") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while writing to the CSV file 'DatabaseTargetReadinesses'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }

}


function Test-PathValidity {
    param (
        [string]$Path
    )

    # Check if the path is null or empty
    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Output ("Invalid path: Path is null or empty.")
        return $false
    }

    # Use Test-Path to check if the path exists
    if (Test-Path -Path $Path) {
        Write-Output "Valid path: $Path"
        return $true
    }
    else {
        Write-Output "Invalid path: $Path does not exist."
        return $false
    }
}

# Welcome to where the magic begins
write-host ("                                          ") -BackgroundColor DarkGreen
Write-Host ("  Welcome to CMA - DMS Report at scale    ") -ForegroundColor white -BackgroundColor DarkGreen
write-host ("  Customer Migration Accelerator for SQL  ") -ForegroundColor white -BackgroundColor DarkGreen
   
try {
    # Combine the folder path and file name
    $transcriptFileName = "Transcript-ReportCMA.txt"
    $transcriptFilePath = [IO.Path]::Combine($FolderPath, $transcriptFileName)

    if (Test-PathValidity -Path $transcriptFilePath) {
        Write-Host ("The path is valid. Initiating the transcript.") 
    }
    else {
        Write-Host ("The path is not valid. Stopping the execution.") -ForegroundColor Red
        exit
    }
    Start-Transcript -Path $transcriptFilePath -Append
}
catch {
    # Error message
    Write-Host ("An error occurred while creating the transcript file.") -ForegroundColor Red
    Write-Host ("Please check the log file for more details.") -ForegroundColor Red
    #Write-host ($_.Exception)
    Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    # Ensure the transcript is stopped
    Stop-Transcript
    Write-Host ("Transcript stopped.") -ForegroundColor White
    exit 
}


function rename-fileToCsv {
    param (
        [Parameter(Mandatory = $true)]
        [string] $FolderPath,
        [Parameter(Mandatory = $true)]
        [string]$JsonFile,
        [Parameter(Mandatory = $true)]
        [string]$SectionName        
    )
    # Check if the file path ends with '.json'
    if ($JsonFile -like "*.json") {
        
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($JsonFile)
        # Create the new CSV file name by appending the SectionName before the extension
        $csvFileName = "$baseFileName`_$SectionName.csv"
        $csvFilePath = [IO.Path]::Combine($FolderPath, $csvFileName)
    }
    else {

        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("Input file does not end with '.json'") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Stop-Transcript 
        exit
    }

    # delete the report if it already exists:
    if (Test-Path $csvFilePath ) {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("Deleting the file {0}." -f $csvFileName) -ForegroundColor Yellow
        Write-Host ("Note: The file already exists in {0} " -f $FolderPath ) -ForegroundColor Yellow
        Write-Host ("========================================================================================") -ForegroundColor White
        Remove-Item $csvFilePath
    }
    else {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("The file {0}. will be created in the directory: {1}." -f $csvFileName, $FolderPath ) -ForegroundColor White
        Write-Host ("========================================================================================") -ForegroundColor White
    }  
    
    return $csvFilePath
}

try {
    # Combine the folder path and file name
    $JsonFilePath = [IO.Path]::Combine($FolderPath, $JsonFile)
   
    if (Test-Path -Path $JsonFilePath) {
        $JsonData = Get-Content $JsonFilePath | Out-String | ConvertFrom-Json
    }
    else {
        Write-Host ("The file does not exist in {0}." -f $JsonFilePath) -ForegroundColor Red
        Stop-Transcript 
        exit
    }
}
catch {
    Write-Host ("========================================================================================") -ForegroundColor White
    Write-Host ("Input file is not a valid JSON.") -ForegroundColor Red
    Write-Host ("========================================================================================") -ForegroundColor White
    Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    # Ensure the transcript is stopped
    Stop-Transcript
    Write-Host "Transcript stopped."
    exit 
}


#New-ExecelFile -CSVFilePath $CSVFilePath
Add-CsvFileServerAssessments -JsonFile $JsonFile -jsonData $JsonData
Add-CsvFileServerProperties -JsonFile $JsonFile -jsonData $JsonData
Add-CsvFileServerImpactedObjects -JsonFile $JsonFile -jsonData $JsonData
Add-CsvFileServerTargetReadinesses -JsonFile $JsonFile -jsonData $JsonData

Add-CsvFileDatabaseAssessments -JsonFile $JsonFile -jsonData $JsonData
Add-CsvFileDatabaseImpactedObjects -JsonFile $JsonFile -jsonData $JsonData
Add-CsvFileDatabaseTargetReadinesses -JsonFile $JsonFile -jsonData $JsonData

Write-Host ("=======================================================================================")     
Write-Host (" The report has been generated and can be found at {0} " -f $CSVFilePath)  -BackgroundColor Black -ForegroundColor Green

Write-Host ("=======================================================================================") 
Write-Host ("  DMS Report at scale finished                                                         ") -ForegroundColor White -BackgroundColor DarkGreen
write-host ("  Customer Migration Accelerator for SQL                                               ") -ForegroundColor White -BackgroundColor DarkGreen
Write-Host ("=======================================================================================") -ForegroundColor White -BackgroundColor DarkGreen
Stop-Transcript