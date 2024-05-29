<#
.SYNOPSIS
    Creates an Excel report for SQL Server assessment at scale.
.DESCRIPTION
    This PowerShell script creates an Excel report for SQL Server assessment at scale.
.PARAMETER FolderPath
    Specifies the path to the target directory.
.PARAMETER JsonFile
    Specifies the JSON file that contains the assessment.
.EXAMPLE
    PS> ./ReportCMA.ps1 -FolderPath "C:\temp\CMA-Report\" -JsonFile "SQLAssessment.json"
    ...
.NOTES
    Author: Tiago Balabuch
    Date: 29/05/2024
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

function New-ExecelFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath
    )

    try {
        Write-Host ("=======================================================================================") -ForegroundColor White
        Write-Host ("Creating Excel file") -ForegroundColor White
        # Create a new Excel application instance
        $excelApp = New-Object -ComObject Excel.Application
        
        $excelApp.Visible = $false
        $excelApp.ScreenUpdating = $false
        $excelApp.DisplayStatusBar = $false
        $excelApp.EnableEvents = $false
        
        # Add a new workbook (Excel file)
        $workbook = $excelApp.Workbooks.Add()

        $sheetNames = @(
            "DatabaseTargetReadinesses",
            "DatabaseImpactedObjects",
            "DatabaseAssessments",
            "ServerProperties",
            "ServerTargetReadinesses",
            "ServerImpactedObjects",
            "ServerAssessments"
        )

        foreach ($name in $sheetNames) {
            $worksheet = $workbook.Worksheets.Add()
            $worksheet.Name = $name
        }

        # Delete the worksheet
        $worksheetToDelete = $workbook.Sheets.Item("Sheet1")
        $worksheetToDelete.Delete()

        # Save the workbook
        $workbook.SaveAs($ExcelFilePath)

        # Close the workbook and Excel application
        $workbook.Close()
        $excelApp.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null

        Write-Host ("Excel file was successfully created") -ForegroundColor White
        Write-Host ("=======================================================================================") -ForegroundColor White
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while creating the excel file.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
        # Ensure the transcript is stopped
        Stop-Transcript
        Write-Host ("Transcript stopped.") -ForegroundColor White
        exit 
    }
}
function Add-ExcelSheetServerAssessments {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "ServerAssessments" 
    try {
        Write-Host ("=======================================================================================") -ForegroundColor White
        Write-Host ("Opening Excel file and writing to sheet 'ServerAssessments'") -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()
        
        $Header1 = "Server Assessment"
        $Headers = @("Status", "Timestamp", "Server Name", "Feature ID", "Issue Category", "More Information", "Description", "ID", "Help Link", "Level", "Message", "Rule Scope", "Applies to Target Platform")
        
        # Set headers
        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel Colors
        Add-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:M1").Merge()
        $worksheet.Range("A1:M2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:M2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:M2").Borders.Item(9).LineStyle = 1

        $Row = 3
        $Status = $JsonData.Status
        foreach ($Server in $JsonData.Servers) {
            foreach ($Assessment in $Server.ServerAssessments) {
                $worksheet.Cells.Item($Row, 1).Value2 = $Status
                $worksheet.Cells.Item($Row, 2).Value2 = $Assessment.Timestamp.ToString()
                $worksheet.Cells.Item($Row, 3).Value2 = $Assessment.ServerName
                $worksheet.Cells.Item($Row, 4).Value2 = $Assessment.FeatureId
                $worksheet.Cells.Item($Row, 5).Value2 = $Assessment.IssueCategory
                $worksheet.Cells.Item($Row, 6).Value2 = $Assessment.MoreInformation
                $worksheet.Cells.Item($Row, 7).Value2 = $Assessment.RuleMetadata.Description
                $worksheet.Cells.Item($Row, 8).Value2 = $Assessment.RuleMetadata.Id
                $worksheet.Cells.Item($Row, 9).Value2 = $Assessment.RuleMetadata.HelpLink
                $worksheet.Cells.Item($Row, 10).Value2 = $Assessment.RuleMetadata.Level.ToString()
                $worksheet.Cells.Item($Row, 11).Value2 = $Assessment.RuleMetadata.Message
                $worksheet.Cells.Item($Row, 12).Value2 = $Assessment.RuleScope
                $worksheet.Cells.Item($Row, 13).Value2 = $Assessment.AppliesToMigrationTargetPlatform

                $Row++
            }
        }

        # Save and close the workbook
        Write-Host ("Saving Excel sheet 'ServerAssessments'") -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating the Excel sheet 'ServerAssessments'..") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
        # Ensure the transcript is stopped
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host ("Closed Excel application") -ForegroundColor White
    }
}

function Add-ExcelSheetServerProperties {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "ServerProperties"
    try {
        Write-Host ("=======================================================================================") -ForegroundColor White
        Write-Host ("Opening Excel file and writing to sheet 'ServerProperties'") -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()

        $Header1 = "Server Properties"
        $Headers = @("Server Name", "FQDN", "Server Version", "Server Edition", "Hosting Platform", "Server Level", "Core Count", "Collation", "Hyperthread Ratio", "Logical CPU", "Physical CPU", "Max Memory In Use", "Number of Databases", "Total Database Size")

        # Set headers
        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel Colors
        Add-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:N1").Merge()
        $worksheet.Range("A1:N2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:N2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:N2").Borders.Item(9).LineStyle = 1

        $Row = 3

        foreach ($Server in $JsonData.Servers) {
            $Properties = $Server.Properties
            $worksheet.Cells.Item($Row, 1).Value2 = $Properties.ServerName
            $worksheet.Cells.Item($Row, 2).Value2 = $Properties.FQDN
            $worksheet.Cells.Item($Row, 3).Value2 = $Properties.ServerVersion
            $worksheet.Cells.Item($Row, 4).Value2 = $Properties.ServerEdition
            $worksheet.Cells.Item($Row, 5).Value2 = $Properties.ServerHostPlatform
            $worksheet.Cells.Item($Row, 6).Value2 = $Properties.ServerLevel
            $worksheet.Cells.Item($Row, 7).Value2 = $Properties.ServerCoreCount.ToString()
            $worksheet.Cells.Item($Row, 8).Value2 = $Properties.ServerCollation
            $worksheet.Cells.Item($Row, 9).Value2 = $Properties.HyperthreadRatio.ToString()
            $worksheet.Cells.Item($Row, 10).Value2 = $Properties.LogicalCpuCount.ToString()
            $worksheet.Cells.Item($Row, 11).Value2 = $Properties.PhysicalCpuCount.ToString()
            $worksheet.Cells.Item($Row, 12).Value2 = $Properties.MaxServerMemoryInUse.ToString()
            $worksheet.Cells.Item($Row, 13).Value2 = $Properties.NumberOfUserDatabases.ToString()
            $worksheet.Cells.Item($Row, 14).Value2 = $Properties.SumOfUserDatabasesSize.ToString()
    
            $Row++
        }

        # Save the workbook
        Write-Host ("Saving Excel sheet 'ServerProperties'") -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating the Excel sheet 'ServerProperties'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host ("Closed Excel application") -ForegroundColor White
    }
}


function Add-ExcelSheetServerImpactedObjects {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "ServerImpactedObjects" 
    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Opening Excel file and writing to sheet 'ServerImpactedObjects'" -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()

        # Set headers
        $Header1 = "Server Impacted Objects"
        $Headers = @("Server Name", "Feature ID", "Object Name", "Object Type", "Details", "Object Type", "Applies To Target Platform")

        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel Colors
        Add-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:G1").Merge()
        $worksheet.Range("A1:G2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:G2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:G2").Borders.Item(9).LineStyle = 1

        $RowStarts = 3

        foreach ($Server in $JsonData.Servers) {
            foreach ($ServerAssessment in $Server.ServerAssessments) {
                $Col_ServerName = $ServerAssessment.ServerName
                $Col_FeatureID = $ServerAssessment.FeatureId
                $Col_AppliesToTargetPlatform = $ServerAssessment.AppliesToMigrationTargetPlatform

                foreach ($ImpactedObject in $ServerAssessment.ImpactedObjects) {
                    $Col_Impact_ObjectName = $ImpactedObject.Name
                    $Col_Impact_ObjectType = $ImpactedObject.ObjectType
                    $Col_Impact_Details = $ImpactedObject.ImpactDetail
                    $Col_Impact_DatabaseObjectType = $ImpactedObject.DatabaseObjectType

                    $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_ServerName
                    $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_FeatureID
                    $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_Impact_ObjectName
                    $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_Impact_ObjectType
                    $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_Impact_Details
                    $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_Impact_DatabaseObjectType
                    $worksheet.Cells.Item($RowStarts, 7).Value2 = $Col_AppliesToTargetPlatform

                    $RowStarts++
                }
            }
        }

        # Save the workbook
        Write-Host "Saving Excel sheet 'ServerImpactedObjects'" -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating to the Excel sheet 'ServerImpactedObjects'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host "Closed Excel application" -ForegroundColor White
    }

}

function Add-ExcelSheetDatabaseAssessments {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "DatabaseAssessments" 
    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Opening Excel file and writing to sheet 'DatabaseAssessments'" -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()

        # Set headers
        $Header1 = "Database Assessment"
        $Headers = @("Timestamp", "Server Name", "Database Name", "Feature ID", "Issue Category", "More Information", "Description", "ID", "Help Link", "Level", "Message", "Rule Scope", "Applies to Target Platform")

        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel Colors
        Add-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:M1").Merge()
        $worksheet.Range("A1:M2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:M2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:M2").Borders.Item(9).LineStyle = 1

        $RowStarts = 3

        foreach ($Server in $JsonData.Servers) {
            foreach ($DatabaseAssessment in $Server.Databases.DatabaseAssessments) {

                $Col_ServerName = $DatabaseAssessment.ServerName
                $Col_DatabaseName = $DatabaseAssessment.DatabaseName
                $Col_FeatureID = $DatabaseAssessment.FeatureId
                $Col_IssueCategory = $DatabaseAssessment.IssueCategory
                $Col_MoreInformation = $DatabaseAssessment.MoreInformation
                $Col_RM_Description = $DatabaseAssessment.RuleMetadata.Description
                $Col_RM_ID = $DatabaseAssessment.RuleMetadata.Id
                $Col_RM_HelpLink = $DatabaseAssessment.RuleMetadata.HelpLink
                $Col_RM_Level = $DatabaseAssessment.RuleMetadata.Level.ToString()
                $Col_RM_Message = $DatabaseAssessment.RuleMetadata.Message
                $Col_RuleScope = $DatabaseAssessment.RuleScope
                $Col_AppliesToTargetPlatform = $DatabaseAssessment.AppliesToMigrationTargetPlatform
                $Col_AssessmentTimestamp = $DatabaseAssessment.Timestamp.ToString()

                ## item( row , column ) 
                $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_AssessmentTimestamp
                $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_ServerName
                $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_DatabaseName
                $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_FeatureID
                $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_IssueCategory
                $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_MoreInformation
                $worksheet.Cells.Item($RowStarts, 7).Value2 = $Col_RM_Description
                $worksheet.Cells.Item($RowStarts, 8).Value2 = $Col_RM_ID
                $worksheet.Cells.Item($RowStarts, 9).Value2 = $Col_RM_HelpLink
                $worksheet.Cells.Item($RowStarts, 10).Value2 = $Col_RM_Level
                $worksheet.Cells.Item($RowStarts, 11).Value2 = $Col_RM_Message
                $worksheet.Cells.Item($RowStarts, 12).Value2 = $Col_RuleScope
                $worksheet.Cells.Item($RowStarts, 13).Value2 = $Col_AppliesToTargetPlatform

                $RowStarts++
            }
        }

        # Save the workbook
        Write-Host "Saving Excel sheet 'DatabaseAssessments'" -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while generating to the Excel sheet 'DatabaseAssessments'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host "Closed Excel application" -ForegroundColor White
    }
}

function Add-ExcelSheetDatabaseImpactedObjects {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "DatabaseImpactedObjects"
    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Opening Excel file and writing to sheet 'DatabaseImpactedObjects'" -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()

        # Set headers
        $Header1 = "Database Impacted Objects"
        $Headers = @("Server Name", "Database Name", "Feature ID", "Object Name", "Object Type", "Details", "Object Type", "Applies To Target Platform")

        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel Colors
        Add-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:H1").Merge()
        $worksheet.Range("A1:H2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:H2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:H2").Borders.Item(9).LineStyle = 1

        $RowStarts = 3

        foreach ($Server in $JsonData.Servers) {
            foreach ($database in $Server.Databases.DatabaseAssessments) {
                $Col_ServerName = $database.ServerName
                $Col_DatabaseName = $database.DatabaseName
                $Col_FeatureID = $database.FeatureId
                $Col_AppliesToTargetPlatform = $database.AppliesToMigrationTargetPlatform

                foreach ($impactedObject in $database.ImpactedObjects) {
                    $Col_Impact_ObjectName = $impactedObject.Name
                    $Col_Impact_ObjectType = $impactedObject.ObjectType
                    $Col_Impact_Details = $impactedObject.ImpactDetail
                    $Col_Impact_DatabaseObjectType = $impactedObject.DatabaseObjectType

                    $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_ServerName
                    $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_DatabaseName
                    $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_FeatureID
                    $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_Impact_ObjectName
                    $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_Impact_ObjectType
                    $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_Impact_Details
                    $worksheet.Cells.Item($RowStarts, 7).Value2 = $Col_Impact_DatabaseObjectType
                    $worksheet.Cells.Item($RowStarts, 8).Value2 = $Col_AppliesToTargetPlatform

                    $RowStarts++
                }
            }
        }

        # Save the workbook
        Write-Host "Saving Excel sheet 'DatabaseImpactedObjects'" -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while writing to the Excel sheet 'DatabaseImpactedObjects'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host "Closed Excel application" -ForegroundColor White
    }
}

function Add-ExcelSheetDatabaseTargetReadinesses {
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "DatabaseTargetReadinesses"
    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Opening Excel file and writing to sheet 'DatabaseTargetReadinesses'" -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()

        # Set headers
        $Header1 = "Database Target Platform Readinesses"
        $Headers = @("Server Name", "Database Name", "Applies To Target Platform", "State", "Recommendation Status", "Num Of Blocker Issues")

        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel Colors
        Add-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:F1").Merge()
        $worksheet.Range("A1:F2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:F2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:F2").Borders.Item(9).LineStyle = 1

        $RowStarts = 3

        foreach ($Server in $JsonData.Servers) {
            $Col_ServerName = $Server.Properties.ServerName

            foreach ($database in $Server.Databases) {
                $Col_DatabaseName = $database.Properties.Name

                # SQL DB
                $Col_AppliesToTargetPlatform = $database.TargetReadinesses.AzureSqlDatabase.AppliesToMigrationTargetPlatform
                $Col_State = $database.TargetReadinesses.AzureSqlDatabase.State
                $Col_RecommendationStatus = $database.TargetReadinesses.AzureSqlDatabase.RecommendationStatus
                $Col_NumOfBlockerIssues = $database.TargetReadinesses.AzureSqlDatabase.NumOfBlockerIssues

                $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_ServerName
                $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_DatabaseName
                $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_AppliesToTargetPlatform
                $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_State
                $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_RecommendationStatus
                $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_NumOfBlockerIssues.ToString()

                $RowStarts++

                # SQL MI
                $Col_AppliesToTargetPlatform = $database.TargetReadinesses.AzureSqlManagedInstance.AppliesToMigrationTargetPlatform
                $Col_State = $database.TargetReadinesses.AzureSqlManagedInstance.State
                $Col_RecommendationStatus = $database.TargetReadinesses.AzureSqlManagedInstance.RecommendationStatus
                $Col_NumOfBlockerIssues = $database.TargetReadinesses.AzureSqlManagedInstance.NumOfBlockerIssues

                $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_ServerName
                $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_DatabaseName
                $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_AppliesToTargetPlatform
                $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_State
                $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_RecommendationStatus
                $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_NumOfBlockerIssues.ToString()

                $RowStarts++
            }
        }

        # Save the workbook
        Write-Host "Saving Excel sheet 'DatabaseTargetReadinesses'" -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while writing to the Excel sheet 'DatabaseTargetReadinesses'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host "Closed Excel application" -ForegroundColor White
    }

}

function Add-ExcelSheetServerTargetReadinesses { 
    param (
        [Parameter(Mandatory = $true)]
        [string] $ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [pscustomobject] $JsonData
    )

    $worksheetName = "ServerTargetReadinesses"
    try {
        Write-Host "=======================================================================================" -ForegroundColor White
        Write-Host "Opening Excel file and writing to sheet 'ServerTargetReadinesses'" -ForegroundColor White

        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $workbook = $excelApp.Workbooks.Open($ExcelFilePath)
        $worksheet = $workbook.Sheets.Item($worksheetName)
        $worksheet.Activate()

        # Set headers
        $Header1 = "Server Target Platform Readinesses"
        $Headers = @("Server Name", "Applies To Target Platform", "Databases List Ready For Migration", "Num Of Databases Ready For Migration", "Total Number Of Databases", "Recommendation Status")

        $worksheet.Cells.Item(1, 1).Value2 = $Header1
        for ($i = 0; $i -lt $Headers.Length; $i++) {
            $worksheet.Cells.Item(2, $i + 1).Value2 = $Headers[$i]
        }

        # Excel ColorsAdd-Type -AssemblyName System.Drawing
        $ExcelColor = [System.Drawing.Color]::FromArgb(218, 242, 208).ToArgb()
        $ExcelColor2 = [System.Drawing.Color]::FromArgb(217, 217, 217).ToArgb()

        # Merge and format headers
        $worksheet.Range("A1:F1").Merge()
        $worksheet.Range("A1:F2").Font.Bold = $true
        $worksheet.Cells.Item(1, 1).Font.Size = 14
        $worksheet.Cells.Item(1, 1).HorizontalAlignment = -4108
        $worksheet.Cells.Item(1, 1).Interior.Color = $ExcelColor
        $worksheet.Range("A2:F2").Interior.Color = $ExcelColor2
        $worksheet.Range("A1:F2").Borders.Item(9).LineStyle = 1

        $RowStarts = 3

        foreach ($Server in $JsonData.Servers) {
            $Col_ServerName = $Server.Properties.ServerName

            # SQL DB
            $Col_AppliesToTargetPlatform = $Server.TargetReadinesses.AzureSqlDatabase.AppliesToMigrationTargetPlatform
            $Col_DatabasesListReadyForMigration = $Server.TargetReadinesses.AzureSqlDatabase.DatabasesListReadyForMigration -join ','
            $Col_NumberOfDatabasesReadyForMigration = $Server.TargetReadinesses.AzureSqlDatabase.NumberOfDatabasesReadyForMigration.ToString()
            $Col_TotalNumberOfDatabases = $Server.TargetReadinesses.AzureSqlDatabase.TotalNumberOfDatabases.ToString()
            $Col_RecommendationStatus = $Server.TargetReadinesses.AzureSqlDatabase.RecommendationStatus

            $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_ServerName
            $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_AppliesToTargetPlatform 
            $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_DatabasesListReadyForMigration
            $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_NumberOfDatabasesReadyForMigration
            $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_TotalNumberOfDatabases
            $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_RecommendationStatus

            $RowStarts++

            # SQL MI
            $Col_AppliesToTargetPlatform = $Server.TargetReadinesses.AzureSqlManagedInstance.AppliesToMigrationTargetPlatform
            $Col_DatabasesListReadyForMigration = $Server.TargetReadinesses.AzureSqlManagedInstance.DatabasesListReadyForMigration -join ','
            $Col_NumberOfDatabasesReadyForMigration = $Server.TargetReadinesses.AzureSqlManagedInstance.NumberOfDatabasesReadyForMigration.ToString()
            $Col_TotalNumberOfDatabases = $Server.TargetReadinesses.AzureSqlManagedInstance.TotalNumberOfDatabases.ToString()
            $Col_RecommendationStatus = $Server.TargetReadinesses.AzureSqlManagedInstance.RecommendationStatus

            $worksheet.Cells.Item($RowStarts, 1).Value2 = $Col_ServerName
            $worksheet.Cells.Item($RowStarts, 2).Value2 = $Col_AppliesToTargetPlatform 
            $worksheet.Cells.Item($RowStarts, 3).Value2 = $Col_DatabasesListReadyForMigration
            $worksheet.Cells.Item($RowStarts, 4).Value2 = $Col_NumberOfDatabasesReadyForMigration
            $worksheet.Cells.Item($RowStarts, 5).Value2 = $Col_TotalNumberOfDatabases
            $worksheet.Cells.Item($RowStarts, 6).Value2 = $Col_RecommendationStatus

            $RowStarts++
        }

        # Save the workbook
        Write-Host "Saving Excel sheet 'ServerTargetReadinesses'" -ForegroundColor White
        $workbook.Save()
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("An error occurred while writing to the Excel sheet 'ServerTargetReadinesses'.") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor White
        Handle-Error -ErrorRecord $_ -FolderPath $FolderPath
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excelApp) { $excelApp.Quit() }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        Write-Host "Closed Excel application" -ForegroundColor White
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

# Check if the file path ends with '.json'
if ($JsonFile -like "*.json") {
    # Change the file path to the corresponding Excel file
    $excelFileName = $JsonFile -replace ".json", ".xlsx"
}
else {

    Write-Host ("========================================================================================") -ForegroundColor White
    Write-Host ("Input file does not end with '.json'") -ForegroundColor Red
    Write-Host ("========================================================================================") -ForegroundColor White
    Stop-Transcript 
    exit
}

try {
    # Combine the folder path and file name
    $ExcelFilePath = [IO.Path]::Combine($FolderPath, $excelFileName)
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

# delete the report if it already exists:
if (Test-Path $ExcelFilePath ) {
    Write-Host ("========================================================================================") -ForegroundColor White
    Write-Host ("Deleting the file {0}." -f $JsonFile) -ForegroundColor Yellow
    Write-Host ("Note: The file already exists in {0} " -f $FolderPath ) -ForegroundColor Yellow
    Write-Host ("========================================================================================") -ForegroundColor White
    Remove-Item $ExcelFilePath
}
else {
    Write-Host ("========================================================================================") -ForegroundColor White
    Write-Host ("The file {0}. will be created in the directory: {1}." -f $excelFileName, $FolderPath ) -ForegroundColor White
    Write-Host ("========================================================================================") -ForegroundColor White
}  
   
New-ExecelFile -ExcelFilePath $ExcelFilePath
Add-ExcelSheetServerAssessments -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Add-ExcelSheetServerTargetReadinesses -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Add-ExcelSheetServerProperties -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Add-ExcelSheetServerImpactedObjects -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Add-ExcelSheetDatabaseAssessments -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Add-ExcelSheetDatabaseImpactedObjects -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Add-ExcelSheetDatabaseTargetReadinesses -ExcelFilePath $ExcelFilePath -jsonData $JsonData
Write-Host ("=======================================================================================")     
Write-Host (" The report has been generated and can be found at {0} " -f $ExcelFilePath)  -BackgroundColor Black -ForegroundColor Green

Write-Host ("=======================================================================================") 
Write-Host ("  DMS Report at scale finished                                                         ") -ForegroundColor White -BackgroundColor DarkGreen
write-host ("  Customer Migration Accelerator for SQL                                               ") -ForegroundColor White -BackgroundColor DarkGreen
Write-Host ("=======================================================================================") -ForegroundColor White -BackgroundColor DarkGreen
Stop-Transcript