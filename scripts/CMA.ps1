<#
.SYNOPSIS
    Automation to assess, get SKU recommendation and collect data at scale.
.DESCRIPTION
    This PowerShell script is an uutomation to assess, get SKU recommendation and collect data at scale.
.PARAMETER Choice1
    ...
.EXAMPLE
    PS> ./ReportCMA.ps1 -FolderPath "C:\temp\CMA-Report\" -JsonFile "SQLAssessment.json"
    ...
.NOTES
    Author: Tiago Balabuch
    Date: 26/05/2024
    Version: 1.0
    GitHub: https://github.com/tiagobalabuch
    Copyright: (c) 2024 by Tiago Balabuch, licensed under MIT
    License: MIT https://opensource.org/licenses/MIT
#>

#Start-Transcript -Path C:\psLogs.txt -Append

write-host ("                                          ") -BackgroundColor DarkGreen
Write-Host ("           Welcome to CMA                 ") -ForegroundColor white -BackgroundColor DarkGreen
write-host ("  Customer Migration Accelerator for SQL  ") -ForegroundColor white -BackgroundColor DarkGreen

# Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
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
    $ErrorLogFileName = "ErrorLogCMALogs.log"
    $ErrorLogFilePath = [IO.Path]::Combine($FolderPath, $ErrorLogFileName)

    Add-Content -Path $ErrorLogFilePath -Value $ErrorDetails

    # Display the error details to the user
    Write-Error ($errorDetails)
}

function installChocolatey {
    param (
        [Parameter(Mandatory = $true)]
        [string] $rootPath
    )

    $chocoExePath = Join-Path $env:ProgramData "chocolatey\bin\choco.exe"

    if (Test-Path $chocoExePath) {

        Write-Host ("=======================================================================================") -ForegroundColor White
        Write-Host ("Chocolatey is installed.") -ForegroundColor White
        Write-Host ("=======================================================================================") -ForegroundColor White
    }
    else {
        Write-Host ("=======================================================================================") -ForegroundColor Yellow
        Write-Host ("Chocolatey is not installed.") -ForegroundColor Yellow
        Write-Host ("Installing Chocolatey.") -ForegroundColor Yellow
   
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))    
            Write-Host ("Chocolatey was installed successfully.") -ForegroundColor Green
            Write-Host ("=======================================================================================") -ForegroundColor Yellow
        }
        catch {
     
            Write-Host ("=======================================================================================") -ForegroundColor Red
            Write-Host ("Error occurred while installing Chocolatey.") -ForegroundColor Red
            Write-Host ("=======================================================================================") -ForegroundColor Red
            Handle-Error -ErrorRecord $_ -FolderPath $rootPath
            exit
        }
    }
}

# Functions 
function installDotNetRuntime {
    param (
        [Parameter(Mandatory = $true)]
        [string] $rootPath
    )


    #$dotnetVersion = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' | Get-ItemPropertyValue -Name Release -ErrorAction SilentlyContinue
    #if (Test-Path (Join-Path $env:ProgramFiles "dotnet\sdk\6.0.420")) { 
    #if ($dotnetVersion -eq 528040) {

    # Define the version you want to check
    $versionToCheck = "6.0.*"

    # Construct the path to check
    $dotnetPath = Join-Path -Path "$env:ProgramFiles\dotnet\sdk" -ChildPath $versionToCheck
    # Check if any version matching the pattern is installed
    if (Test-Path $dotnetPath) {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("DotNet 6.0 is installed.") -ForegroundColor White
        Write-Host ("========================================================================================") -ForegroundColor White
    }
    else {
        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("DotNet 6.0 is not installed.") -ForegroundColor Yellow
        Write-Host ("Installing DotNet 6.0")  -ForegroundColor Yellow
        try {
            installChocolatey -rootPath $rootPath
            choco install dotnet-6.0-sdk -y --force
            Write-Host ("DotNet 6.0 was installed successfully.") -ForegroundColor Green
            Write-Host ("========================================================================================") -ForegroundColor White
        }
        catch {
            Write-Host ("========================================================================================") -ForegroundColor Red
            Write-Host ("Error occurred while installing DotNet 6.0") -ForegroundColor Red
            Write-Host ("========================================================================================") -ForegroundColor Red
            Handle-Error -ErrorRecord $_ -FolderPath $rootPath
            exit
        }
    
    }
}

function installAzDataMigration {
    param (
        [Parameter(Mandatory = $true)]
        [string] $rootPath
    )
    try {
        # Validate if Az.DataMigration module is installed. 
        # In case it isn't, it will install two modules: Az.Accounts and Az.DataMigration. 
        # Az.Accounts is a prerequite for Az.DataMigration
        if (-Not(Get-Module -ListAvailable -Name Az.DataMigration)) {
            Write-Host ("========================================================================================") -ForegroundColor White
            Write-Host ("Az.DataMigration not available") -ForegroundColor Yellow
            Write-Host ("========================================================================================") -ForegroundColor White
            Write-Host ("Installing NuGet package manager") -ForegroundColor Yellow
            #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-PackageProvider -Name NuGet -Force -Confirm:$false
            Write-Host ("NuGet package manager was installed successfully") -ForegroundColor Green
            Write-Host ("========================================================================================") -ForegroundColor White
            Write-Host ("Az.DataMigration module requires Az.Accounts module.") -ForegroundColor Yellow
            Write-Host ("Az.Accounts module does not exist. Installing Az.Accounts module") -ForegroundColor Yellow
            Install-Module -Name Az.Accounts -Force -Confirm:$false 
            Write-Host ("Az.Accounts was installed successfully") -ForegroundColor Green
            Write-Host ("Az.DataMigration module does not exist. Installing Az.DataMigration module") -ForegroundColor Yellow
            Install-Module -Name Az.DataMigration -Force -Confirm:$false 
            Write-Host ("Az.DataMigration was installed successfully") -ForegroundColor Green
            Write-Host ("Importing Az.DataMigration and Az.Accounts modules") -ForegroundColor White
            Import-Module Az.Accounts -Force
            Import-Module Az.DataMigration -Force
            Write-Host ("Az.DataMigration and Az.Accounts modules imported successfully") -ForegroundColor Green
            Write-Host ("========================================================================================") -ForegroundColor White
        }
        else {
            # In case it is installed, it will update two modules(Az.Accounts and Az.DataMigration) if a new version is available
            # This ensure we're running always the latest version.  
            # Az.Accounts is a prerequite for Az.DataMigration
            $modulesToUpdate = "Az.Accounts", "Az.DataMigration"
            
            foreach ($moduleName in $modulesToUpdate) {
                $installedModule = Get-InstalledModule -Name $moduleName
                $currentVersion = $installedModule.Version.ToString()
                $onlineVersion = (Find-Module -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1).Version.ToString()
                
                if ([version]$currentVersion -ge [version]$onlineVersion) {
                    Write-Output "========================================================================================"

                    Write-Host ("Module: {0}" -f $moduleName ) -ForegroundColor Yellow
                    Write-Host ("Installed version {0} is equal or greater than {1}" -f $currentVersion, $onlineVersion ) -ForegroundColor White
                    Write-Host "Importing module $moduleName"
                    Import-Module -Name $moduleName -MinimumVersion $onlineVersion -Force 
                    Write-Host ("{0}  module imported successfully" -f $moduleName) -ForegroundColor Green
                    Write-Host "========================================================================================"
                }
                else {
                    Write-Host ("========================================================================================") -ForegroundColor Yellow
                    Write-Host ("Module: {0}" -f $moduleName ) -ForegroundColor Yellow
                    Write-Host ("Installed {0} is lower version than {1}" -f $currentVersion, $onlineVersion ) -ForegroundColor Yellow
                    Write-Host ("Updating to the lastest version: {0}" -f $onlineVersion ) -ForegroundColor White

                    # Get all versions of the specified module
                    $modulesToDelete = Get-InstalledModule -Name $moduleName | Select-Object -ExpandProperty Name -Unique

                    # Iterate through each version of the module and uninstall it
                    foreach ($module in $modulesToDelete) {
                        Uninstall-Module -Name $module -Force -Confirm:$false
                        Write-Host("Module {0} version {1} has been uninstalled" -f $module, $currentVersion ) -ForegroundColor White
                    }

                    # Install the new version
                    Install-Module -Name $moduleName -RequiredVersion $onlineVersion -Force -Confirm:$false 

                    Write-Host("New version {0} of {1} has been installed" -f $onlineVersion, $module ) -ForegroundColor White
                    # End of new code
                    Write-Host ("Module was updated successfully") -ForegroundColor Green
                    Write-Host ("Importing {0} module" -f $moduleName) -ForegroundColor White
                    Import-Module $moduleName -MinimumVersion $onlineVersion -Force
                    Write-Host ("{0} module imported successfully" -f $moduleName) -ForegroundColor Green
                    Write-Host ("========================================================================================") -ForegroundColor Yellow
                }
            }
        }
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor Red
        Write-Host ("Error occurred while installing or updating Az.DataMigration module") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor Red
        Handle-Error -ErrorRecord $_ -FolderPath $rootPath
        exit
    }
}

# Function to create folders
function CreateFolder {
    param (
        [Parameter(Mandatory = $true)]
        [String] $FolderName
    )
    try {
        if (Test-Path $FolderName) {
            Write-Host ("{0} folder already exists" -f $FolderName) -ForegroundColor Yellow
        }
        else {
            # Create directory if not exists
            New-Item $FolderName -ItemType Directory
            Write-Host ("{0} folder was created successfully" -f $FolderName) -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor Red
        Write-Host ("Error occurred while creating the folder {0}" -f $FolderName) -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor Red
        Handle-Error -ErrorRecord $_ -FolderPath $configFilePath
        exit
    }
}

function RenameFile {
    param (
        [Parameter(Mandatory = $true)]
        [String] $Path,
        [Parameter(Mandatory = $true)]
        [ValidateSet("AzureSqlDatabase", "AzureSqlManagedInstance", "AzureSqlVirtualMachine", "Any")] 
        [String] $TargetPlatform,
        [Parameter(Mandatory = $true)]
        [string] $rootPath
    )
    try {
        $files = Get-ChildItem -Path $path | Where-Object { 
            ($_.Extension -eq '.html' -or $_.Extension -eq '.json') -and 
            ($_.Name -notmatch 'AzureSqlDatabase|AzureSqlManagedInstance|AzureSqlVirtualMachine') 
        }
        
        foreach ($file in $files) {

            $newFileName = $file.FullName.Replace(".", "-" + $TargetPlatform + ".")
            
            # If a file already exists at the new path, delete it
            if (Test-Path $newFileName) {

                Write-Host ("========================================================================================") -ForegroundColor White
                Write-Host ("File '{0}' already exists. Deleting it..." -f $newFileName) -ForegroundColor White
                Remove-Item -Path $newFileName -Force
                Write-Host ("File '{0}' was removed successfully." -f $newFileName) -ForegroundColor White
            }
           
            #Renaming file
            Write-Host ("========================================================================================") -ForegroundColor White
            Write-Host ("Renaming file '{0}' to '{1}'..." -f $file.Name, $newFileName) -ForegroundColor White
            Rename-Item -Path $file.FullName -NewName $newFileName -Force
            Write-Host ("File '{0}' was renamed to '{1}' successfully." -f $file.Name, $newFileName) -ForegroundColor Green
            Write-Host ("========================================================================================") -ForegroundColor White
            
        }
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor Red
        Write-Host ("Error occurred while renaming files") -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor Red
        Handle-Error -ErrorRecord $_ -FolderPath $rootPath
        exit
    }
}

function CMA_Assessment {
    param (
        [Parameter(Mandatory = $true)]
        [string] $configFilePath
    )

    $configFile = "configAssessment.json"
    $configFilePathFinal = [IO.Path]::Combine($configFilePath, $configFile)

    if ([System.IO.File]::Exists($configFilePathFinal)) {

        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("Running SQL assessment. Please wait ... ") -ForegroundColor White
        Write-Host ("========================================================================================") -ForegroundColor White
        try {
            Get-AzDataMigrationAssessment -ConfigFilePath $configFilePathFinal -ErrorAction Stop
            Write-Host ("========================================================================================") -ForegroundColor White
        }
        catch {
            Write-Host ("========================================================================================") -ForegroundColor Red
            Write-Host ("Error occurred while running the assessment.") -ForegroundColor Red
            Handle-Error -ErrorRecord $_ -FolderPath $configFilePath
            Write-Host ("========================================================================================") -ForegroundColor Red
            exit
        }
    }
    else { 
        Write-Host ("========================================================================================") -ForegroundColor Red
        Write-Host ("Config file 'configAssessment.json' do not exist. Make sure you have a config file on this folder {0}" -f $configFilePath)-ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor Red
        exit
    }
}

function CMA_PerformanceDataCollect {
    param (
        [Parameter(Mandatory = $true)]
        [string] $configFilePath
    )
     
    $configFile = "configPerformanceDataCollect.json"
    $configFilePathFinal = [IO.Path]::Combine($configFilePath, $configFile)
    
    if ([System.IO.File]::Exists($configFilePathFinal)) {

        Write-Host ("For how many hours should the data be collected? Please specify the duration.") -ForegroundColor Green
        Write-Host ("=====================================================================") -ForegroundColor White
        Write-Host ("Valid inputs for the hour range should fall within the range of 1 to 168 hours.") -ForegroundColor White

        $helpDaysToHoursText = @"
Day 1: 24 hours
Day 2: 48 hours
Day 3: 72 hours
Day 4: 96 hours
Day 5: 120 hours
Day 6: 144 hours
Day 7: 168 hours
"@
        Write-Host $helpDaysToHoursText -ForegroundColor White
        
        $validHoursRange = 1..168
        do {
            [int]$hoursToPerform = Read-Host -Prompt "Please enter the number of hours"     
            if (-not $validHoursRange.Contains($hoursToPerform)) { 
                Write-Host ("Please select a valid input. Choose a number between 1 and 168.") -ForegroundColor Red
            }
        } until ($validHoursRange.Contains($hoursToPerform))

        [int]$duration = $hoursToPerform * 3600
     
        if ($duration -ge 3600) {
            $time = $duration / 3600
            $timeDesc = "hours"
        } 
        else {
            $time = $duration / 60
            $timeDesc = "minutes"
        }
      
        <#
        $debug = $false
        if (!$debug) { 
            $duration = 300 
        }
        #>

        Write-Host ("========================================================================================") -ForegroundColor White
        Write-Host ("Performing data collection.") -ForegroundColor White
        Write-Host ("Collecting data for {0} {1}" -f $time, $timeDesc) -ForegroundColor White
        Write-Host ("========================================================================================") -ForegroundColor White
      
        try {
            Get-AzDataMigrationPerformanceDataCollection -ConfigFilePath $configFilePathFinal -Time $duration -ErrorAction Stop
            Write-Host ("========================================================================================") -ForegroundColor White
            Write-Host ("Data collection completed successfully.") -ForegroundColor Green
            Write-Host ("========================================================================================") -ForegroundColor White
        }
        catch {
            Write-Host ("========================================================================================") -ForegroundColor Red
            Write-Host ("Error occurred during data collection.") -ForegroundColor Red
            Handle-Error -ErrorRecord $_ -FolderPath $configFilePath
            Write-Host ("========================================================================================") -ForegroundColor Red
            exit
        }
    }
    else { 
        Write-Host ("========================================================================================") -ForegroundColor Red
        Write-Host ("Config file 'configPerformanceDataCollect.json' does not exist. Ensure the file is in the folder: {0}" -f $configFilePath) -ForegroundColor Red
        Write-Host ("========================================================================================") -ForegroundColor Red
        exit
    }
}

function CMA_SkuRecommendation {
    param (
        [Parameter(Mandatory = $true)]
        [string] $configFilePath,
        [Parameter(Mandatory = $true)]
        [string] $DataCollectionPath
    )
    try {
        if (![string]::IsNullOrEmpty($configFilePath)) {

            $TargetPlatforms = "AzureSqlDatabase", "AzureSqlManagedInstance", "AzureSqlVirtualMachine", "Any" 

            foreach ($TargetPlatform in $TargetPlatforms) {
                Write-Host ("========================================================================================") -ForegroundColor White
                Write-Host ("Running SKU recommendation for {0}. Please wait ..." -f $TargetPlatform) -ForegroundColor White
                Write-Host ("========================================================================================") -ForegroundColor White

                try {  
                    Get-AzDataMigrationSkuRecommendation -OutputFolder $DataCollectionPath -Overwrite -DisplayResult -TargetPlatform $TargetPlatform -ScalingFactor 100
                    RenameFile -TargetPlatform $TargetPlatform -Path $DataCollectionPath -rootPath $configFilePath
                    Write-Host ("SKU recommendation for {0} completed successfully." -f $TargetPlatform) -ForegroundColor Green
                }
                catch {
                    Write-Host ("Error occurred during SKU recommendation for {0}." -f $TargetPlatform) -ForegroundColor Red
                    Write-Host ("Error details: {0}" -f $_.Exception.Message) -ForegroundColor Red
                }
            }
        }
        else {
                Write-Host ("========================================================================================") -ForegroundColor Red
                Write-Host ("Parameter configFilePath is null or empty. Please provide a valid path.") -ForegroundColor Red
                Write-Host ("========================================================================================") -ForegroundColor Red
                exit
            }
    }
    catch {
        Write-Host ("========================================================================================") -ForegroundColor Red
        Write-Host ("Error occurred while running the SKU recommendation.") -ForegroundColor Red
        Handle-Error -ErrorRecord $_ -FolderPath $configFilePath
        Write-Host ("========================================================================================") -ForegroundColor Red
        exit
    }

}

# Magic start here!!! :)

Write-Host ("Please select the operation to perform:") -ForegroundColor Green
Write-Host ("=====================================================================")
Write-Host ("1. Perform Assessment, Performance Data Collection, and SKU recommendation")
Write-Host ("2. Perform Assessment only")
Write-Host ("3. Perform Performance Data Collection only")
Write-Host ("4. Perform SKU recommendation only")
Write-Host ("5. Exit")

do {
    [int]$operationToPerform = Read-Host -Prompt "Enter the operation number (1-5)"
    if ($operationToPerform -lt 1 -or $operationToPerform -gt 5) { 
        Write-Host ("Please specify a number between 1 and 5.") 
    }
} until ($operationToPerform -ge 1 -and $operationToPerform -le 5)

if ($operationToPerform -eq 5) {
    Write-Host ("Exiting the script based on user input.")
    exit
}

$configFilePath = Read-Host -Prompt ("Please specify the root folder where config files are")

Write-Host ("=======================================================================================")
Write-Host ("Start checking prerequisites") 
Write-Host ("=======================================================================================")
Write-Host ("Checking folders") 
Write-Host ("=======================================================================================") 
try {
    #Destination Path
    $rootPath = $configFilePath
    $CMAAssessmentPath = [IO.Path]::Combine($rootPath, "CMAAssessment")
    $CMADataCollectionPath = [IO.Path]::Combine($rootPath, "CMADataCollection")
    
    #Create Folders
    CreateFolder $CMAAssessmentPath
    CreateFolder $CMADataCollectionPath
}
catch {
    Write-Host ("Error to create folders")  -ForegroundColor Red
    exit
}

Write-Host ("=======================================================================================")
Write-Host ("Reviewing DotNet 6.0 on this machine...") 
Write-Host ("=======================================================================================")

# Ensuring that all prerequisites are installed or updated!
try {
    installDotNetRuntime -rootPath $rootPath
}
catch {
    Write-Host ("Error reviewing DotNet 6.0 on this machine...")
    exit
}

Write-Host ("=======================================================================================")
Write-Host ("Reviewing Az.DataMigration PowerShell modules on this machine...") 
Write-Host ("=======================================================================================")

try {
    installAzDataMigration -rootPath $rootPath
}
catch {
    Write-Host ("Error reviewing Azure Data Migration PowerShell modules on this machine...") -ForegroundColor Red
    exit
}

# Perform selected operation
switch ($operationToPerform) {
    1 {
        try {
            # Perform Assessment
            CMA_Assessment $configFilePath
        }
        catch {
            Write-Host ("Error running the assessment.") -ForegroundColor Red
        }

        try {
            # Perform Performance Data Collection
            CMA_PerformanceDataCollect -configFilePath $configFilePath
        }
        catch {
            Write-Host ("Error performing data collection.") -ForegroundColor Red
        }

        try {
            # Perform SKU recommendation
            CMA_SkuRecommendation -configFilePath $configFilePath -DataCollectionPath $CMADataCollectionPath
        }
        catch {
            Write-Host ("Error running the SKU recommendation.") -ForegroundColor Red
        }
    }
    2 {
        try {
            # Perform Assessment only
            CMA_Assessment $configFilePath
        }
        catch {
            Write-Host ("Error running the assessment.") -ForegroundColor Red
        }
    }
    3 {
        try {
            # Perform Performance Data Collection only
            CMA_PerformanceDataCollect -configFilePath $configFilePath
        }
        catch {
            Write-Host ("Error performing data collection.") -ForegroundColor Red
        }
    }
    4 {
        try {
            # Perform SKU recommendation only
            CMA_SkuRecommendation -configFilePath $configFilePath -DataCollectionPath $CMADataCollectionPath
        }
        catch {
            Write-Host ("Error running the SKU recommendation.") -ForegroundColor Red
        }
    }
}

Write-Host ("==========================================") 
write-host ("                                          ") -BackgroundColor DarkGreen
Write-Host ("          Thanks for using CMA            ") -ForegroundColor white -BackgroundColor DarkGreen
write-host ("  Customer Migration Accelerator for SQL  ") -ForegroundColor white -BackgroundColor DarkGreen
Write-Host ("==========================================") -ForegroundColor white -BackgroundColor DarkGreen

