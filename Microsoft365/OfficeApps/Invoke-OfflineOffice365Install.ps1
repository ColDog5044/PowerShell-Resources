#Requires -Version 5.1

<#
.SYNOPSIS
    Installs Office 365 from an offline repository with optional update.

.DESCRIPTION
    Copies Office 365 installation files from a network repository and performs an offline installation using configuration.xml. 
    Optionally downloads the latest Office updates to the repository before installation.

.PARAMETER Repository
    The path to the Office 365 installation repository. Default is the network share.

.PARAMETER DestinationPath
    The local path where Office 365 files will be copied. Default is "C:\Office365Installer\".

.PARAMETER InstallerFile
    The name of the setup executable. Default is "setup.exe".

.PARAMETER UpdateApps
    Switch parameter to download the latest Office updates to the repository before installation.

.INPUTS
    None

.OUTPUTS
    None

.NOTES
    Version:          1.1
    Author:           Collin Laney
    Creation Date:    2025-09-02

    Exit Codes:
    Code 0 = Success
    Code 1 = Failure
    Code 2 = Script Error

.EXAMPLE
    .\Invoke-OfflineOffice365Install.ps1
    Runs the script with default parameters.

.EXAMPLE
    .\Invoke-OfflineOffice365Install.ps1 -UpdateApps
    Downloads latest updates to the repository before installation.

.EXAMPLE
    .\Invoke-OfflineOffice365Install.ps1 -Repository "D:\OfficeFiles\" -DestinationPath "C:\Temp\Office\" -UpdateApps
    Runs the script with custom paths and updates the repository.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$Repository = "\\networkshare\Office365Installer\",
    
    [Parameter(Mandatory = $false)]
    [string]$DestinationPath = "C:\Office365Installer\",
    
    [Parameter(Mandatory = $false)]
    [string]$InstallerFile = "setup.exe",
    
    [Parameter(Mandatory = $false)]
    [switch]$UpdateApps
)

begin {
    # Initialize variables and functions here
    function Test-IsElevated {
        $isElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($isElevated) {
            Write-Verbose "The script is running with elevated permissions"
            return $true
        }
        else {
            Write-Error "The script must be run with elevated permissions"
            exit 2
        }
    }
    
    function Update-OfficeRepository {
        param (
            [string]$RepositoryPath
        )
        
        Write-Output "Updating Office 365 repository with latest updates..."
        
        # Verify setup.exe exists in repository for download operation
        $RepoSetupPath = Join-Path $RepositoryPath $InstallerFile
        if (-not (Test-Path $RepoSetupPath)) {
            Write-Error "Setup file '$RepoSetupPath' not found in repository"
            exit 1
        }
        
        # Change to repository directory
        try {
            Set-Location $RepositoryPath -ErrorAction Stop
            Write-Verbose "Changed directory to repository: $RepositoryPath"
        }
        catch {
            Write-Error "Failed to change directory to repository '$RepositoryPath': $($_.Exception.Message)"
            exit 2
        }
        
        # Download latest updates
        try {
            Write-Verbose "Downloading latest Office 365 updates to repository"
            $DownloadProcess = Start-Process $RepoSetupPath -ArgumentList "/download configuration.xml" -Wait -PassThru -ErrorAction Stop
            
            if ($DownloadProcess.ExitCode -eq 0) {
                Write-Output "Office 365 repository updated successfully"
            }
            else {
                Write-Error "Failed to update Office 365 repository. Exit code: $($DownloadProcess.ExitCode)"
                exit 1
            }
        }
        catch {
            Write-Error "Failed to start Office 365 download process: $($_.Exception.Message)"
            exit 2
        }
    }
}

process {
    # Main code here
    Test-IsElevated
    
    # Verify source repository exists
    if (-not (Test-Path $Repository)) {
        Write-Error "Repository path '$Repository' does not exist"
        exit 1
    }
    
    # Update repository if requested
    if ($UpdateApps) {
        Update-OfficeRepository -RepositoryPath $Repository
    }
    
    # Remove existing destination if it exists
    if (Test-Path $DestinationPath) {
        Write-Verbose "Removing existing installation directory"
        try {
            Remove-Item $DestinationPath -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to remove existing directory '$DestinationPath': $($_.Exception.Message)"
            exit 2
        }
    }
    
    # Copy files from repository
    Write-Verbose "Copying files from '$Repository' to '$DestinationPath'"
    try {
        Copy-Item $Repository $DestinationPath -Recurse -Force -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to copy files from '$Repository' to '$DestinationPath': $($_.Exception.Message)"
        exit 2
    }
    
    # Verify setup.exe exists
    $SetupPath = Join-Path $DestinationPath $InstallerFile
    if (-not (Test-Path $SetupPath)) {
        Write-Error "Setup file '$SetupPath' not found"
        exit 1
    }
    
    # Verify configuration.xml exists
    $ConfigPath = Join-Path $DestinationPath "configuration.xml"
    if (-not (Test-Path $ConfigPath)) {
        Write-Error "Configuration file '$ConfigPath' not found"
        exit 1
    }
    
    try {
        Set-Location $DestinationPath -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to change directory to '$DestinationPath': $($_.Exception.Message)"
        exit 2
    }
    
    # Install the Office 365 apps
    Write-Verbose "Starting Office 365 installation"
    try {
        $Process = Start-Process $SetupPath -ArgumentList "/configure configuration.xml" -Wait -PassThru -ErrorAction Stop
        
        if ($Process.ExitCode -eq 0) {
            Write-Output "Office 365 installation completed successfully"
            exit 0
        }
        else {
            Write-Error "Office 365 installation failed with exit code: $($Process.ExitCode)"
            exit 1
        }
    }
    catch {
        Write-Error "Failed to start Office 365 installation: $($_.Exception.Message)"
        exit 2
    }
}

end {
    # Clean up code here
    if (Test-Path $DestinationPath) {
        Write-Verbose "Cleaning up installation files"
        try {
            Remove-Item $DestinationPath -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to clean up installation files: $($_.Exception.Message)"
        }
    }
}