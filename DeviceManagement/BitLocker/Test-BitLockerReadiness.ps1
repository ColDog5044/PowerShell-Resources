#Requires -Version 5.1

<#
.SYNOPSIS
    BitLocker readiness assessment and status checker.

.DESCRIPTION
    This script performs a comprehensive assessment of BitLocker readiness and current status 
    on the local system. It checks TPM availability, hardware requirements, BitLocker feature 
    availability, current encryption status, and system configuration to determine if the 
    system is ready for BitLocker deployment. The script provides informational output about 
    current encryption status without requiring BitLocker to be enabled.

.PARAMETER OutputPath
    Optional path to export the BitLocker report to a file.

.PARAMETER IncludeRecoveryKeys
    Switch to include recovery key information in the output (requires elevated permissions).

.PARAMETER CheckAllDrives
    Switch to check all available drives, not just the system drive.

.INPUTS
    None

.OUTPUTS
    BitLocker status report object and optional file export.

.NOTES
    Version:          1.0
    Author:           Collin Laney
    Creation Date:    2025-09-02

    Exit Codes:
    Code 0 = Success - System ready for BitLocker or BitLocker functioning properly
    Code 1 = Warning - BitLocker cannot be enabled due to requirements not met
    Code 2 = Script Error - Unable to run checks

.EXAMPLE
    .\Test-BitLockerReadiness.ps1
    Runs a basic BitLocker check on the system drive.

.EXAMPLE
    .\Test-BitLockerReadiness.ps1 -CheckAllDrives -IncludeRecoveryKeys -OutputPath "C:\Reports\BitLocker-Report.txt"
    Performs a comprehensive check of all drives with recovery key information and exports to file.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeRecoveryKeys,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckAllDrives
)

begin {
    # Initialize variables and functions here
    $script:ExitCode = 0
    $script:BitLockerReport = @{
        ComputerName        = $env:COMPUTERNAME
        ScanDate            = Get-Date
        TPMStatus           = $null
        BitLockerCapability = $null
        DriveStatus         = @()
        SystemRequirements  = @{}
        RecoveryKeys        = @()
        Policies            = @{}
        Warnings            = @()
        Errors              = @()
    }

    function Test-IsElevated {
        $isElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($isElevated) {
            Write-Verbose "The script is running with elevated permissions"
            return $true
        }
        else {
            Write-Warning "Some checks require elevated permissions for complete results"
            return $false
        }
    }

    function Test-TPMStatus {
        Write-Verbose "Checking TPM status..."
        try {
            $tpm = Get-WmiObject -Namespace "Root\CIMv2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue
            if ($tpm) {
                $tpmInfo = @{
                    Present             = $true
                    Enabled             = $tpm.IsEnabled_InitialValue
                    Activated           = $tpm.IsActivated_InitialValue
                    Owned               = $tpm.IsOwned_InitialValue
                    Version             = $tpm.SpecVersion
                    ManufacturerVersion = $tpm.ManufacturerVersion
                }
            }
            else {
                # Try alternative method
                $tpmWin32 = Get-CimInstance -ClassName Win32_Tpm -Namespace "Root\CIMv2\Security\MicrosoftTpm" -ErrorAction SilentlyContinue
                if ($tpmWin32) {
                    $tpmInfo = @{
                        Present             = $true
                        Enabled             = $tpmWin32.IsEnabled_InitialValue
                        Activated           = $tpmWin32.IsActivated_InitialValue
                        Owned               = $tpmWin32.IsOwned_InitialValue
                        Version             = $tpmWin32.SpecVersion
                        ManufacturerVersion = $tpmWin32.ManufacturerVersion
                    }
                }
                else {
                    $tpmInfo = @{
                        Present             = $false
                        Enabled             = $false
                        Activated           = $false
                        Owned               = $false
                        Version             = "Not Available"
                        ManufacturerVersion = "Not Available"
                    }
                }
            }
        }
        catch {
            Write-Warning "Unable to retrieve TPM information: $($_.Exception.Message)"
            $tpmInfo = @{
                Present             = "Unknown"
                Enabled             = "Unknown"
                Activated           = "Unknown"
                Owned               = "Unknown"
                Version             = "Unknown"
                ManufacturerVersion = "Unknown"
                Error               = $_.Exception.Message
            }
        }
        return $tpmInfo
    }

    function Test-BitLockerCapability {
        Write-Verbose "Checking BitLocker capability..."
        try {
            $capability = @{
                BitLockerAvailable = $false
                SupportedVolumes   = @()
                RequirementsCheck  = @{}
            }

            # Check if BitLocker feature is available
            $feature = Get-WindowsOptionalFeature -Online -FeatureName "BitLocker" -ErrorAction SilentlyContinue
            if ($feature -and $feature.State -eq "Enabled") {
                $capability.BitLockerAvailable = $true
            }

            # Check system requirements
            $capability.RequirementsCheck = @{
                OSVersion         = (Get-WmiObject -Class Win32_OperatingSystem).Version
                OSEdition         = (Get-WmiObject -Class Win32_OperatingSystem).Caption
                UEFIBoot          = Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State"
                SecureBootEnabled = $false
            }

            # Check Secure Boot status
            try {
                $secureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
                $capability.RequirementsCheck.SecureBootEnabled = $secureBoot
            }
            catch {
                $capability.RequirementsCheck.SecureBootEnabled = "Unable to determine"
            }

            return $capability
        }
        catch {
            Write-Error "Error checking BitLocker capability: $($_.Exception.Message)"
            return $null
        }
    }

    function Get-BitLockerVolumeStatus {
        param(
            [string[]]$DriveLetters
        )
        
        Write-Verbose "Checking BitLocker volume status..."
        $volumeStatus = @()
        
        try {
            # Check if BitLocker module is available
            $bitlockerModule = Get-Module -ListAvailable -Name BitLocker -ErrorAction SilentlyContinue
            if (-not $bitlockerModule) {
                Write-Verbose "BitLocker PowerShell module not available - limited volume status information"
                return @()
            }
            
            # Import BitLocker module if not already loaded
            if (-not (Get-Module -Name BitLocker)) {
                Import-Module BitLocker -ErrorAction SilentlyContinue
            }
            
            if ($DriveLetters) {
                foreach ($drive in $DriveLetters) {
                    $volume = Get-BitLockerVolume -MountPoint $drive -ErrorAction SilentlyContinue
                    if ($volume) {
                        $volumeStatus += $volume
                    }
                    else {
                        # Create a basic status object for drives without BitLocker info
                        $volumeStatus += [PSCustomObject]@{
                            MountPoint           = $drive
                            VolumeStatus         = "FullyDecrypted"
                            EncryptionMethod     = "None"
                            ProtectionStatus     = "Off"
                            EncryptionPercentage = 0
                        }
                    }
                }
            }
            else {
                $volumeStatus = Get-BitLockerVolume -ErrorAction SilentlyContinue
                if (-not $volumeStatus) {
                    # If no BitLocker volumes found, get basic drive info
                    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
                    foreach ($drive in $drives) {
                        $volumeStatus += [PSCustomObject]@{
                            MountPoint           = $drive.DeviceID
                            VolumeStatus         = "FullyDecrypted"
                            EncryptionMethod     = "None"
                            ProtectionStatus     = "Off"
                            EncryptionPercentage = 0
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Error retrieving BitLocker volume status: $($_.Exception.Message)"
            $script:BitLockerReport.Errors += "BitLocker volume status check failed: $($_.Exception.Message)"
        }
        
        return $volumeStatus
    }

    function Get-BitLockerRecoveryKeys {
        param(
            [string]$VolumeId
        )
        
        if (-not $script:IsElevated) {
            Write-Verbose "Elevated permissions required for recovery key information"
            return @()
        }
        
        try {
            $recoveryKeys = @()
            $volume = Get-BitLockerVolume | Where-Object { $_.VolumeStatus -ne "FullyDecrypted" }
            
            foreach ($vol in $volume) {
                $keys = $vol | Get-BitLockerVolume | Select-Object -ExpandProperty KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" }
                foreach ($key in $keys) {
                    $recoveryKeys += @{
                        VolumeId         = $vol.VolumeId
                        MountPoint       = $vol.MountPoint
                        KeyId            = $key.KeyProtectorId
                        RecoveryPassword = if ($IncludeRecoveryKeys) { $key.RecoveryPassword } else { "Hidden (use -IncludeRecoveryKeys)" }
                    }
                }
            }
            
            return $recoveryKeys
        }
        catch {
            Write-Warning "Error retrieving recovery keys: $($_.Exception.Message)"
            return @()
        }
    }

    function Test-BitLockerPolicies {
        Write-Verbose "Checking BitLocker policies..."
        $policies = @{}
        
        try {
            # Check common BitLocker registry policies
            $policyPath = "HKLM:\SOFTWARE\Policies\Microsoft\FVE"
            
            if (Test-Path $policyPath) {
                $policyKeys = Get-ChildItem $policyPath -Recurse -ErrorAction SilentlyContinue
                foreach ($key in $policyKeys) {
                    $keyName = $key.Name.Split('\')[-1]
                    $policies[$keyName] = @{}
                    
                    $properties = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                    if ($properties) {
                        foreach ($prop in $properties.PSObject.Properties) {
                            if ($prop.Name -notlike "PS*") {
                                $policies[$keyName][$prop.Name] = $prop.Value
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Error checking BitLocker policies: $($_.Exception.Message)"
        }
        
        return $policies
    }

    function Write-BitLockerReport {
        param(
            [hashtable]$Report
        )
        
        Write-Host "`n=== BitLocker Status Report ===" -ForegroundColor Cyan
        Write-Host "Computer: $($Report.ComputerName)" -ForegroundColor Green
        Write-Host "Scan Date: $($Report.ScanDate)" -ForegroundColor Green
        
        # TPM Status
        Write-Host "`n--- TPM Status ---" -ForegroundColor Yellow
        if ($Report.TPMStatus) {
            Write-Host "TPM Present: $($Report.TPMStatus.Present)" -ForegroundColor $(if ($Report.TPMStatus.Present) { "Green" } else { "Red" })
            Write-Host "TPM Enabled: $($Report.TPMStatus.Enabled)" -ForegroundColor $(if ($Report.TPMStatus.Enabled) { "Green" } else { "Red" })
            Write-Host "TPM Activated: $($Report.TPMStatus.Activated)" -ForegroundColor $(if ($Report.TPMStatus.Activated) { "Green" } else { "Red" })
            Write-Host "TPM Version: $($Report.TPMStatus.Version)"
        }
        
        # BitLocker Capability
        Write-Host "`n--- BitLocker Capability ---" -ForegroundColor Yellow
        if ($Report.BitLockerCapability) {
            Write-Host "BitLocker Available: $($Report.BitLockerCapability.BitLockerAvailable)" -ForegroundColor $(if ($Report.BitLockerCapability.BitLockerAvailable) { "Green" } else { "Red" })
            Write-Host "OS Version: $($Report.BitLockerCapability.RequirementsCheck.OSVersion)"
            Write-Host "UEFI Boot: $($Report.BitLockerCapability.RequirementsCheck.UEFIBoot)" -ForegroundColor $(if ($Report.BitLockerCapability.RequirementsCheck.UEFIBoot) { "Green" } else { "Yellow" })
            Write-Host "Secure Boot: $($Report.BitLockerCapability.RequirementsCheck.SecureBootEnabled)" -ForegroundColor $(if ($Report.BitLockerCapability.RequirementsCheck.SecureBootEnabled -eq $true) { "Green" } else { "Yellow" })
        }
        
        # Drive Status
        Write-Host "`n--- Drive Encryption Status ---" -ForegroundColor Yellow
        if ($Report.DriveStatus.Count -gt 0) {
            foreach ($drive in $Report.DriveStatus) {
                $statusColor = switch ($drive.VolumeStatus) {
                    "FullyEncrypted" { "Green" }
                    "EncryptionInProgress" { "Yellow" }
                    "DecryptionInProgress" { "Yellow" }
                    "FullyDecrypted" { "Red" }
                    default { "White" }
                }
                Write-Host "Drive $($drive.MountPoint): $($drive.VolumeStatus)" -ForegroundColor $statusColor
                Write-Host "  Encryption Method: $($drive.EncryptionMethod)"
                Write-Host "  Protection Status: $($drive.ProtectionStatus)"
                if ($drive.EncryptionPercentage) {
                    Write-Host "  Encryption Progress: $($drive.EncryptionPercentage)%"
                }
            }
        }
        else {
            Write-Host "No BitLocker volumes found" -ForegroundColor Red
        }
        
        # Recovery Keys
        if ($Report.RecoveryKeys.Count -gt 0) {
            Write-Host "`n--- Recovery Keys ---" -ForegroundColor Yellow
            foreach ($key in $Report.RecoveryKeys) {
                Write-Host "Volume $($key.MountPoint): $($key.RecoveryPassword)"
            }
        }
        
        # Warnings and Errors
        if ($Report.Warnings.Count -gt 0) {
            Write-Host "`n--- Warnings ---" -ForegroundColor Yellow
            foreach ($warning in $Report.Warnings) {
                Write-Host "⚠️  $warning" -ForegroundColor Yellow
            }
        }
        
        if ($Report.Errors.Count -gt 0) {
            Write-Host "`n--- Errors ---" -ForegroundColor Red
            foreach ($err in $Report.Errors) {
                Write-Host "❌ $err" -ForegroundColor Red
            }
        }
    }

    # Initialize elevation check
    $script:IsElevated = Test-IsElevated
}

process {
    # Main code here
    Write-Host "Starting BitLocker comprehensive check..." -ForegroundColor Cyan
    
    try {
        # Check TPM Status
        Write-Verbose "Performing TPM status check..."
        $script:BitLockerReport.TPMStatus = Test-TPMStatus
        
        # Check BitLocker Capability
        Write-Verbose "Checking BitLocker capability..."
        $script:BitLockerReport.BitLockerCapability = Test-BitLockerCapability
        
        # Determine which drives to check
        $drivesToCheck = @()
        if ($CheckAllDrives) {
            # Get all drives with drive letters
            $drivesToCheck = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object { $_.DeviceID }
            Write-Verbose "Checking all drives: $($drivesToCheck -join ', ')"
        }
        else {
            # Just check system drive
            $systemDrive = $env:SystemDrive
            $drivesToCheck = @($systemDrive)
            Write-Verbose "Checking system drive: $systemDrive"
        }
        
        # Check BitLocker Volume Status
        Write-Verbose "Checking BitLocker volume status..."
        $volumes = Get-BitLockerVolumeStatus -DriveLetters $drivesToCheck
        $script:BitLockerReport.DriveStatus = $volumes
        
        # Get Recovery Keys if requested and elevated
        if ($IncludeRecoveryKeys) {
            Write-Verbose "Retrieving recovery keys..."
            $script:BitLockerReport.RecoveryKeys = Get-BitLockerRecoveryKeys
        }
        
        # Check BitLocker Policies
        Write-Verbose "Checking BitLocker policies..."
        $script:BitLockerReport.Policies = Test-BitLockerPolicies
        
        # Analyze results and set warnings/errors
        Write-Verbose "Analyzing results..."
        
        # Check if TPM is available and properly configured
        if ($script:BitLockerReport.TPMStatus.Present -eq $false) {
            $script:BitLockerReport.Warnings += "TPM is not present on this system. BitLocker can still be used with alternative authentication methods."
        }
        elseif ($script:BitLockerReport.TPMStatus.Enabled -eq $false) {
            $script:BitLockerReport.Warnings += "TPM is present but not enabled. Enable TPM in BIOS/UEFI settings before enabling BitLocker."
            $script:ExitCode = 1
        }
        elseif ($script:BitLockerReport.TPMStatus.Activated -eq $false) {
            $script:BitLockerReport.Warnings += "TPM is enabled but not activated. Windows may need to initialize the TPM before enabling BitLocker."
            $script:ExitCode = 1
        }
        
        # Check if BitLocker is available
        if ($script:BitLockerReport.BitLockerCapability -and -not $script:BitLockerReport.BitLockerCapability.BitLockerAvailable) {
            $script:BitLockerReport.Warnings += "BitLocker feature is not available on this system or Windows edition."
            $script:ExitCode = 1
        }
        
        # Check encryption status (informational only - not required)
        $encryptedVolumes = $script:BitLockerReport.DriveStatus | Where-Object { $_.VolumeStatus -eq "FullyEncrypted" }
        $unencryptedVolumes = $script:BitLockerReport.DriveStatus | Where-Object { $_.VolumeStatus -eq "FullyDecrypted" }
        $inProgressVolumes = $script:BitLockerReport.DriveStatus | Where-Object { $_.VolumeStatus -like "*InProgress" }
        
        if ($unencryptedVolumes.Count -gt 0) {
            foreach ($vol in $unencryptedVolumes) {
                $script:BitLockerReport.Warnings += "Drive $($vol.MountPoint) is not encrypted. Ready for BitLocker enablement when needed."
            }
        }
        
        if ($inProgressVolumes.Count -gt 0) {
            foreach ($vol in $inProgressVolumes) {
                $script:BitLockerReport.Warnings += "Drive $($vol.MountPoint) has encryption/decryption in progress ($($vol.EncryptionPercentage)%)."
            }
        }
        
        if ($encryptedVolumes.Count -gt 0) {
            foreach ($vol in $encryptedVolumes) {
                $script:BitLockerReport.Warnings += "Drive $($vol.MountPoint) is already encrypted with BitLocker."
            }
        }
        
        # Check for Secure Boot (recommended but not required)
        if ($script:BitLockerReport.BitLockerCapability.RequirementsCheck.SecureBootEnabled -eq $false) {
            $script:BitLockerReport.Warnings += "Secure Boot is not enabled. While not required, it's recommended for enhanced security when enabling BitLocker."
        }
        
        # Success message based on readiness
        if ($script:ExitCode -eq 0) {
            Write-Host "✅ BitLocker readiness check completed successfully - System is ready for BitLocker!" -ForegroundColor Green
        }
        elseif ($script:ExitCode -eq 1) {
            Write-Host "⚠️ BitLocker readiness check found requirements not met - Address issues before enabling BitLocker" -ForegroundColor Yellow
        }
        else {
            Write-Host "❌ BitLocker readiness check encountered errors - Review script execution issues" -ForegroundColor Red
        }
        
    }
    catch {
        Write-Error "Critical error during BitLocker check: $($_.Exception.Message)"
        $script:BitLockerReport.Errors += "Script execution error: $($_.Exception.Message)"
        $script:ExitCode = 2
    }
}

end {
    # Clean up code here and display results
    
    # Display the comprehensive report
    Write-BitLockerReport -Report $script:BitLockerReport
    
    # Export to file if requested
    if ($OutputPath) {
        try {
            Write-Verbose "Exporting report to: $OutputPath"
            
            # Create a formatted text report
            $reportText = @"
BitLocker Status Report
======================
Computer: $($script:BitLockerReport.ComputerName)
Scan Date: $($script:BitLockerReport.ScanDate)

TPM Status:
-----------
Present: $($script:BitLockerReport.TPMStatus.Present)
Enabled: $($script:BitLockerReport.TPMStatus.Enabled)
Activated: $($script:BitLockerReport.TPMStatus.Activated)
Owned: $($script:BitLockerReport.TPMStatus.Owned)
Version: $($script:BitLockerReport.TPMStatus.Version)

BitLocker Capability:
--------------------
Available: $($script:BitLockerReport.BitLockerCapability.BitLockerAvailable)
OS Version: $($script:BitLockerReport.BitLockerCapability.RequirementsCheck.OSVersion)
UEFI Boot: $($script:BitLockerReport.BitLockerCapability.RequirementsCheck.UEFIBoot)
Secure Boot: $($script:BitLockerReport.BitLockerCapability.RequirementsCheck.SecureBootEnabled)

Drive Encryption Status:
-----------------------
$($script:BitLockerReport.DriveStatus | ForEach-Object { 
    "Drive $($_.MountPoint): $($_.VolumeStatus) - $($_.EncryptionMethod) - Protection: $($_.ProtectionStatus)" + 
    $(if ($_.EncryptionPercentage) { " - Progress: $($_.EncryptionPercentage)%" } else { "" })
} | Out-String)

Recovery Keys:
-------------
$($script:BitLockerReport.RecoveryKeys | ForEach-Object { 
    "Volume $($_.MountPoint): $($_.RecoveryPassword)"
} | Out-String)

Warnings:
--------
$($script:BitLockerReport.Warnings | ForEach-Object { "⚠️ $_" } | Out-String)

Errors:
------
$($script:BitLockerReport.Errors | ForEach-Object { "❌ $_" } | Out-String)

Exit Code: $script:ExitCode
"@
            
            $reportText | Out-File -FilePath $OutputPath -Encoding UTF8
            Write-Host "Report exported to: $OutputPath" -ForegroundColor Green
            
            # Also export as JSON for programmatic use
            $jsonPath = $OutputPath -replace '\.[^.]*$', '.json'
            $script:BitLockerReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
            Write-Host "JSON report exported to: $jsonPath" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to export report: $($_.Exception.Message)"
        }
    }
    
    # Provide recommendations based on findings
    Write-Host "`n=== BitLocker Readiness Assessment ===" -ForegroundColor Cyan
    
    if ($script:ExitCode -eq 1) {
        Write-Host "🔧 Requirements Not Met - Address Before Enabling BitLocker:" -ForegroundColor Yellow
        foreach ($warning in $script:BitLockerReport.Warnings | Where-Object { $_ -like "*TPM*" -or $_ -like "*BitLocker feature*" }) {
            Write-Host "   • $warning" -ForegroundColor Yellow
        }
    }
    
    if ($script:BitLockerReport.Warnings.Count -gt 0) {
        Write-Host "💡 Current Status & Recommendations:" -ForegroundColor Cyan
        foreach ($warning in $script:BitLockerReport.Warnings) {
            Write-Host "   • $warning" -ForegroundColor Gray
        }
    }
    
    if ($script:ExitCode -eq 0) {
        Write-Host "✅ System is ready for BitLocker deployment when needed." -ForegroundColor Green
    }
    
    # Additional recommendations
    Write-Host "`n📋 BitLocker Deployment Preparation:" -ForegroundColor Cyan
    Write-Host "   • Plan recovery key backup strategy (Azure AD, Active Directory, or secure file storage)" -ForegroundColor Gray
    Write-Host "   • Test BitLocker enablement on a pilot group of devices first" -ForegroundColor Gray
    Write-Host "   • Ensure BIOS/UEFI firmware is updated before deployment" -ForegroundColor Gray
    Write-Host "   • Consider enabling Secure Boot for enhanced security" -ForegroundColor Gray
    Write-Host "   • Document BitLocker policies and procedures for your organization" -ForegroundColor Gray
    Write-Host "   • Plan for regular BitLocker status monitoring and compliance checks" -ForegroundColor Gray
    
    Write-Host "`nBitLocker readiness check completed with exit code: $script:ExitCode" -ForegroundColor $(
        switch ($script:ExitCode) {
            0 { "Green" }
            1 { "Yellow" }
            default { "Red" }
        }
    )
    
    # Return the report object for programmatic use
    return $script:BitLockerReport
    
    # Exit with appropriate code
    exit $script:ExitCode
}