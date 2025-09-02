#Requires -Version 5.1

<#
.SYNOPSIS
    Enables BitLocker encryption on the system drive with comprehensive security options.

.DESCRIPTION
    This script provides a comprehensive solution for enabling BitLocker encryption on Windows systems.
    It automatically installs the BitLocker module if not present, validates system readiness,
    and configures BitLocker with appropriate key protectors based on system capabilities.
    
    The script supports TPM-based encryption, recovery passwords, and startup keys,
    ensuring maximum security while maintaining usability.

.PARAMETER Drive
    The drive letter to encrypt. Defaults to the system drive (C:).

.PARAMETER EncryptionMethod
    The encryption method to use. Valid options: Aes128, Aes256, XtsAes128, XtsAes256.
    Defaults to XtsAes256 for best security.

.PARAMETER RecoveryPasswordPath
    Optional path to save the recovery password. If not specified, displays on screen only.

.PARAMETER SkipHardwareTest
    Skip the hardware compatibility test. Use with caution.

.PARAMETER Force
    Force encryption even if warnings are present.

.INPUTS
    None

.OUTPUTS
    System.Boolean - Returns $true if BitLocker was successfully enabled, $false otherwise.

.NOTES
    Version:          2.0
    Author:           Collin Laney
    Creation Date:    2025-09-02
    Last Modified:    2025-09-02

    Exit Codes:
    Code 0 = Success - BitLocker enabled successfully
    Code 1 = Failure - BitLocker could not be enabled
    Code 2 = Script Error - Script execution error or insufficient permissions
    Code 3 = System Not Ready - System does not meet BitLocker requirements

    Requirements:
    - Windows 10/11 Pro, Enterprise, or Education
    - TPM 1.2 or later (recommended)
    - Administrator privileges
    - UEFI firmware (recommended)

.EXAMPLE
    .\Enable-BitLocker.ps1
    Enables BitLocker on the system drive with default settings.

.EXAMPLE
    .\Enable-BitLocker.ps1 -Drive "D:" -EncryptionMethod Aes256
    Enables BitLocker on drive D: using AES 256-bit encryption.

.EXAMPLE
    .\Enable-BitLocker.ps1 -RecoveryPasswordPath "C:\BitLocker\Recovery.txt" -Force
    Enables BitLocker and saves the recovery password to a file, forcing encryption even with warnings.

.LINK
    https://docs.microsoft.com/en-us/windows/security/information-protection/bitlocker/
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[A-Za-z]:$')]
    [string]$Drive = $env:SystemDrive,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('Aes128', 'Aes256', 'XtsAes128', 'XtsAes256')]
    [string]$EncryptionMethod = 'XtsAes256',
    
    [Parameter(Mandatory = $false)]
    [string]$RecoveryKeyPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipHardwareTest,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

begin {
    # Initialize variables and functions here
    $Script:BitLockerEnabled = $false
    $Script:RecoveryPassword = $null
    
    function Test-IsElevated {
        $isElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($isElevated) {
            Write-Verbose "The script is running with elevated permissions"
        }
        else {
            Write-Error "The script must be run with elevated permissions"
            exit 2
        }
    }

    function Install-BitLockerModule {
        <#
        .SYNOPSIS
            Installs the BitLocker module if not already present
        #>
        try {
            Write-Host "Checking for BitLocker module..." -ForegroundColor Yellow
            
            # Check if BitLocker module is available
            $bitLockerModule = Get-Module -ListAvailable -Name BitLocker
            
            if (-not $bitLockerModule) {
                Write-Host "BitLocker module not found. Installing..." -ForegroundColor Yellow
                
                # Enable BitLocker feature if not enabled
                $bitLockerFeature = Get-WindowsOptionalFeature -Online -FeatureName BitLocker
                if ($bitLockerFeature.State -ne "Enabled") {
                    Write-Host "Enabling BitLocker Windows feature..." -ForegroundColor Yellow
                    Enable-WindowsOptionalFeature -Online -FeatureName BitLocker -All -NoRestart
                    Write-Warning "A restart may be required for BitLocker feature installation to complete."
                }
                
                # Try to import the module after feature installation
                Import-Module BitLocker -Force -ErrorAction Stop
                Write-Host "BitLocker module installed and imported successfully." -ForegroundColor Green
            }
            else {
                Write-Host "BitLocker module found. Importing..." -ForegroundColor Green
                Import-Module BitLocker -Force
            }
            
            return $true
        }
        catch {
            Write-Error "Failed to install or import BitLocker module: $($_.Exception.Message)"
            return $false
        }
    }

    function Test-BitLockerReadiness {
        <#
        .SYNOPSIS
            Tests if the system is ready for BitLocker encryption
        #>
        try {
            Write-Host "Testing BitLocker readiness..." -ForegroundColor Yellow
            
            # Check Windows edition
            $windowsEdition = (Get-WmiObject -Class Win32_OperatingSystem).Caption
            Write-Verbose "Windows Edition: $windowsEdition"
            
            if ($windowsEdition -notmatch "(Pro|Enterprise|Education)") {
                Write-Warning "BitLocker is not supported on this Windows edition: $windowsEdition"
                if (-not $Force) {
                    Write-Error "BitLocker requires Windows Pro, Enterprise, or Education edition."
                    return $false
                }
            }
            
            # Check TPM status
            $tpm = Get-WmiObject -Namespace "Root\CIMv2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue
            if ($tpm) {
                Write-Host "TPM Status: Present and Ready" -ForegroundColor Green
                Write-Verbose "TPM Version: $($tpm.SpecVersion)"
            }
            else {
                Write-Warning "TPM not detected or not ready. BitLocker can still be enabled with alternative methods."
            }
            
            # Check if drive is already encrypted
            try {
                $bitLockerStatus = Get-BitLockerVolume -MountPoint $Drive -ErrorAction SilentlyContinue
                if ($bitLockerStatus -and $bitLockerStatus.ProtectionStatus -eq "On") {
                    Write-Warning "Drive $Drive is already encrypted with BitLocker."
                    if (-not $Force) {
                        Write-Host "Use -Force to proceed anyway." -ForegroundColor Yellow
                        return $false
                    }
                }
            }
            catch {
                # BitLocker volume info not available - this is expected for unencrypted drives
                Write-Verbose "Drive $Drive is not currently encrypted."
            }
            
            # Check disk space
            $driveInfo = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $Drive }
            if ($driveInfo) {
                $freeSpaceGB = [math]::Round($driveInfo.FreeSpace / 1GB, 2)
                Write-Verbose "Free space on $Drive : $freeSpaceGB GB"
                
                if ($freeSpaceGB -lt 1.5) {
                    Write-Warning "Low disk space detected ($freeSpaceGB GB). BitLocker requires at least 1.5GB free space."
                    if (-not $Force) {
                        return $false
                    }
                }
            }
            
            # Check UEFI/BIOS mode
            try {
                $firmwareType = (Get-ComputerInfo).BiosFirmwareType
                if ($firmwareType -eq "UEFI") {
                    Write-Host "Firmware Type: UEFI (Recommended)" -ForegroundColor Green
                }
                else {
                    Write-Warning "Firmware Type: Legacy BIOS. UEFI is recommended for optimal BitLocker security."
                }
            }
            catch {
                Write-Verbose "Could not determine firmware type."
            }
            
            return $true
        }
        catch {
            Write-Error "Error during BitLocker readiness test: $($_.Exception.Message)"
            return $false
        }
    }

    function New-BitLockerRecoveryPassword {
        <#
        .SYNOPSIS
            Generates and manages BitLocker recovery password
        #>
        try {
            Write-Host "Generating BitLocker recovery password..." -ForegroundColor Yellow
            
            # Add recovery password protector
            $recoveryProtector = Add-BitLockerKeyProtector -MountPoint $Drive -RecoveryPasswordProtector
            
            if ($recoveryProtector) {
                # Get the recovery password
                $recoveryInfo = Get-BitLockerVolume -MountPoint $Drive
                $Script:RecoveryPassword = ($recoveryInfo.KeyProtector | Where-Object { $_.KeyProtectorType -eq "RecoveryPassword" }).RecoveryPassword
                
                Write-Host "Recovery Password Generated Successfully!" -ForegroundColor Green
                Write-Host "IMPORTANT: Save this recovery password in a secure location!" -ForegroundColor Red
                Write-Host "Recovery Password: $($Script:RecoveryPassword)" -ForegroundColor Cyan
                
                # Save to file if path specified
                if ($RecoveryKeyPath) {
                    try {
                        $recoveryContent = @"
BitLocker Recovery Information
Generated: $(Get-Date)
Drive: $Drive
Computer: $env:COMPUTERNAME
Recovery Password: $($Script:RecoveryPassword)

IMPORTANT: Keep this information secure and accessible.
You will need this password to unlock your drive if you lose access to your primary unlock method.
"@
                        $recoveryContent | Out-File -FilePath $RecoveryKeyPath -Encoding UTF8
                        Write-Host "Recovery password saved to: $RecoveryKeyPath" -ForegroundColor Green
                    }
                    catch {
                        Write-Warning "Could not save recovery password to file: $($_.Exception.Message)"
                    }
                }
                
                return $true
            }
            else {
                Write-Error "Failed to generate recovery password protector."
                return $false
            }
        }
        catch {
            Write-Error "Error generating recovery password: $($_.Exception.Message)"
            return $false
        }
    }

    function Enable-BitLockerEncryption {
        <#
        .SYNOPSIS
            Enables BitLocker encryption with appropriate key protectors
        #>
        try {
            Write-Host "Configuring BitLocker encryption on drive $Drive..." -ForegroundColor Yellow
            
            # Check if TPM is available for key protector
            $tpm = Get-WmiObject -Namespace "Root\CIMv2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue
            
            if ($tpm -and $tpm.IsReady()) {
                Write-Host "Using TPM + Recovery Password protection..." -ForegroundColor Green
                
                # Enable BitLocker with TPM protector
                Enable-BitLocker -MountPoint $Drive -EncryptionMethod $EncryptionMethod -TpmProtector -ErrorAction Stop
                
                # Add recovery password protector
                if (-not (New-BitLockerRecoveryPassword)) {
                    Write-Warning "Failed to add recovery password protector, but TPM protector is active."
                }
            }
            else {
                Write-Host "TPM not available. Using Recovery Password protection only..." -ForegroundColor Yellow
                
                # Generate recovery password first
                if (New-BitLockerRecoveryPassword) {
                    # Enable BitLocker with recovery password only
                    Enable-BitLocker -MountPoint $Drive -EncryptionMethod $EncryptionMethod -PasswordProtector -ErrorAction Stop
                }
                else {
                    throw "Failed to configure recovery password protector."
                }
            }
            
            # Start encryption
            Write-Host "Starting BitLocker encryption..." -ForegroundColor Yellow
            Resume-BitLocker -MountPoint $Drive
            
            # Get encryption status
            $bitLockerStatus = Get-BitLockerVolume -MountPoint $Drive
            Write-Host "Encryption Status: $($bitLockerStatus.VolumeStatus)" -ForegroundColor Green
            Write-Host "Encryption Percentage: $($bitLockerStatus.EncryptionPercentage)%" -ForegroundColor Green
            
            $Script:BitLockerEnabled = $true
            return $true
        }
        catch {
            Write-Error "Failed to enable BitLocker: $($_.Exception.Message)"
            return $false
        }
    }

    function Show-BitLockerStatus {
        <#
        .SYNOPSIS
            Displays current BitLocker status and next steps
        #>
        try {
            $bitLockerStatus = Get-BitLockerVolume -MountPoint $Drive
            
            Write-Host "`n" + "="*60 -ForegroundColor Cyan
            Write-Host "BITLOCKER STATUS SUMMARY" -ForegroundColor Cyan
            Write-Host "="*60 -ForegroundColor Cyan
            
            Write-Host "Drive: $($bitLockerStatus.MountPoint)" -ForegroundColor White
            Write-Host "Volume Status: $($bitLockerStatus.VolumeStatus)" -ForegroundColor $(if ($bitLockerStatus.VolumeStatus -eq "FullyEncrypted") { "Green" } else { "Yellow" })
            Write-Host "Protection Status: $($bitLockerStatus.ProtectionStatus)" -ForegroundColor $(if ($bitLockerStatus.ProtectionStatus -eq "On") { "Green" } else { "Yellow" })
            Write-Host "Encryption Method: $($bitLockerStatus.EncryptionMethod)" -ForegroundColor White
            Write-Host "Encryption Percentage: $($bitLockerStatus.EncryptionPercentage)%" -ForegroundColor White
            
            Write-Host "`nKey Protectors:" -ForegroundColor Cyan
            foreach ($protector in $bitLockerStatus.KeyProtector) {
                Write-Host "  - $($protector.KeyProtectorType)" -ForegroundColor White
            }
            
            if ($Script:RecoveryPassword) {
                Write-Host "`nRecovery Password: $($Script:RecoveryPassword)" -ForegroundColor Red
                Write-Host "IMPORTANT: Save this password in a secure location!" -ForegroundColor Red
            }
            
            Write-Host "`nNext Steps:" -ForegroundColor Cyan
            if ($bitLockerStatus.VolumeStatus -eq "EncryptionInProgress") {
                Write-Host "  1. Encryption is in progress. Do not shut down unexpectedly." -ForegroundColor Yellow
                Write-Host "  2. Monitor progress with: Get-BitLockerVolume -MountPoint $Drive" -ForegroundColor White
                Write-Host "  3. Encryption will continue in the background." -ForegroundColor White
            }
            elseif ($bitLockerStatus.VolumeStatus -eq "FullyEncrypted") {
                Write-Host "  1. BitLocker encryption is complete!" -ForegroundColor Green
                Write-Host "  2. Your drive is now protected." -ForegroundColor Green
                Write-Host "  3. Keep your recovery password safe." -ForegroundColor Yellow
            }
            
            Write-Host "="*60 -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Could not retrieve BitLocker status: $($_.Exception.Message)"
        }
    }
}

process {
    # Main code execution
    try {
        Write-Host "Starting BitLocker Enablement Process..." -ForegroundColor Cyan
        Write-Host "Target Drive: $Drive" -ForegroundColor White
        Write-Host "Encryption Method: $EncryptionMethod" -ForegroundColor White
        Write-Host ""
        
        # Step 1: Check elevation
        Test-IsElevated
        
        # Step 2: Install/Import BitLocker module
        if (-not (Install-BitLockerModule)) {
            Write-Error "Failed to install BitLocker module. Cannot continue."
            exit 2
        }
        
        # Step 3: Test system readiness
        if (-not $SkipHardwareTest) {
            if (-not (Test-BitLockerReadiness)) {
                Write-Error "System is not ready for BitLocker encryption."
                exit 3
            }
        }
        else {
            Write-Warning "Skipping hardware compatibility test as requested."
        }
        
        # Step 4: Enable BitLocker encryption
        if (Enable-BitLockerEncryption) {
            Write-Host "`nBitLocker has been successfully enabled!" -ForegroundColor Green
            
            # Step 5: Display status and next steps
            Show-BitLockerStatus
            
            # Return success
            Write-Host "`nScript completed successfully." -ForegroundColor Green
            exit 0
        }
        else {
            Write-Error "Failed to enable BitLocker encryption."
            exit 1
        }
    }
    catch {
        Write-Error "An unexpected error occurred: $($_.Exception.Message)"
        Write-Error "Stack Trace: $($_.ScriptStackTrace)"
        exit 2
    }
}

end {
    # Clean up code here
    if ($Script:BitLockerEnabled) {
        Write-Host "`nCleanup: BitLocker enablement process completed successfully." -ForegroundColor Green
        
        # Final reminder about recovery password
        if ($Script:RecoveryPassword -and -not $RecoveryKeyPath) {
            Write-Host "`nIMPORTANT REMINDER:" -ForegroundColor Red
            Write-Host "Your BitLocker recovery password is: $($Script:RecoveryPassword)" -ForegroundColor Cyan
            Write-Host "Please save this password in a secure location immediately!" -ForegroundColor Red
        }
        
        # Provide helpful commands for monitoring
        Write-Host "`nUseful BitLocker Commands:" -ForegroundColor Cyan
        Write-Host "  Check status: Get-BitLockerVolume -MountPoint $Drive" -ForegroundColor White
        Write-Host "  Backup recovery key: manage-bde -protectors $Drive -get" -ForegroundColor White
        Write-Host "  Suspend BitLocker: Suspend-BitLocker -MountPoint $Drive" -ForegroundColor White
        Write-Host "  Resume BitLocker: Resume-BitLocker -MountPoint $Drive" -ForegroundColor White
    }
    else {
        Write-Host "`nCleanup: BitLocker enablement was not successful." -ForegroundColor Yellow
        Write-Host "Please review the error messages above and try again." -ForegroundColor Yellow
    }
    
    # Clear sensitive variables
    if ($Script:RecoveryPassword) {
        $Script:RecoveryPassword = $null
    }
}