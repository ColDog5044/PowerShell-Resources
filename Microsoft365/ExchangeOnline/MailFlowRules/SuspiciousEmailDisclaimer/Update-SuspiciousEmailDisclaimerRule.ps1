#Requires -Version 5.1

<#
.SYNOPSIS
    Updates Exchange Online mail flow rule to include regex patterns for detecting suspicious emails.

.DESCRIPTION
    This script connects to Exchange Online and updates the "Apply Disclaimer for Suspicious Emails" transport rule
    with a comprehensive list of regex patterns to detect suspicious email content in both subject lines and message bodies.

.PARAMETER RuleName
    The name of the transport rule to update. Defaults to "Apply Disclaimer for Suspicious Emails".

.INPUTS
    RuleName - The name of the Exchange Online transport rule to update

.OUTPUTS
    Updates Exchange Online transport rule with suspicious email regex patterns

.NOTES
    Version:          1.0
    Author:           Collin Laney
    Creation Date:    2025-08-08

    Exit Codes:
    Code 0 = Success
    Code 1 = Failure
    Code 2 = Script Error

    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Exchange Online administrator permissions

    Credit to SwiftOnSecurity for the suspicious email regex patterns.
    https://github.com/SwiftOnSecurity/PhishingRegex/blob/master/PhishingRegex.txt

.EXAMPLE
    .\Update-SuspiciousEmailDisclaimerRule.ps1
    Runs the script with default parameters.

.EXAMPLE
    .\Update-SuspiciousEmailDisclaimerRule.ps1 -RuleName "Custom Suspicious Email Rule"
    Runs the script with a custom rule name.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/set-transportrule
    https://github.com/SwiftOnSecurity/PhishingRegex/blob/master/PhishingRegex.txt
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$RuleName = "Apply Disclaimer for Suspicious Emails"
)

begin {
    # Initialize variables and functions here
    
    function Get-SuspiciousPatterns {
        return @(
            # SUSPICIOUS PATTERNS - High Priority (URLs and domains)
            "blocked\ your?\ online",
            "suspicious\ activit",
            "updated?\ your\ account\ record",
            "sign\ in\S{0,7}(with\ )?\ your\ email\ address",
            "Verify\ your\ ID\s",
            "dear\ \w{3,8}(\ banking)?\ user",
            "limit\ (and\ suspend\ )?your\ account",
            "sitey\.me",
            "myfreesites\.net",
            "owaportal",
            "outlook\W365",
            "pyapal",
            "pay\Sa\S{0,2}login",
            "%20paypal",
            "\.invoice\.php",
            "security-?err",
            "simplefileupload",
            "security-?warning",
            "-(un)?b?locked",
            "\.my-free\.website",
            "mail-?update",
            "\.yolasite\.com",
            "\.freetemplate\.site",
            "\.ezweb123\.com",
            "\.tripod\.com",
            "\.myfreesites\.net",
            "mailowa",
            "-icloud",
            "icloud-",
            "contabo\.net",
            "\.xyz/",
            "ownership\ validation\ (has\ )?expired",
            "icloudcom",
            "\w\.jar(?=\b)",
            "\.000webhost(app)?\.com",
            "is\.gd/",
            "\.weebly\.com",
            "\.wix\.com",
            "tiny\.cc/",
            "\.joburg",
            "\.top/",
            
            # Additional high-value patterns
            "/GoogleDrive/",
            "/googledocs?/",
            "/Dropfile/",
            "/uploadfile/",
            "/helpdesk/",
            "/natwest/",
            "/dro?pbo?x/",
            "/newdropbox/",
            "/www/amazon",
            "//helpdesk(?!\.)",
            "//webmail(?!\.)",
            "/office\S{0,3}365/",
            "/docu\S{0,3}sign\S{1,4}/",
            "-icloud\Wcom",
            "/https?/www/",
            
            # SUSPICIOUS PHRASES - Critical
            "word must be installed",
            "prevent further unauthorized",
            "prevent further unauthorised",
            "informations has been",
            "fallow our process",
            "confirm your informations",
            "failed to validate",
            "unable to verify",
            "delayed payment",
            "activate your account",
            "Update your payment",
            "submit your payment",
            "via Paypal",
            "has been compromised",
            "FRAUD NOTICE",
            "your account will be closed",
            "your apple id was used to sign in to",
            "was blocked for violation",
            "urged to download",
            "that you validate your account",
            "multiple login attempt",
            "trying to access your account",
            "suspend your account",
            "restricted if you fail to update",
            "informations on your account",
            "update your account information",
            "update in our security",
            "Account Was Limited",
            "verify and reactivate",
            
            # SUSPICIOUS SUBJECTS - Most Common
            "has\ been\ limited",
            "We\ have\ locked",
            "has\ been\ suspended",
            "unusual\ activity",
            "notifications\ pending",
            "your\ (customer\ )?account\ has",
            "your\ (customer\ )?account\ was",
            "new voice(\ )?mail",
            "Periodic\ Maintenance",
            "refund\ not\ approved",
            "account\ (is\ )?on\ hold",
            "wire\ transfer",
            "secure\ update",
            "temporar(il)?y\ deactivated",
            "verification\ required"
        )
    }
    
    function Test-ModuleInstalled {
        param (
            [string]$ModuleName
        )
        
        try {
            $module = Get-Module -ListAvailable -Name $ModuleName
            if ($module) {
                Write-Verbose "$ModuleName module is installed (Version: $($module.Version -join ', '))"
                return $true
            }
            else {
                Write-Verbose "$ModuleName module is not installed"
                return $false
            }
        }
        catch {
            Write-Verbose "Error checking for $ModuleName module: $($_.Exception.Message)"
            return $false
        }
    }
    
    function Install-RequiredModule {
        param (
            [string]$ModuleName
        )
        
        try {
            Write-Host "Installing $ModuleName module..." -ForegroundColor Yellow
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
            Write-Host "$ModuleName module installed successfully" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to install $ModuleName module: $($_.Exception.Message)"
            return $false
        }
    }
    
    function Test-ExchangeOnlineConnection {
        try {
            $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
            if ($session) {
                Write-Verbose "Exchange Online session is already established"
                return $true
            }
            else {
                Write-Verbose "No active Exchange Online session found"
                return $false
            }
        }
        catch {
            Write-Verbose "Error checking Exchange Online connection: $($_.Exception.Message)"
            return $false
        }
    }
}

process {
    try {
        # Check if ExchangeOnlineManagement module is installed
        if (-not (Test-ModuleInstalled -ModuleName "ExchangeOnlineManagement")) {
            Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
            if (-not (Install-RequiredModule -ModuleName "ExchangeOnlineManagement")) {
                Write-Error "Failed to install ExchangeOnlineManagement module. Cannot continue."
                exit 2
            }
        }
        
        # Import the module if not already loaded
        if (-not (Get-Module -Name "ExchangeOnlineManagement")) {
            Write-Host "Importing ExchangeOnlineManagement module..." -ForegroundColor Yellow
            Import-Module ExchangeOnlineManagement -Force
        }

        # Check if Exchange Online connection exists, if not connect
        if (-not (Test-ExchangeOnlineConnection)) {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
            Connect-ExchangeOnline
        }

        # Get suspicious email patterns
        Write-Host "Loading suspicious email patterns..." -ForegroundColor Yellow
        $suspiciousPatterns = Get-SuspiciousPatterns

        Write-Verbose "Loaded $($suspiciousPatterns.Count) suspicious patterns"

        # Update the mail flow rule with both subject and body matching
        Write-Host "Updating transport rule '$RuleName'..." -ForegroundColor Yellow
        Set-TransportRule -Identity $RuleName -SubjectOrBodyMatchesPatterns $suspiciousPatterns

        Write-Host "Successfully updated transport rule with $($suspiciousPatterns.Count) suspicious patterns" -ForegroundColor Green
        exit 0
    }
    catch {
        Write-Error "Script execution failed: $($_.Exception.Message)"
        exit 1
    }
}

end {
    # Clean up code here
    try {
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Script completed successfully" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error during cleanup: $($_.Exception.Message)"
        exit 2
    }
}