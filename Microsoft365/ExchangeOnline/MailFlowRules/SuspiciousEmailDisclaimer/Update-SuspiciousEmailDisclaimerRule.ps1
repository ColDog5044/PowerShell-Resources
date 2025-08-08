#Requires -Version 5.1

<#
.SYNOPSIS
    Updates Exchange Online mail flow rule to include regex patterns for detecting suspicious emails.

.DESCRIPTION
    This script connects to Exchange Online and updates the "Apply Disclaimer for Suspicious Emails" transport rule
    with a comprehensive list of regex patterns to detect suspicious email content in both subject lines and message bodies.

.PARAMETER RuleName
    The name of the transport rule to update. Defaults to "Apply Disclaimer for Suspicious Emails".

.PARAMETER CreateIfNotExists
    Switch parameter. If specified, the script will create a new transport rule if one doesn't exist.
    The new rule will include the suspicious patterns and a default HTML disclaimer.

.INPUTS
    RuleName - The name of the Exchange Online transport rule to update
    CreateIfNotExists - Switch to enable automatic rule creation

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
    Runs the script with default parameters to update an existing rule.

.EXAMPLE
    .\Update-SuspiciousEmailDisclaimerRule.ps1 -RuleName "Custom Suspicious Email Rule"
    Runs the script with a custom rule name.

.EXAMPLE
    .\Update-SuspiciousEmailDisclaimerRule.ps1 -CreateIfNotExists
    Creates a new rule if it doesn't exist, otherwise updates the existing rule.

.EXAMPLE
    .\Update-SuspiciousEmailDisclaimerRule.ps1 -RuleName "My Phishing Rule" -CreateIfNotExists
    Creates or updates a rule with a custom name.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/set-transportrule
    https://github.com/SwiftOnSecurity/PhishingRegex/blob/master/PhishingRegex.txt
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$RuleName = "Apply Disclaimer for Suspicious Emails",
    
    [Parameter(Mandatory = $false)]
    [switch]$CreateIfNotExists
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
    
    function Test-ExchangeOnlineConnection {
        try {
            $session = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
            if ($session) {
                Write-Verbose "Exchange Online connection found"
                return $true
            }
            else {
                Write-Verbose "No active Exchange Online connection found"
                return $false
            }
        }
        catch {
            Write-Verbose "Error checking Exchange Online connection: $($_.Exception.Message)"
            return $false
        }
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
    
    function Get-DefaultDisclaimerHtml {
        return @"
<div style="
		background-color: #f0fdfa;
		border: 1pt solid #14b8a6;
		border-left: 6pt solid #0f766e;
		border-radius: 4pt;
		margin: 10pt 0;
		padding: 12pt;
		font-family: Verdana, sans-serif;">
	<div style="margin-bottom: 8pt; color: #0f766e; font-weight: bold; font-size: 12pt">SUSPICIOUS EMAIL</div>
	<div style="font-size: 10pt; color: #374151; line-height: 1.4; margin-bottom: 8pt">
		This email has been flagged for containing potentially suspicious patterns, including phishing keywords, urgent language, or
		questionable attachments.
	</div>
	<div style="background-color: #ccfbf1; padding: 8pt; border-radius: 3pt; font-size: 9pt; color: #0f766e">
		<strong>RECOMMENDED ACTIONS:</strong>
		<br />• Do not click links or download attachments without verification <br />• Independently verify sender identity through known
		contact methods <br />• Report to the IT Department if you believe this is a legitimate email
	</div>
</div>
"@
    }

    function Test-TransportRuleExists {
        param (
            [string]$RuleName
        )
        
        try {
            $rule = Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue
            if ($rule) {
                Write-Verbose "Transport rule '$RuleName' found"
                return $true
            }
            else {
                Write-Verbose "Transport rule '$RuleName' not found"
                return $false
            }
        }
        catch {
            Write-Verbose "Error checking for transport rule '$RuleName': $($_.Exception.Message)"
            return $false
        }
    }

    function New-SuspiciousEmailRule {
        param (
            [string]$RuleName,
            [string[]]$SuspiciousPatterns
        )
        
        try {
            $disclaimerHtml = Get-DefaultDisclaimerHtml
            
            Write-Host "Creating new transport rule '$RuleName'..." -ForegroundColor Yellow
            
            New-TransportRule -Name $RuleName `
                -SubjectOrBodyMatchesPatterns $SuspiciousPatterns `
                -ApplyHtmlDisclaimerText $disclaimerHtml `
                -ApplyHtmlDisclaimerLocation Prepend `
                -ApplyHtmlDisclaimerFallbackAction Wrap `
                -Comments "Automatically created by Update-SuspiciousEmailDisclaimerRule.ps1 script to detect suspicious email patterns and apply warning disclaimers."
            
            Write-Host "Successfully created transport rule '$RuleName'" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to create transport rule '$RuleName': $($_.Exception.Message)"
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

        # Check if transport rule exists
        if (-not (Test-TransportRuleExists -RuleName $RuleName)) {
            if ($CreateIfNotExists) {
                # Create new rule with patterns and disclaimer
                if (New-SuspiciousEmailRule -RuleName $RuleName -SuspiciousPatterns $suspiciousPatterns) {
                    Write-Host "Successfully created and configured transport rule with $($suspiciousPatterns.Count) suspicious patterns" -ForegroundColor Green
                    exit 0
                } else {
                    Write-Error "Failed to create transport rule '$RuleName'. Cannot continue."
                    exit 1
                }
            } else {
                Write-Error "Transport rule '$RuleName' does not exist. Use -CreateIfNotExists parameter to create it automatically."
                exit 1
            }
        }

        # Update existing mail flow rule with patterns
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