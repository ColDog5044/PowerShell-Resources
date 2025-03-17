#Requires -Version 5.1

<#
.SYNOPSIS
    A brief description of the script or function.

.DESCRIPTION
    A detailed description of the script or function.

.PARAMETER ParameterName
    A description of the parameter.

.INPUTS
    <Inputs if any, otherwise state None>

.OUTPUTS
    <Outputs if any, otherwise state None>

.NOTES
    Version:          1.0
    Author:           Your Name
    Creation Date:    <Date>

    Exit Codes:
    Code 0 = Success
    Code 1 = Failure
    Code 2 = Script Error

.EXAMPLE
    .\New-ScriptTemplate.ps1
    Runs the script with no parameters.

.LINK
    https://www.example.com/help/New-ScriptTemplate
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$parameter1
)

begin {
    # Initialize variables and functions here
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

    
}

process {
    # Main code here
    Test-IsElevated


}

end {
    # Clean up code here


}
