#Requires -Version 5.1

<#
.SYNOPSIS
    A brief description of the script or function.

.DESCRIPTION
    A detailed description of the script or function.

.PARAMETER ParameterName
    A description of the parameter.

.INPUTS
    None

.OUTPUTS
    [string] The name of the entity.

.NOTES
    Author: Your Name
    Date:   The date
    Version: 1.0


.EXAMPLE
    .\New-ScriptTemplate.ps1
    Runs the script with no parameters.

.LINK
    https://www.example.com/help/New-ScriptTemplate
#>

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