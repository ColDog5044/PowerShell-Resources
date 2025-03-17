# New-ScriptTemplate.ps1

## Overview

`New-ScriptTemplate.ps1` is a PowerShell script template designed to provide a structured starting point for creating new scripts. It includes sections for metadata, parameters, inputs, outputs, examples, and links to documentation.

## Features

-   Includes a `Test-IsElevated` function to check if the script is running with elevated permissions.
-   Provides a structured template for adding metadata, parameters, and examples.
-   Supports PowerShell version 5.1 or later.

## Requirements

-   **PowerShell Version**: 5.1 or later
-   **Permissions**: Must be run with elevated permissions (Administrator)

## Parameters

| Name         | Type   | Mandatory | Description              |
| ------------ | ------ | --------- | ------------------------ |
| `parameter1` | String | No        | A placeholder parameter. |

## Exit Codes

| Code | Description                       |
| ---- | --------------------------------- |
| 0    | Success                           |
| 1    | Failure                           |
| 2    | Script Error (e.g., not elevated) |

## Usage

### Example

```powershell
.\New-ScriptTemplate.ps1
```

Runs the script with no parameters.

## Functions

### Test-IsElevated

Checks if the script is running with elevated permissions. If not, it exits with code `2`.

## Notes

-   Version 1.0
-   Author Your Name
-   Creation Date `<Date>`
