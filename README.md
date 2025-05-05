# Office Audit PowerShell Script

## Overview
The `OfficeAuditv3.ps1` script is a comprehensive tool designed to audit Microsoft Office installations on Windows systems. Created by Jesus Ayala from Sarah Lawrence College, this script identifies and reports key information about Office installations including:

- Installation type (Click-to-Run or MSI-based)
- Architecture (32-bit or 64-bit)
- Product name
- Version details (including build numbers)
- Marketing version (e.g., Office 2016, 2019, 2021, 2024)

## Features

- **Detailed Reporting**: Provides comprehensive information about Office installations in an easy-to-read format
- **Multiple Detection Methods**: Uses several approaches to detect Office installations:
  - Registry queries for Click-to-Run installations
  - WMI queries for MSI-based installations
  - File system checks for additional verification
- **Version Mapping**: Automatically maps build numbers to marketing versions (2016, 2019, 2021, 2024)
- **System Context**: Includes computer name, current user, and OS version information

## Technical Details

### Functions

#### `Get-InteractiveUser`
Retrieves the current interactive user logged into the system, stripping domain information if present.

### Detection Logic

The script follows a hierarchical detection approach:

1. **Click-to-Run Detection**:
   - Checks registry paths for Click-to-Run configuration
   - Extracts platform, version, and product information
   - Captures full version string and build number

2. **MSI-based Office Detection** (if Click-to-Run not found):
   - Uses WMI to query installed products
   - Identifies Microsoft Office products (excluding Click-to-Run variants)
   - Captures version and name information

3. **Binary Architecture Detection**:
   - Scans program directories for Office executables
   - Determines architecture based on installation path

4. **Version Mapping**:
   - Maps build numbers to marketing versions using the following logic:
     - Build ≥ 17000: Office 2024
     - Build ≥ 14000: Office 2021
     - Build ≥ 10300: Office 2019
     - Build ≥ 4266: Office 2016
     - Also detects Office 2013 (15.0) and Office 2010 (14.0)

### Output

The script produces a formatted list containing:
- Computer name
- Username
- OS version
- Office type (Click-to-Run or MSI)
- Office architecture (32-bit or 64-bit)
- Office product name
- Marketing version (e.g., 2016, 2019, 2021)
- Build number
- Full version string
- Timestamp of audit

## Usage

1. Run the script in PowerShell:
   ```powershell
   .\OfficeAuditv3.ps1
   ```

2. Review the detailed output displayed in the console.

## Requirements

- Windows operating system
- PowerShell 3.0 or higher
- Administrative privileges (recommended for complete detection)

## Example Output

```
Office Audit Summary:

ComputerName      : DESKTOP-ABCD123
Username          : jdoe
OSVersion         : Microsoft Windows 11 Pro
OfficeType        : Click-to-Run
OfficeArch        : x64
OfficeName        : O365ProPlus
OfficeVersion     : 2021
OfficeBuildNumber : 14931
OfficeFullVersion : 16.0.14931.20132
Timestamp         : 2025-05-05 10:15:30
```

## Notes

- The script handles both Click-to-Run and traditional MSI-based Office installations
- If an Office installation is not detected, the script will display "None" as the OfficeType
- Build numbers are continuously updated with new Office releases

## Author
Created by Jesus Ayala from Sarah Lawrence College
