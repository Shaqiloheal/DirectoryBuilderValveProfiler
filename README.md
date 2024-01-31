# Valve Profile Directory Organizer

This PowerShell script is designed to organize Excel files containing valve profiles into designated directories for various subsea tree projects.

## Overview

The script performs the following actions:

1. Defines a source directory where the Excel files are located.
2. Filters and retrieves all Excel files within the source directory.
3. Matches each file name against a specified pattern to identify the valve name, state suffix, and valve state (Open or Close).
4. Creates a new directory named after the valve and its state if it doesn't exist.
5. Moves each Excel file into its corresponding directory.

## File Naming Convention

The script expects Excel files to follow a naming convention as outlined below:

- Optional date-time prefix in the format `YYYYMMDD-HHMMSS_`
- Valve name which may include:
  - A standard prefix (`ESV`, `HV`, `FV`, `XV`)
  - A sequence of digits
  - An optional letter
  - Additional optional digits
- An optional state suffix (`-C` or `-O`)
- A mandatory state indicator (`_Open` or `_Close`)
- The file extension `.xlsx`

Example: `20180610-145906_FV80414_Open.xlsx`

## How to Use

To use the script, follow these steps:

1. Ensure that PowerShell is installed on your system.
2. Save the script to a `.ps1` file, for instance, `DirectoryBuilder.ps1`.
3. Modify the `$sourceDirectory` variable in the script to point to your actual directory containing the valve profile Excel files.
4. Open PowerShell and navigate to the directory where the script is saved.
5. Execute the script by typing `.\DirectoryBuilder.ps1` and pressing Enter.

## Output

The script will:

- Create subdirectories within the source directory following the naming convention `<ValveName>_<State>`.
- Move each Excel file into its corresponding directory.
- Print a message for files that do not follow the naming convention and are not moved.

## Notes

- It is essential that the Excel files adhere to the expected naming convention for the script to process them correctly.
- Files that do not match the pattern will be reported in the PowerShell console and will not be moved.
- The script will output "Script completed." once it has finished running.

For any queries or issues, please refer to the script documentation or contact the support team.
