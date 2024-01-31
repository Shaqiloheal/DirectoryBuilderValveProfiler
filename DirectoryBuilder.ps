# Define the directory where the Excel files are located
$sourceDirectory = "C:\Valve Profiles"

# Get all Excel files in the source directory
$files = Get-ChildItem -Path $sourceDirectory -Filter "*.xlsx"

foreach ($file in $files) {
    # This regex pattern will match files with the following format:
    # Optional date-time prefix, valve name with optional following letter and digits, optional -C or -O suffix, and _Open or _Close state.
    if ($file.Name -match "(\d{8}-\d{6}_)?(ESV\d+[A-Z]?\d*|HV\d+[A-Z]?\d*|FV\d+[A-Z]?\d*|XV\d+[A-Z]?\d*)(-C|-O)?_(Open|Close)\.xlsx") {
        $valveName = $matches[2]
        $stateSuffix = $matches[3]
        $state = $matches[4]

        # Create the folder name based on valve name and state
        # If there is a -C or -O suffix, it will be appended to the valve name
        $folderName = $valveName
        if ($stateSuffix) {
            $folderName += $stateSuffix
        }
        
        # Add an underscore and the state to the folder name
        $folderName += "_$state"

        $folderPath = Join-Path -Path $sourceDirectory -ChildPath $folderName

        # Check if the folder exists, if not, create it
        if (-not (Test-Path -Path $folderPath)) {
            New-Item -ItemType Directory -Path $folderPath
        }

        # Move the file to the new folder
        $newFilePath = Join-Path -Path $folderPath -ChildPath $file.Name
        Move-Item -Path $file.FullName -Destination $newFilePath
    } else {
        Write-Host "File name format is not recognized: $($file.Name)"
    }
}

Write-Host "Script completed."
