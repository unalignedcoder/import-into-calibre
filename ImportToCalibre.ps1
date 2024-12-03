# ============= Developer variables ==============
$log = $true
$verbose = $false
$logFile = Join-Path $PSScriptRoot "ImportIntoCalibre.log"

# Path to Calibre's calibredb.exe
$calibredbPath = "C:\Program Files\Calibre2\calibredb.exe"

# # # Automerge option for calibredb 'add'
# Options: "ignore", "overwrite", "new_record"
# read more about it here:
# https://manual.calibre-ebook.com/generated/en/calibredb.html#id3
$automerge = "overwrite"  # Default behavior

# Automatically kill calibre.exe if running
# Use $false to be prompted
$AutoKillCalibre = $false

# ============= FUNCTIONS ==============

# Function to get the list of Calibre libraries from gui.json
function Get-CalibreLibraries {
    try {
        # Locate the Calibre config directory
        $calibreConfigPath = Join-Path $env:APPDATA "calibre"
        $guiFilePath = Join-Path $calibreConfigPath "gui.json"

        if (-not (Test-Path $guiFilePath)) {
            LogThis "Calibre gui.json not found at $guiFilePath." -verboseMessage $false
            throw "Calibre GUI settings file not found at $guiFilePath."
        }

        # Read and parse gui.json
        $guiJson = Get-Content -Path $guiFilePath -Raw -Encoding UTF8 | ConvertFrom-Json

        # Extract library paths from library_usage_stats
        if (-not $guiJson.library_usage_stats) {
            LogThis "No library usage statistics found in $guiFilePath." -verboseMessage $false
            throw "No library usage statistics found in $guiFilePath."
        }

        $libraries = @()
        foreach ($libraryPath in $guiJson.library_usage_stats.PSObject.Properties.Name) {
            $libraries += [PSCustomObject]@{
                Path = $libraryPath
                Name = Split-Path $libraryPath -Leaf
            }
        }

        if ($libraries.Count -eq 0) {
            LogThis "No libraries found in Calibre GUI settings." -verboseMessage $false
            throw "No libraries found in Calibre GUI settings."
        }

        return $libraries
    } catch {
        LogThis "Error retrieving Calibre libraries: $_" -verboseMessage $false
        throw $_
    }
}



# Check if calibre.exe is running
function Check-CalibreRunning {
    $process = Get-Process -Name "calibre" -ErrorAction SilentlyContinue
    if ($process) {
        return $true
    }
    return $false
}

# Kill calibre.exe process
function Kill-Calibre {
    try {
        Stop-Process -Name "calibre" -Force -ErrorAction Stop
        LogThis "calibre.exe was running and has been terminated."
    } catch {
        LogThis "Failed to terminate calibre.exe: $_"
        throw "Unable to proceed while calibre.exe is running."
    }
}

# Prompt user to terminate calibre.exe
function Prompt-CalibreTermination {
    $response = Read-Host "calibre.exe is running. Would you like to terminate it? (y/n)"
    if ($response -match "^[yY]") {
        Kill-Calibre
    } else {
        throw "calibre.exe must be closed to proceed."
    }
}

# Create the logging system
function LogThis {
    param (
        [string]$message,
        [bool]$verboseMessage = $false
    )

    if ($log) {
        if ($verboseMessage -and -not $verbose) {
            return
        }

        if (IsRunningFromTerminal) {
            Write-Output "$message"
        #} else {
            Add-Content -Path $logFile -Value "$message"
        }
    }
}

# Determine if the script runs interactively
function IsRunningFromTerminal {
    return $true
}

# ============= RUNTIME ==============

try {
    # Check if calibre.exe is running
    if (Check-CalibreRunning) {
        if ($AutoKillCalibre) {
            Kill-Calibre
        } else {
            Prompt-CalibreTermination
        }
    }

    # Ensure at least one ebook file is provided
    if (-not $args) {
        LogThis "Usage: Drag and drop ebook files onto this script or pass them as arguments."
        exit
    }

    # Validate that the calibredb.exe path is correct
    if (-not (Test-Path $calibredbPath)) {
        throw "Could not find calibredb.exe at $calibredbPath. Please update the script with the correct path."
    }

    # Get the list of libraries
    $libraries = Get-CalibreLibraries

    if (-not $libraries) {
        throw "No Calibre libraries found. Ensure Calibre is configured properly."
    }

    # Display available libraries and let the user choose
    LogThis "Available Calibre Libraries:"
    for ($i = 0; $i -lt $libraries.Count; $i++) {
        LogThis "$($i + 1): $($libraries[$i].Name) ($($libraries[$i].Path))"
    }

    # Prompt for library selection
    $libraryIndex = Read-Host "Enter the number of the library to import into (1-$($libraries.Count))"

    if (-not ($libraryIndex -as [int]) -or $libraryIndex -lt 1 -or $libraryIndex -gt $libraries.Count) {
        throw "Invalid selection. Exiting."
    }

    $selectedLibrary = $libraries[$libraryIndex - 1].Path

    # Collect all valid files into an array
    $validFiles = @()

    foreach ($file in $args) {
        if (-not (Test-Path $file)) {
            LogThis "File not found: $file"
        } else {
            $validFiles += $file
        }
    }

    # Check if there are any valid files to import
    if ($validFiles.Count -eq 0) {
        LogThis "No valid files to import. Exiting."
        exit
    }

    # Construct the calibredb add command for all files
    try {
        LogThis "Importing files into $($libraries[$libraryIndex - 1].Name) with automerge=$automerge..."
        $output = & $calibredbPath add --library-path "$selectedLibrary" --automerge $automerge @validFiles 2>&1

        if ($LASTEXITCODE -eq 0) {
            LogThis "Successfully imported files: $($validFiles -join ', ')"
        } else {
            throw "Failed to import files:`n$output"
        }
    } catch {
        LogThis "Error during file import: $_"
    }
} catch {
    LogThis "Critical error: $_"
    exit 1
}
