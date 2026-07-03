#Requires -Version 5.1

<#
.SYNOPSIS
A PowerShell GUI tool to compare source files/folders/archives against target media folders,
extract matching archives, rename specific image files, and move processed images.

.DESCRIPTION
This tool provides a graphical interface to manage media-related image files.
- Compares items in a source folder against 'Show Name (YYYY)' folders in target locations.
- Extracts matching .zip, .rar, or .7z archives in the source folder using 7-Zip. (Button 1)
- Renames backdrop images ('* - Backdrop') found within source subfolders to 'backdrop.ext'. (Button 2)
- Moves loose backdrop images ('Show Name (YYYY) - Backdrop.ext') found directly in the source folder to the target show folder, renaming them appropriately (overwrites existing). (Button 2)
- Renames season poster images ('Show Name (YYYY) - Season X.ext') found within source subfolders to 'seasonXX-poster.ext' or 'season-specials-poster.ext' (overwrites existing). (Button 3)
- Renames show poster images ('Show Name (YYYY).ext') found within source subfolders to 'folder.ext' (overwrites existing). (Button 3)
- Moves loose season poster images ('Show Name (YYYY) - Season X.ext') found directly in the source folder to the target show folder, renaming them appropriately (overwrites existing). (Button 3)
- Moves loose show poster images ('Show Name (YYYY).ext') found directly in the source folder to the target show folder, renaming them to 'folder.ext' (overwrites existing). (Button 3)
- Renames episode thumbnail images ('*SXXEXX*') found within source subfolders to match corresponding video filenames ('video_filename-thumb.ext') (overwrites existing). (Button 4)
- Moves processed images from source subfolders (thumbs, season posters, backdrops, folder images) to the appropriate target show/season folders (overwrites existing). (Button 5)
- Allows users to select source and target folders, which are saved for future use.
- Provides individual buttons for each step and a "Run All" button.

.NOTES
Date:       2025-04-20
Updated:    2026-07-02
Requires:   PowerShell 5.1+, .NET Framework 4.5+, 7-Zip (7z.exe must be in system PATH).
Backup your data before extensive use! Overwriting is enabled.

--- Changelog ---

2026-07-02 — v1.0 (Project renamed: Shipwreck)
- Rebranded project as "Shipwreck". In the real world, shipwrecks act as time capsules,
  preserving history at the bottom of the ocean. This tool is about preservation — finding,
  downloading, and locally saving high-quality artwork so your library looks perfect even
  if online metadata sites ever go down.
- Title bar updated to "Shipwreck v1.0".
- Config file renamed from MediaToolConfig.json to manifest.json.
- Log file saved as shipwreck.log.
- All log messages updated with nautical theme: diving/scouting for shows, salvaging archives,
  cataloguing renames, hauling cargo, lost cargo for errors, emptying the hold for cleanup.
- Straggler detection: files that remain in the source folder after a move step (no matching
  video found) are flagged as "adrift at sea" in the log.

2026-06-28
- Fix: Episode thumbnails now route to the exact season folder the video file lives in,
  rather than scanning folder names and potentially matching the wrong folder (e.g. routing
  to a stale "Season 01" instead of the correct "Season 1"). Rename-EpisodeImages now records
  a thumb filename -> video directory map ($script:episodeTargetDirMap) which Move-ImageFilesToServer
  consults directly. The old season folder name scan is kept as a fallback for standalone use.

- Fix: Three-tier cascading match (Find-TargetShowMatch) replaces all direct hashtable lookups
  for show name matching. Tier 1: exact. Tier 2: Mediux spacing fix ("Name- Title" -> "Name - Title").
  Tier 3: loose match strips hyphens and colons from both sides before comparing, handling cases
  like "Fullmetal Alchemist- Brotherhood (2009)" vs "Fullmetal Alchemist Brotherhood (2009)".
  Write-MatchLog logs which tier fired (cyan for normalized, gold for loose) so every fallback
  match is visible and verifiable in the log.

- Refactor: Process-Action now builds the target show folder map once before any step runs
  and passes it to all five core functions via an optional $PrebuiltMap parameter (each function
  falls back to building locally if not provided). Trim-ImageFilenamesInFolder also moved into
  Process-Action and runs once per operation rather than redundantly inside individual buttons.
  Scan progress is logged at startup so the UI never appears frozen.

- Fix: S00/Specials episode thumbnails were never routed to the Specials folder because
  Move-ImageFilesToServer assigned the Specials path to $script:targetSpecialsPath (script scope)
  but all downstream checks referenced $targetSpecialsPath (local scope). Fixed to local scope.

- Feature: Trim-ImageFilenamesInFolder extended to also fix Mediux hyphen spacing in source
  filenames at the physical file level (e.g. "Dragon Ball Z- Fusion Reborn (1995).jpg" renamed
  to "Dragon Ball Z - Fusion Reborn (1995).jpg") in addition to trimming whitespace. Output
  switched from Write-Host to Add-LogEntry so changes appear in the GUI log. Silent for files
  that need no changes.

- Feature: Get-NormalizedShowName helper added. Corrects Mediux's missing-space-before-hyphen
  pattern at comparison time as a safety net for files inside subfolders that bypass the trim step.

- GUI: Default window size increased to 1400x900, minimum to 900x650. Button dimensions
  increased (170x30 -> 240x40), target listbox height increased, log textbox initial size
  increased so it fills the window properly. Form.Activate() added to Load event to bring
  the window to the foreground when launched from a console.
#>

# --- Configuration ---
# WPF assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
# Keep WinForms loaded solely for FolderBrowserDialog (no native WPF folder picker)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Path for storing configuration (Source/Target folders)
# Uses the script's directory for the config file, or PWD if run interactively.
$scriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptPath)) {
    # Fallback to current working directory if script path isn't available (e.g., running code directly in ISE/console)
    Write-Warning "Cannot determine script path. Using current working directory ($PWD) for config file."
    $scriptDir = $PWD.Path
} else {
    # Get the directory containing the running script
    $scriptDir = Split-Path -Parent $scriptPath
}
# *** CONFIG FILE LOCATION ***: This will be in the same directory as the .ps1 script file by default.
$ConfigFilePath = Join-Path $scriptDir "manifest.json"
$LogFilePath    = Join-Path $scriptDir "shipwreck.log"

# --- Version ---
$script:AppVersion = "v1.0"


# --- Global Variables ---
$sourcePath = $null
$targetPaths = [System.Collections.Generic.List[string]]::new()
$logTextBox = $null # Will be assigned the GUI textbox object
$sevenZipPath = $null # Path to 7z.exe

# --- Utility Functions ---

# Normalizes image filenames in the source folder before any other steps run.
# Fixes two issues in a single pass:
#   1. Leading/trailing whitespace in the base name.
#   2. Mediux hyphen-spacing bug: "Show Name- Title (YYYY)" -> "Show Name - Title (YYYY)"
#      (Mediux omits the space before hyphens, and converts colons to hyphens without proper spacing.)
# Files that need neither fix are skipped silently. Results are logged to the GUI log.
function Trim-ImageFilenamesInFolder {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$sourcePath
    )

    # Ensure the folder exists
    if (-not (Test-Path $sourcePath -PathType Container)) {
        Add-LogEntry "⚠️ Trim/Normalize: Folder '$sourcePath' not found." -ColorInput ([System.Drawing.Color]::Orange)
        return
    }

    # Define image extensions to scan
    $imageExtensions = @(".jpg", ".jpeg", ".png", ".webp")

    $files = Get-ChildItem -Path $sourcePath -File -ErrorAction SilentlyContinue | Where-Object {
        $imageExtensions -contains $_.Extension.ToLower()
    }

    $fixedCount = 0
    foreach ($file in $files) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $ext  = $file.Extension

        # Step 1: trim leading/trailing whitespace
        $newBase = $base.Trim()

        # Step 2: fix Mediux hyphen spacing — insert missing space before "- "
        # e.g. "Dragon Ball Z- Fusion Reborn (1995)" -> "Dragon Ball Z - Fusion Reborn (1995)"
        # e.g. "2001- A Space Odyssey (1968)"        -> "2001 - A Space Odyssey (1968)"
        # The pattern only fires when a non-whitespace char is immediately followed by "-<space>",
        # so internal hyphens like "season-specials-poster" are never affected.
        $newBase = $newBase -replace '(\S)-\s', '$1 - '

        if ($newBase -ne $base) {
            $newName = "$newBase$ext"
            try {
                Rename-Item -LiteralPath $file.FullName -NewName $newName -Force -ErrorAction Stop
                Add-LogEntry "✅ Normalized: '$($file.Name)' -> '$newName'" -ColorInput ([System.Drawing.Color]::Green)
                $fixedCount++
            } catch {
                Add-LogEntry "⚠️ Failed to normalize '$($file.Name)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Orange)
            }
        }
    }

    if ($fixedCount -gt 0) {
        Add-LogEntry "   Normalized $fixedCount filename(s) in source folder." -ColorInput ([System.Drawing.Color]::Green)
    }
}

# Function to normalize show names for comparison against target folders.
# Mediux downloads omit the space before hyphens and convert colons to hyphens without proper spacing.
# Examples of what this fixes:
#   "Dragon Ball Z- Fusion Reborn (1995)"  ->  "Dragon Ball Z - Fusion Reborn (1995)"
#   "2001- A Space Odyssey (1968)"         ->  "2001 - A Space Odyssey (1968)"
# This normalization is applied to SOURCE names only, at lookup time. Source files are never renamed by this.
function Get-NormalizedShowName {
    param([string]$Name)
    # Match any non-whitespace character immediately followed by "- " and insert the missing space.
    # This handles both the "missing space before hyphen" and the "colon converted to hyphen" Mediux issues.
    return ($Name -replace '(\S)-\s', '$1 - ')
}

# Three-tier cascading match against the target show folder map.
# Returns @{ Path; MatchType = "exact"|"normalized"|"loose"; MatchedKey } or $null.
#
# Tier 1 — Exact: direct hashtable lookup (map stores both actual and pre-normalized keys).
# Tier 2 — Normalized: apply Get-NormalizedShowName (fixes Mediux missing-space-before-hyphen).
# Tier 3 — Loose: strip all hyphens/colons from both sides and compare. Handles cases like
#           "Fullmetal Alchemist- Brotherhood (2009)" vs "Fullmetal Alchemist Brotherhood (2009)"
#           where Mediux inserts a hyphen that doesn't exist in the Jellyfin folder name at all.
function Find-TargetShowMatch {
    param(
        [string]$SourceName,
        [hashtable]$TargetMap
    )

    # Tier 1: direct lookup — covers exact names and pre-normalized variants already in the map
    if ($TargetMap.ContainsKey($SourceName)) {
        return @{ Path = $TargetMap[$SourceName]; MatchType = "exact"; MatchedKey = $SourceName }
    }

    # Tier 2: Mediux spacing fix — "Name- Title" -> "Name - Title"
    $normalized = Get-NormalizedShowName -Name $SourceName
    if ($TargetMap.ContainsKey($normalized)) {
        return @{ Path = $TargetMap[$normalized]; MatchType = "normalized"; MatchedKey = $normalized }
    }

    # Tier 3: loose match — strip hyphens/colons, collapse spaces, compare case-insensitively
    $sourceLoose = ($SourceName -replace '[-:]', ' ' -replace '\s+', ' ').Trim()
    foreach ($key in $TargetMap.Keys) {
        $keyLoose = ($key -replace '[-:]', ' ' -replace '\s+', ' ').Trim()
        if ([string]::Compare($sourceLoose, $keyLoose, [System.StringComparison]::OrdinalIgnoreCase) -eq 0) {
            return @{ Path = $TargetMap[$key]; MatchType = "loose"; MatchedKey = $key }
        }
    }

    return $null
}

# Logs which match tier fired. Exact matches are silent (normal flow).
# Normalized and loose matches log a note so the user can verify correctness.
function Write-MatchLog {
    param(
        [string]$SourceName,
        [hashtable]$MatchResult
    )
    switch ($MatchResult.MatchType) {
        "normalized" {
            Add-LogEntry "   ↳ Spacing fix: '$SourceName' -> '$($MatchResult.MatchedKey)'" -ColorInput ([System.Drawing.Color]::DarkCyan)
        }
        "loose" {
            Add-LogEntry "   ↳ Loose match (punctuation ignored): '$SourceName' -> '$($MatchResult.MatchedKey)'" -ColorInput ([System.Drawing.Color]::Goldenrod)
        }
    }
}

# Function to check if 7z.exe is available in PATH
function Test-7ZipPath {
    $exe = "7z.exe"
    $envPaths = $env:Path -split ';'
    $foundPath = $null
    foreach ($p in $envPaths) {
        # Skip empty path segments
        if ([string]::IsNullOrWhiteSpace($p)) { continue }

        $fullPath = Join-Path $p $exe -ErrorAction SilentlyContinue
        # Check if $fullPath is not null and exists
        if ($null -ne $fullPath -and (Test-Path $fullPath -PathType Leaf)) {
            $foundPath = $fullPath
            break
        }
    }
    # Also check common Program Files locations as a fallback
    if (-not $foundPath) {
        $progFiles = @(
            "$env:ProgramFiles\7-Zip\$exe",
            "$env:ProgramFiles(x86)\7-Zip\$exe"
        )
        foreach ($pf in $progFiles) {
            if (Test-Path $pf -PathType Leaf) {
                $foundPath = $pf
                break
            }
        }
    }
    return $foundPath
}

# Maps a System.Drawing.Color name to the Shipwreck WPF theme palette.
# Call sites pass System.Drawing.Color so signatures stay identical across all functions.
function Get-WpfBrush {
    param($ColorInput)
    $hex = if ($ColorInput -ne $null -and $ColorInput -is [System.Drawing.Color]) {
        switch ($ColorInput.Name) {
            "Green"       { "#A2BFA6" }   # Seafoam  — success
            "DarkGreen"   { "#4A8553" }   # Coral green — strong success / hold emptied
            "Red"         { "#D65036" }   # Burnt orange — error / lost cargo
            "DarkRed"     { "#B03020" }   # Deep red — stack trace
            "Blue"        { "#7BA7BC" }   # Muted blue — in-progress operations
            "DarkBlue"    { "#C9A054" }   # Pirate gold — finished section banners
            "Orange"      { "#C9A054" }   # Pirate gold — warnings / no-result notices
            "Goldenrod"   { "#C9A054" }   # Pirate gold — deletions / cleanup
            "Gray"        { "#7A8A7E" }   # Dim — informational
            "DarkGray"    { "#5A6A5E" }   # Very dim — verbose detail
            "DarkCyan"    { "#5BBFAF" }   # Teal — tier-2 match (spacing fix)
            "DarkMagenta" { "#C9A054" }   # Pirate gold — Run All banners
            default       { "#A2BFA6" }   # Default seafoam
        }
    } else { "#A2BFA6" }
    return New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($hex)
    )
}

# Appends a styled line to the WPF RichTextBox log.
# Signature is identical to the WinForms version — all call sites are unchanged.
function Add-LogEntry {
    param(
        [string]$Message,
        [Parameter(Mandatory=$false)]
        $ColorInput = $null,
        [Parameter(Mandatory=$false)]
        [switch]$Bold,
        [bool]$NewLine = $true
    )

    $brush = Get-WpfBrush -ColorInput $ColorInput

    if ($script:logTextBox -ne $null) {
        try {
            $script:logTextBox.Dispatcher.Invoke([Action]{
                $para = New-Object System.Windows.Documents.Paragraph
                $para.Margin = New-Object System.Windows.Thickness(0)
                $para.LineHeight = 17
                $run = New-Object System.Windows.Documents.Run($Message)
                $run.Foreground = $brush
                if ($Bold.IsPresent) {
                    $run.FontWeight = [System.Windows.FontWeights]::Bold
                }
                $para.Inlines.Add($run)
                $script:logTextBox.Document.Blocks.Add($para)
                $script:logTextBox.ScrollToEnd()
            }, [System.Windows.Threading.DispatcherPriority]::Background)
        } catch {
            Write-Host $Message
        }
    } else {
        # Fallback to console before the GUI is up
        $consoleColor = switch ($ColorInput.Name) {
            "Green" {"Green"} "DarkGreen" {"DarkGreen"} "Red" {"Red"} "DarkRed" {"DarkRed"}
            "Blue" {"Blue"} "Orange" {"DarkYellow"} "Goldenrod" {"DarkYellow"}
            "Gray" {"Gray"} "DarkGray" {"DarkGray"} "DarkCyan" {"DarkCyan"}
            default { if ($Host.UI.RawUI.BackgroundColor -eq 'Black') {'White'} else {'Black'} }
        }
        Write-Host $Message -ForegroundColor $consoleColor
    }
}


# Function to load saved configuration
function Load-Configuration {
    # Ensure ConfigFilePath is valid before testing
    if ([string]::IsNullOrWhiteSpace($ConfigFilePath)) {
        Add-LogEntry "Configuration file path could not be determined. Cannot load configuration." -ColorInput ([System.Drawing.Color]::Red)
        return
    }

    if (Test-Path $ConfigFilePath -PathType Leaf) { # Check it's a file
        try {
            $configContent = Get-Content $ConfigFilePath -Raw -ErrorAction Stop
            # Handle potentially empty config file
            if ([string]::IsNullOrWhiteSpace($configContent)) {
                 Add-LogEntry "Configuration file '$ConfigFilePath' is empty." -ColorInput ([System.Drawing.Color]::Orange)
                 return
            }
            $config = $configContent | ConvertFrom-Json -ErrorAction Stop

            if ($config -ne $null) {
                if ($config.PSObject.Properties.Name -contains 'SourcePath' -and -not [string]::IsNullOrWhiteSpace($config.SourcePath)) {
                    $script:sourcePath = $config.SourcePath
                }
                if ($config.PSObject.Properties.Name -contains 'TargetPaths' -and $config.TargetPaths -ne $null) {
                    # Clear existing before adding loaded paths
                    $script:targetPaths.Clear()
                    # Ensure we only add non-empty strings
                    $config.TargetPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
                        $script:targetPaths.Add($_)
                    }
                }
                Add-LogEntry "Manifest loaded from '$ConfigFilePath'" -ColorInput ([System.Drawing.Color]::Blue)
            } else {
                 Add-LogEntry "Failed to parse JSON from configuration file '$ConfigFilePath'." -ColorInput ([System.Drawing.Color]::Red)
            }
        } catch {
            Add-LogEntry "Error loading or parsing configuration: $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
            Add-LogEntry "Configuration file might be corrupt or inaccessible: '$ConfigFilePath'" -ColorInput ([System.Drawing.Color]::Red)
        }
    } else {
        Add-LogEntry "Configuration file not found at '$ConfigFilePath'. Using defaults (blank)." -ColorInput ([System.Drawing.Color]::Orange)
    }
}

# Function to save configuration
function Save-Configuration {
    # Ensure ConfigFilePath is valid before saving
    if ([string]::IsNullOrWhiteSpace($ConfigFilePath)) {
        Add-LogEntry "Configuration file path could not be determined. Cannot save configuration." -ColorInput ([System.Drawing.Color]::Red)
        return
    }

    $config = @{
        SourcePath  = $script:sourcePath
        TargetPaths = $script:targetPaths.ToArray() # Convert List to Array for JSON
    }
    try {
        $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigFilePath -Encoding UTF8 -Force -ErrorAction Stop
        Add-LogEntry "Manifest saved to '$ConfigFilePath'" -ColorInput ([System.Drawing.Color]::Blue)
    } catch {
        Add-LogEntry "Error saving configuration: $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
    }
}


# Function to browse for a folder
function Select-FolderDialog {
    param(
        [string]$InitialDirectory = [Environment]::GetFolderPath('MyComputer'),
        [string]$Description = "Select a folder"
    )
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    # Ensure initial directory exists before setting it
    if (-not [string]::IsNullOrWhiteSpace($InitialDirectory) -and (Test-Path $InitialDirectory -PathType Container)) {
        $folderBrowser.SelectedPath = $InitialDirectory
    } else {
        # Default to MyComputer if initial is invalid/empty
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    }
    $folderBrowser.ShowNewFolderButton = $true # Allow creating new folders
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        return $null
    }
}

# --- Core Logic Functions (Adapted from User Input) ---

# Gets unique 'Show Name (YYYY)' directory names from all target paths
function Get-TargetShowNames {
    param(
        [System.Collections.Generic.List[string]]$TargetBasePaths
    )
    # Use a HashSet for efficient unique storage, ignoring case
    $allShowNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($targetBasePath in $TargetBasePaths) {
        if (Test-Path $targetBasePath -PathType Container) {
            # Get only immediate subdirectories
            Get-ChildItem -Path $targetBasePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                # Strict validation for 'Name (YYYY)' format at the end of the name
                if ($_.Name -match '^.+ \(\d{4}\)$') {
                    [void]$allShowNames.Add($_.Name)
                    # Also add the normalized form so Mediux-named archives (missing space before hyphen) can match
                    $normalizedName = Get-NormalizedShowName -Name $_.Name
                    if ($normalizedName -ne $_.Name) {
                        [void]$allShowNames.Add($normalizedName)
                    }
                } else {
                     # Add-LogEntry "Debug: Folder '$($_.Name)' in '$targetBasePath' skipped (doesn't match 'Name (YYYY)' format)." -ColorInput ([System.Drawing.Color]::DarkGray)
                }
            }
        } else {
             Add-LogEntry "Target path not found or not a directory: '$targetBasePath'" -ColorInput ([System.Drawing.Color]::Orange)
        }
    }
    # Pipe HashSet directly to Sort-Object
    return $allShowNames | Sort-Object
}

# Gets map of target show names to full paths
# Keys are stored in normalized form (space before hyphens enforced) so that Mediux-named
# source files ("Show Name- Title (YYYY)") resolve to the correct target path.
function Get-TargetShowFoldersMap {
     param(
        [System.Collections.Generic.List[string]]$TargetBasePaths
    )
    $map = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($targetRoot in $TargetBasePaths) {
         if (Test-Path $targetRoot -PathType Container) {
             Get-ChildItem -Path $targetRoot -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -match '^.+ \(\d{4}\)$'} | ForEach-Object {
                 # Key by actual name
                 if (-not $map.ContainsKey($_.Name)) {
                    $map.Add($_.Name, $_.FullName)
                 }
                 # Also key by normalized name so source files with missing spaces before hyphens still resolve
                 $normalizedName = Get-NormalizedShowName -Name $_.Name
                 if ($normalizedName -ne $_.Name -and -not $map.ContainsKey($normalizedName)) {
                    $map.Add($normalizedName, $_.FullName)
                 }
            }
         }
    }
    return $map
}


# 1. Extract Archives
function Extract-MatchingArchives {
    param(
        [string]$SourceRoot,
        [System.Collections.Generic.List[string]]$TargetRoots,
        [hashtable]$PrebuiltMap = $null
    )
    Add-LogEntry "--- Starting Archive Salvage ---" -ColorInput ([System.Drawing.Color]::Green)
    if (-not $script:sevenZipPath) {
         Add-LogEntry "7-Zip executable (7z.exe) not found in PATH or common locations. Cannot salvage .rar/.7z files." -ColorInput ([System.Drawing.Color]::Red)
         Add-LogEntry "--- Aborted Archive Salvage ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
         return
    }

    # Use the prebuilt map if provided, otherwise build locally (fallback for standalone use)
    if ($null -ne $PrebuiltMap -and $PrebuiltMap.Count -gt 0) {
        $targetMap = $PrebuiltMap
    } else {
        $targetMap = Get-TargetShowFoldersMap -TargetBasePaths $TargetRoots
    }
    if ($null -eq $targetMap -or $targetMap.Count -eq 0) {
        Add-LogEntry "   No shows found at any port. Nothing to salvage." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Archive Salvage (No Ports) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }
    Add-LogEntry "   Scouting $($targetMap.Count) show entries to match against." -ColorInput ([System.Drawing.Color]::Gray)

    $archiveExtensions = @('.zip', '.rar', '.7z')
    $archiveFiles = Get-ChildItem -Path $SourceRoot -File -ErrorAction SilentlyContinue | Where-Object { $archiveExtensions -contains $_.Extension.ToLower() }

    if ($null -eq $archiveFiles -or $archiveFiles.Count -eq 0) {
        Add-LogEntry "   No archives (.zip, .rar, .7z) found in the hold '$SourceRoot'." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Archive Salvage (No Archives) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }
    Add-LogEntry "   Found $($archiveFiles.Count) archive(s) in the hold." -ColorInput ([System.Drawing.Color]::Gray)

    $matchingArchiveFound = $false
    $imageExtensions = @(".jpg", ".jpeg", ".png") # Keep only these extensions (lowercase)

    foreach ($archiveFile in $archiveFiles) {
        $archiveBaseName = $archiveFile.BaseName
        $matchResult = Find-TargetShowMatch -SourceName $archiveBaseName -TargetMap $targetMap
        if ($null -ne $matchResult) {
            Add-LogEntry "      ✅ Match found: '$archiveBaseName'" -ColorInput ([System.Drawing.Color]::Green) -Bold
            Write-MatchLog -SourceName $archiveBaseName -MatchResult $matchResult
            $matchingArchiveFound = $true

            $archiveFilePath = $archiveFile.FullName
            $destinationPath = Join-Path $SourceRoot $archiveBaseName

            # Create destination folder
            if (-not (Test-Path -Path $destinationPath -PathType Container)) {
                Add-LogEntry "         📁 Creating extraction folder: '$destinationPath'" -ColorInput ([System.Drawing.Color]::Green)
                try {
                    New-Item -Path $destinationPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                } catch {
                     Add-LogEntry "         ❌ Error creating directory '$destinationPath': $($_.Exception.Message). Skipping." -ColorInput ([System.Drawing.Color]::Red)
                     continue
                }
            } else {
                 Add-LogEntry "         📁 Extraction folder already exists: '$destinationPath'" -ColorInput ([System.Drawing.Color]::Gray)
            }

            Add-LogEntry "         📦 Salvaging '$($archiveFile.Name)' to '$destinationPath' using 7-Zip..." -ColorInput ([System.Drawing.Color]::Blue)

            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = $script:sevenZipPath
            # Quote paths appropriately for the 7z command line
            $quotedArchivePath = "`"$archiveFilePath`""
            # Note: No space between -o and the quoted path for 7zip
            $quotedDestinationArg = "-o`"$destinationPath`""
            # Construct the final argument string
            $processInfo.Arguments = "x $quotedArchivePath $quotedDestinationArg -y"

            Add-LogEntry "   Running: `"$($processInfo.FileName)`" $($processInfo.Arguments)" -ColorInput ([System.Drawing.Color]::DarkGray) # Log the exact command string

            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $extractionSuccess = $false # Flag to track success

            try {
                 $process.Start() | Out-Null
                 # Capture potential errors from stderr
                 $errors = $process.StandardError.ReadToEnd()
                 $process.WaitForExit()

                 if ($process.ExitCode -ne 0) {
                     # Log any errors captured from stderr
                     if (-not [string]::IsNullOrWhiteSpace($errors)) {
                        Add-LogEntry "   7-Zip Errors: $errors" -ColorInput ([System.Drawing.Color]::Red) -ErrorAction SilentlyContinue
                     }
                     Throw "7-Zip extraction failed with exit code $($process.ExitCode)."
                 }

                Add-LogEntry "         ✅ Salvage complete." -ColorInput ([System.Drawing.Color]::Green)
                $extractionSuccess = $true # Mark extraction as successful

            } catch {
                Add-LogEntry "         ❌ Error during 7-Zip execution for '$archiveFilePath': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                # Log $process.ExitCode if available
                if ($process -ne $null -and $process.HasExited) { Add-LogEntry "   7-Zip Exit Code: $($process.ExitCode)" -ColorInput ([System.Drawing.Color]::Red) }
            } finally {
                 if ($process -ne $null) { $process.Dispose() }
            }

            # Only proceed with delete and cleanup if extraction was successful
            if ($extractionSuccess) {
                # *** ADDED RETRY LOGIC FOR DELETION ***
                $deleteSuccess = $false
                $maxRetries = 5
                $retryDelaySeconds = 1
                for ($retry = 1; $retry -le $maxRetries; $retry++) {
                    try {
                        Remove-Item -Path $archiveFilePath -Force -ErrorAction Stop
                        Add-LogEntry "         🗑️ Deleted archive: '$archiveFilePath'" -ColorInput ([System.Drawing.Color]::Goldenrod)
                        $deleteSuccess = $true
                        break
                    } catch [System.IO.IOException] {
                        if ($retry -eq $maxRetries) {
                            Add-LogEntry "         ⚠️ Could not delete archive '$archiveFilePath' after $maxRetries attempts: $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                        } else {
                             Add-LogEntry "         File lock on '$($archiveFile.Name)', retrying in $retryDelaySeconds second(s)... (Attempt $retry/$maxRetries)" -ColorInput ([System.Drawing.Color]::Orange)
                             Start-Sleep -Seconds $retryDelaySeconds
                        }
                    } catch {
                         Add-LogEntry "         ⚠️ Unexpected error deleting archive '$archiveFilePath': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                         break
                    }
                } # End retry loop

                # --- Post-Extraction Cleanup within the new folder ---
                if ($deleteSuccess) {
                    Add-LogEntry "         🧹 Cleaning up non-image files in '$destinationPath'..." -ColorInput ([System.Drawing.Color]::Blue)
                    $filesInNewFolder = Get-ChildItem -Path $destinationPath -Recurse -File -ErrorAction SilentlyContinue
                    if ($null -ne $filesInNewFolder) {
                        $filesToDelete = $filesInNewFolder | Where-Object { $imageExtensions -notcontains $_.Extension.ToLower() }
                        if ($filesToDelete.Count -gt 0) {
                            foreach ($fileItem in $filesToDelete) {
                                try {
                                    Remove-Item $fileItem.FullName -Force -ErrorAction Stop
                                    Add-LogEntry "         🧹 Removed non-image: $($fileItem.FullName)" -ColorInput ([System.Drawing.Color]::Goldenrod)
                                } catch {
                                    Add-LogEntry "         ⚠️ Could not remove: $($fileItem.FullName) — $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                                }
                            }
                        } else {
                            Add-LogEntry "         No non-image files to clean up." -ColorInput ([System.Drawing.Color]::Gray)
                        }
                    } else {
                        Add-LogEntry "         No files found inside '$destinationPath' after extraction." -ColorInput ([System.Drawing.Color]::Gray)
                    }
                }
            } # End if extractionSuccess

        } # End if match found for archive
    } # End foreach ($archiveFile in $archiveFiles)

    if (-not $matchingArchiveFound) {
        Add-LogEntry "   No archives in the hold matched any known port." -ColorInput ([System.Drawing.Color]::Orange)
    }
    Add-LogEntry "--- Finished Archive Salvage ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
}


# 2. Rename Existing Backdrops & Move Loose Backdrops
function Rename-ExistingBackdrops {
    param(
        [string]$SourceRoot,
        [System.Collections.Generic.List[string]]$TargetRoots,
        [hashtable]$PrebuiltMap = $null
    )
    Add-LogEntry "--- Starting Backdrop Catalogue ---" -ColorInput ([System.Drawing.Color]::Green)
    # Use the prebuilt map if provided, otherwise build locally (fallback for standalone use)
    if ($null -ne $PrebuiltMap -and $PrebuiltMap.Count -gt 0) {
        $targetShowFoldersMap = $PrebuiltMap
    } else {
        $targetShowFoldersMap = Get-TargetShowFoldersMap -TargetBasePaths $TargetRoots
    }
    if ($targetShowFoldersMap.Count -eq 0) {
        Add-LogEntry "   No shows found at any port. Cannot catalogue backdrops." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Backdrop Catalogue (No Ports) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }

    $renamedCount = 0
    $movedLooseCount = 0
    $imageExtensionsForBackdrop = @(".jpg", ".jpeg", ".png")
    $regexLooseBackdrop = '^(.+ \(\d{4}\)) - Backdrop$'

    # --- Process Loose Backdrop files in SourceRoot first ---
    Add-LogEntry "   🔍 Checking for loose backdrop files in hold '$SourceRoot'..." -ColorInput ([System.Drawing.Color]::Gray)
    $looseSourceFiles = Get-ChildItem -Path $SourceRoot -File -Force -ErrorAction SilentlyContinue | Where-Object { $imageExtensionsForBackdrop -contains $_.Extension.ToLower() }

    if ($null -ne $looseSourceFiles) {
        foreach ($looseFile in $looseSourceFiles) {
            # Check if the loose file matches the 'Show Name (YYYY) - Backdrop' pattern
            if ($looseFile.BaseName -match $regexLooseBackdrop) {
                $showNamePart = $matches[1]
                # Three-tier match: exact -> normalized -> loose punctuation
                $matchResult = Find-TargetShowMatch -SourceName $showNamePart -TargetMap $targetShowFoldersMap
                if ($null -ne $matchResult) {
                    $targetShowPath = $matchResult.Path
                    Write-MatchLog -SourceName $showNamePart -MatchResult $matchResult
                    $targetFileName = "backdrop$($looseFile.Extension)"
                    $targetPath = Join-Path $targetShowPath $targetFileName

                    Add-LogEntry "      📋 Loose backdrop matched: '$($looseFile.Name)' -> '$targetFileName'" -ColorInput ([System.Drawing.Color]::Gray)
                    try {
                        if (-not (Test-Path $targetShowPath -PathType Container)) {
                            Add-LogEntry "         ❌ Port '$targetShowPath' does not exist! Cannot haul '$($looseFile.Name)'." -ColorInput ([System.Drawing.Color]::Red)
                        } else {
                            Move-Item -Path $looseFile.FullName -Destination $targetPath -Force -ErrorAction Stop
                            Add-LogEntry "         🚢 Hauled loose backdrop '$($looseFile.Name)' -> '$targetPath'" -ColorInput ([System.Drawing.Color]::Green)
                            $movedLooseCount++
                        }
                    } catch {
                         Add-LogEntry "         ⚠️ Lost cargo '$($looseFile.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                    }
                } # End if show name matches target
            } # End if loose file matches backdrop pattern
        } # End foreach loose file
    } else {
         Add-LogEntry "   No loose image files found in '$SourceRoot' to check for backdrops." -ColorInput ([System.Drawing.Color]::Gray)
    }
    # --- END Loose File Processing ---


    # --- Process Subfolders ---
    $sourceFolders = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue
    if($null -eq $sourceFolders -or $sourceFolders.Count -eq 0) {
        Add-LogEntry "   No subfolders found in the hold '$SourceRoot'." -ColorInput ([System.Drawing.Color]::Gray)
    } else {
        foreach ($folder in $sourceFolders) {
            $folderMatchResult = Find-TargetShowMatch -SourceName $folder.Name -TargetMap $targetShowFoldersMap
            if ($null -ne $folderMatchResult) {
                Write-MatchLog -SourceName $folder.Name -MatchResult $folderMatchResult
                Add-LogEntry "      🔍 Checking for backdrops in: '$($folder.FullName)'" -ColorInput ([System.Drawing.Color]::Gray)
                Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
                    Where-Object { ($imageExtensionsForBackdrop -contains $_.Extension.ToLower()) -and ($_.BaseName -like "* - Backdrop") } |
                    ForEach-Object {
                        $fileToRename = $_
                        $newFileName = "backdrop$($fileToRename.Extension)"
                        $newPath = Join-Path $fileToRename.DirectoryName $newFileName
                        try {
                            Rename-Item -Path $fileToRename.FullName -NewName $newFileName -Force -ErrorAction Stop
                            Add-LogEntry "         📋 Catalogued: '$($fileToRename.Name)' -> '$newFileName'" -ColorInput ([System.Drawing.Color]::Green)
                            $renamedCount++
                        } catch {
                            Add-LogEntry "         ⚠️ Could not catalogue '$($fileToRename.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                        }
                    }
            } else {
                # Add-LogEntry "Debug: Skipping source folder '$($folder.Name)' as it doesn't match any target show name." -ColorInput ([System.Drawing.Color]::DarkGray)
            }
        } # End foreach folder
    } # End else

    if ($renamedCount -gt 0) {
        Add-LogEntry "   📋 Catalogued $renamedCount backdrop file(s) within subfolders." -ColorInput ([System.Drawing.Color]::Green)
    }
     if ($movedLooseCount -gt 0) {
        Add-LogEntry "   🚢 Hauled $movedLooseCount loose backdrop file(s) to port." -ColorInput ([System.Drawing.Color]::Green)
    }
    if ($renamedCount -eq 0 -and $movedLooseCount -eq 0) {
        Add-LogEntry "   No backdrop files found that required cataloguing." -ColorInput ([System.Drawing.Color]::Orange)
    }
    Add-LogEntry "--- Finished Backdrop Catalogue ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
}


# 3. Rename Media Images (Season Posters, Loose Season Posters, Loose Show Posters, Folder Images in Subfolders)
function Rename-SeasonPosters { # Keep name for button binding, but logic expanded
    param(
        [string]$SourceRoot,
        [System.Collections.Generic.List[string]]$TargetRoots,
        [hashtable]$PrebuiltMap = $null
    )
     Add-LogEntry "--- Starting Media Image Catalogue ---" -ColorInput ([System.Drawing.Color]::Green)

    # Use the prebuilt map if provided, otherwise build locally (fallback for standalone use)
    if ($null -ne $PrebuiltMap -and $PrebuiltMap.Count -gt 0) {
        $targetShowFoldersMap = $PrebuiltMap
    } else {
        $targetShowFoldersMap = Get-TargetShowFoldersMap -TargetBasePaths $TargetRoots
    }
    if ($targetShowFoldersMap.Count -eq 0) {
        Add-LogEntry "No shows found at any port. Cannot catalogue media images." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Media Image Catalogue (No Ports) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }

    $imageExtensions = @(".jpg", ".jpeg", ".png")
    $renamedSeasonCount = 0
    $renamedFolderCount = 0 # Counter for folder images renamed within subfolders
    $movedLooseSeasonPosterCount = 0
    $movedLooseShowPosterCount = 0 # Counter for loose show posters moved/renamed to folder.ext

    # --- Process loose files in SourceRoot first ---
    Add-LogEntry "🔍 Checking for loose image files in source root '$SourceRoot'..." -ColorInput ([System.Drawing.Color]::Gray)
    $looseSourceFiles = Get-ChildItem -Path $SourceRoot -File -Force -ErrorAction SilentlyContinue | Where-Object { $imageExtensions -contains $_.Extension.ToLower() }

    if ($null -ne $looseSourceFiles) {
        foreach ($looseFile in $looseSourceFiles) {
            # Check 1: Loose Season Poster ('Show Name (YYYY) - Season X')
            if ($looseFile.BaseName -match '^(.+ \(\d{4}\)) - Season (\d+)$') {
                $showNamePart = $matches[1]
                $seasonNum = [int]$matches[2]
                $extension = $looseFile.Extension
                # Three-tier match: exact -> normalized -> loose punctuation
                $matchResult = Find-TargetShowMatch -SourceName $showNamePart -TargetMap $targetShowFoldersMap

                if ($null -ne $matchResult) {
                    $targetShowPath = $matchResult.Path
                    Write-MatchLog -SourceName $showNamePart -MatchResult $matchResult
                    $targetFileName = ""
                    if ($seasonNum -eq 0) { $targetFileName = "season-specials-poster$extension" }
                    else { $targetFileName = "season{0:D2}-poster{1}" -f $seasonNum, $extension }
                    $targetPath = Join-Path $targetShowPath $targetFileName

                    Add-LogEntry "      📋 Loose season poster matched: '$($looseFile.Name)' -> '$targetFileName'" -ColorInput ([System.Drawing.Color]::Gray)
                    try {
                        if (-not (Test-Path $targetShowPath -PathType Container)) { Add-LogEntry "         ❌ Port '$targetShowPath' does not exist! Cannot haul '$($looseFile.Name)'." -ColorInput ([System.Drawing.Color]::Red) }
                        else {
                            Move-Item -Path $looseFile.FullName -Destination $targetPath -Force -ErrorAction Stop
                            Add-LogEntry "         🚢 Hauled loose season poster '$($looseFile.Name)' -> '$targetPath'" -ColorInput ([System.Drawing.Color]::Green)
                            $movedLooseSeasonPosterCount++
                        }
                    } catch { Add-LogEntry "         ⚠️ Lost cargo '$($looseFile.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red) }
                } # End if show name matches target
            }
            # Check 2: Loose Show Poster ('Show Name (YYYY).ext')
            else {
                 $matchResult = Find-TargetShowMatch -SourceName $looseFile.BaseName -TargetMap $targetShowFoldersMap
                 if ($null -ne $matchResult) {
                 $targetShowPath = $matchResult.Path
                 Write-MatchLog -SourceName $looseFile.BaseName -MatchResult $matchResult
                 $targetFileName = "folder$($looseFile.Extension)" # Target name is folder.ext
                 $targetPath = Join-Path $targetShowPath $targetFileName

                 Add-LogEntry "      📋 Loose show poster matched: '$($looseFile.Name)' -> '$targetFileName'" -ColorInput ([System.Drawing.Color]::Gray)
                 try {
                     if (-not (Test-Path $targetShowPath -PathType Container)) { Add-LogEntry "         ❌ Port '$targetShowPath' does not exist! Cannot haul '$($looseFile.Name)'." -ColorInput ([System.Drawing.Color]::Red) }
                     else {
                         Move-Item -Path $looseFile.FullName -Destination $targetPath -Force -ErrorAction Stop
                         Add-LogEntry "         🚢 Hauled loose show poster '$($looseFile.Name)' -> '$targetPath'" -ColorInput ([System.Drawing.Color]::Green)
                         $movedLooseShowPosterCount++
                     }
                 } catch { Add-LogEntry "         ⚠️ Lost cargo '$($looseFile.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red) }
                 } # End if loose file matches show pattern
            } # End else
        } # End foreach loose file
    } else {
         Add-LogEntry "   No loose image files found in '$SourceRoot' to check." -ColorInput ([System.Drawing.Color]::Gray)
    }
    # --- END Loose File Processing ---


    # --- Process Subfolders (For Season Posters and Folder Images) ---
    $sourceFolders = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue
     if($null -eq $sourceFolders -or $sourceFolders.Count -eq 0) {
        Add-LogEntry "   No subfolders found in the hold '$SourceRoot'." -ColorInput ([System.Drawing.Color]::Gray)
     } else {
        foreach ($folder in $sourceFolders) {
            $folderMatchResult = Find-TargetShowMatch -SourceName $folder.Name -TargetMap $targetShowFoldersMap
            if ($null -ne $folderMatchResult) {
                Write-MatchLog -SourceName $folder.Name -MatchResult $folderMatchResult
                $sourceFolderPath = $folder.FullName
                Add-LogEntry "      🔍 Checking for images in: '$sourceFolderPath'" -ColorInput ([System.Drawing.Color]::Gray)

                $imageFiles = Get-ChildItem -Path $sourceFolderPath -File -ErrorAction SilentlyContinue | Where-Object { $imageExtensions -contains $_.Extension.ToLower() }

                foreach ($file in $imageFiles) {
                    if ($file.BaseName -match ' - Season (\d+)$') {
                        $seasonNum = [int]$matches[1]
                        $extension = $file.Extension
                        $expectedPrefix = $folder.Name

                        if ($file.BaseName -like "$expectedPrefix - Season $seasonNum") {
                            $newName = ""
                            if ($seasonNum -eq 0) {
                                $newName = "season-specials-poster$extension"
                            } else {
                                $formattedSeason = "{0:D2}" -f $seasonNum
                                $newName = "season$formattedSeason-poster$extension"
                            }
                            $newPath = Join-Path $sourceFolderPath $newName
                            try {
                                Rename-Item -Path $file.FullName -NewName $newName -Force -ErrorAction Stop
                                Add-LogEntry "         📋 Catalogued season poster: '$($file.Name)' -> '$newName'" -ColorInput ([System.Drawing.Color]::Green)
                                $renamedSeasonCount++
                            } catch {
                                Add-LogEntry "         ❌ Failed to catalogue '$($file.Name)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                            }
                        }
                    }
                    elseif ($file.BaseName -eq $folder.Name) {
                        $extension = $file.Extension
                        $newName = "folder$extension"
                        $newPath = Join-Path $sourceFolderPath $newName
                        try {
                            Rename-Item -Path $file.FullName -NewName $newName -Force -ErrorAction Stop
                            Add-LogEntry "         📋 Catalogued folder image: '$($file.Name)' -> '$newName'" -ColorInput ([System.Drawing.Color]::Green)
                            $renamedFolderCount++
                        } catch {
                            Add-LogEntry "         ❌ Failed to catalogue '$($file.Name)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                        }
                    }
                } # End foreach file
            } # End if folder name matches target
        } # End foreach folder
     } # End else (source folders exist)

    $totalRenamedSubfolder = $renamedSeasonCount + $renamedFolderCount
    $totalProcessedLoose   = $movedLooseSeasonPosterCount + $movedLooseShowPosterCount

    if ($totalRenamedSubfolder -gt 0) {
         Add-LogEntry "   📋 Catalogued $renamedSeasonCount season poster(s) and $renamedFolderCount folder image(s) within subfolders." -ColorInput ([System.Drawing.Color]::Green)
    }
    if ($movedLooseSeasonPosterCount -gt 0) {
         Add-LogEntry "   🚢 Hauled $movedLooseSeasonPosterCount loose season poster(s) to port." -ColorInput ([System.Drawing.Color]::Green)
    }
    if ($movedLooseShowPosterCount -gt 0) {
        Add-LogEntry "   🚢 Hauled $movedLooseShowPosterCount loose show poster(s) to port." -ColorInput ([System.Drawing.Color]::Green)
    }

    if ($totalRenamedSubfolder -eq 0 -and $totalProcessedLoose -eq 0) {
         Add-LogEntry "   No season or show poster images found that required cataloguing." -ColorInput ([System.Drawing.Color]::Orange)
    }
    Add-LogEntry "--- Finished Media Image Catalogue ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
}


# 4. Rename Episode Images in Source Folders
function Rename-EpisodeImages {
     param(
        [string]$SourceRoot,
        [System.Collections.Generic.List[string]]$TargetRoots,
        [hashtable]$PrebuiltMap = $null
    )
    Add-LogEntry "--- Starting Episode Image Catalogue ---" -ColorInput ([System.Drawing.Color]::Green)

    # Use the prebuilt map if provided, otherwise build locally (fallback for standalone use)
    if ($null -ne $PrebuiltMap -and $PrebuiltMap.Count -gt 0) {
        $targetShowFoldersMap = $PrebuiltMap
    } else {
        $targetShowFoldersMap = Get-TargetShowFoldersMap -TargetBasePaths $TargetRoots
    }
    if ($targetShowFoldersMap.Count -eq 0) {
        Add-LogEntry "No shows found at any port. Cannot catalogue episode images." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Episode Image Catalogue (No Ports) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }

    $sourceFolders = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue
    if($null -eq $sourceFolders -or $sourceFolders.Count -eq 0) {
        Add-LogEntry "No subfolders found in the hold '$SourceRoot'." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Episode Image Catalogue (Empty Hold) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }

    # Extensions to scan (using patterns for -Filter)
    $imageExtensionFilters = @("*.jpg", "*.jpeg", "*.png")
    $videoExtensions = @("*.mkv", "*.mp4", "*.avi", "*.mov", "*.mpg", "*.ts") # Add more if needed
    # Regex to find SxxExx
    $regexPattern = '(?i)S(\d{1,2})\s*E(\d{1,2})'
    $renamedCount = 0

    foreach ($sourceFolder in $sourceFolders) {
        $folderMatchResult = Find-TargetShowMatch -SourceName $sourceFolder.Name -TargetMap $targetShowFoldersMap
        if ($null -ne $folderMatchResult) {
            $targetShowPath = $folderMatchResult.Path
            Write-MatchLog -SourceName $sourceFolder.Name -MatchResult $folderMatchResult
            Add-LogEntry "      ✅ Match found: '$($sourceFolder.Name)' -> '$targetShowPath'" -ColorInput ([System.Drawing.Color]::Green)

            $videoFilesHash = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
            Add-LogEntry "         🤿 Diving for video files in '$targetShowPath' (recursive)..." -ColorInput ([System.Drawing.Color]::Gray)
            $targetVideoFiles = Get-ChildItem -Path $targetShowPath -Include $videoExtensions -Recurse -File -ErrorAction SilentlyContinue

            if ($null -ne $targetVideoFiles) {
                foreach ($videoFile in $targetVideoFiles) {
                    if ($videoFile.Name -match $regexPattern) {
                        $season = $matches[1].PadLeft(2, '0')
                        $episode = $matches[2].PadLeft(2, '0')
                        $key = "S${season}E${episode}"
                        if (-not $videoFilesHash.ContainsKey($key)) {
                            $videoFilesHash.Add($key, @{
                                BaseName  = $videoFile.BaseName
                                Directory = $videoFile.DirectoryName
                            })
                        }
                    }
                }
            }

            if ($videoFilesHash.Count -eq 0) {
                Add-LogEntry "         ⚠️ No video files matching SxxExx found in '$targetShowPath'. Skipping." -ColorInput ([System.Drawing.Color]::Orange)
                continue
            } else {
                Add-LogEntry "         Found $($videoFilesHash.Count) unique SxxExx video keys." -ColorInput ([System.Drawing.Color]::Gray)
            }

            Add-LogEntry "         🔍 Scouting for episode images in '$($sourceFolder.FullName)'..." -ColorInput ([System.Drawing.Color]::Gray)
            $foundSourceImages = $false
            foreach ($filter in $imageExtensionFilters) {
                $imageFiles = Get-ChildItem -Path $sourceFolder.FullName -Filter $filter -File -ErrorAction SilentlyContinue

                if ($null -ne $imageFiles) {
                    $foundSourceImages = $true
                    foreach ($imageFile in $imageFiles) {
                        $regexMatch = $imageFile.Name -match $regexPattern

                        if ($imageFile.BaseName -notmatch '(?i)^(folder|.*-thumb)$' -and $imageFile.BaseName -ne $sourceFolder.Name -and $regexMatch) {
                            $season = $matches[1].PadLeft(2, '0')
                            $episode = $matches[2].PadLeft(2, '0')
                            $key = "S${season}E${episode}"

                            if ($videoFilesHash.ContainsKey($key)) {
                                $videoEntry    = $videoFilesHash[$key]
                                $videoBaseName = $videoEntry.BaseName
                                $videoDir      = $videoEntry.Directory
                                $newFileName   = "${videoBaseName}-thumb$($imageFile.Extension)"
                                $newFilePath   = Join-Path -Path $imageFile.DirectoryName -ChildPath $newFileName

                                try {
                                    Rename-Item -Path $imageFile.FullName -NewName $newFileName -Force -ErrorAction Stop
                                    Add-LogEntry "            📋 Catalogued: '$($imageFile.Name)' -> '$newFileName'" -ColorInput ([System.Drawing.Color]::Green)
                                    $renamedCount++

                                    if (-not $script:episodeTargetDirMap.ContainsKey($newFileName)) {
                                        $script:episodeTargetDirMap[$newFileName] = $videoDir
                                    }
                                } catch {
                                    Add-LogEntry "            ❌ Failed to catalogue '$($imageFile.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
                                }
                            } else {
                                Add-LogEntry "            🏝️ Key '$key' for '$($imageFile.Name)' NOT found in video hash — adrift at sea." -ColorInput ([System.Drawing.Color]::Red)
                            }
                        }
                    } # End foreach imageFile
                } # End if imageFiles not null
            } # End foreach filter

            if (-not $foundSourceImages) {
                Add-LogEntry "         No episode images found in '$($sourceFolder.FullName)'." -ColorInput ([System.Drawing.Color]::Gray)
            }

        } # End if source folder matches target show
    } # End foreach source folder

     if ($renamedCount -gt 0) {
         Add-LogEntry "   📋 Catalogued $renamedCount episode image file(s)." -ColorInput ([System.Drawing.Color]::Green)
     } else {
         Add-LogEntry "   No episode images found that required cataloguing." -ColorInput ([System.Drawing.Color]::Orange)
     }
     Add-LogEntry "--- Finished Episode Image Catalogue ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
}


# 5. Move Processed Image Files from Source Subfolders to Target
function Move-ImageFilesToServer {
    param(
        [string]$SourceRoot,
        [System.Collections.Generic.List[string]]$TargetRoots,
        [hashtable]$PrebuiltMap = $null
    )
    Add-LogEntry "--- Starting Cargo Haul ---" -ColorInput ([System.Drawing.Color]::Green)

    # Use the prebuilt map if provided, otherwise build locally (fallback for standalone use)
    if ($null -ne $PrebuiltMap -and $PrebuiltMap.Count -gt 0) {
        $targetShowFoldersMap = $PrebuiltMap
    } else {
        $targetShowFoldersMap = Get-TargetShowFoldersMap -TargetBasePaths $TargetRoots
    }
    if ($targetShowFoldersMap.Count -eq 0) {
        Add-LogEntry "No shows found at any port. Cannot haul cargo." -ColorInput ([System.Drawing.Color]::Orange)
        Add-LogEntry "--- Finished Cargo Haul (No Ports) ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
        return
    }

    # Define image extensions for filtering (lowercase)
    $imageExtensionsToMove = @(".jpg", ".jpeg", ".png", ".webp")

    # Regex Patterns for different image types (Case-insensitive)
    $regexEpisodeThumb = '^(.*)-thumb\.(jpg|jpeg|png|webp)$' # Matches base video name + -thumb.ext
    $regexSeasonPoster = '(?i)^season(\d{2})-poster\.(jpg|jpeg|png|webp)$'
    $regexSpecialsPoster = '(?i)^season-specials-poster\.(jpg|jpeg|png|webp)$'
    $regexBackdrop = '(?i)^backdrop\.(jpg|jpeg|png|webp)$'
    # Regex to find SxxExx within the episode thumb base name for routing
    $regexFindEpisodeKey = '(?i)S(\d{1,2})\s*E(\d{1,2})'

    $movedCount = 0
    $deletedFolderCount = 0 # Count for deleted subfolders

    # *** REMOVED SECTION: Processing loose files moved to Buttons 2 & 3 ***


    # --- Process Subfolders ---
    $sourceFolders = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue
    if($null -eq $sourceFolders -or $sourceFolders.Count -eq 0) {
        Add-LogEntry "   No subfolders found in the hold '$SourceRoot'." -ColorInput ([System.Drawing.Color]::Gray)
    } else {
        # Process each source subfolder
        foreach ($sourceFolder in $sourceFolders) {
            # Three-tier match: exact -> normalized -> loose punctuation
            $folderMatchResult = Find-TargetShowMatch -SourceName $sourceFolder.Name -TargetMap $targetShowFoldersMap
            if ($null -ne $folderMatchResult) {
                $targetShowPath = $folderMatchResult.Path
                Write-MatchLog -SourceName $sourceFolder.Name -MatchResult $folderMatchResult
                Add-LogEntry "🚢 Hauling cargo from '$($sourceFolder.FullName)' -> '$targetShowPath'" -ColorInput ([System.Drawing.Color]::Green)

                # Get target season/specials folder paths for this show (case-insensitive map)
                $targetSeasonMap = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase) # Key: "01", "02" etc. Value: Full Path
                $targetSpecialsPath = $null
                Get-ChildItem -Path $targetShowPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.Name -match '(?i)^Season (\d+)$') {
                        $seasonNumStr = $matches[1].PadLeft(2, '0')
                        if (-not $targetSeasonMap.ContainsKey($seasonNumStr)) {
                            $targetSeasonMap.Add($seasonNumStr, $_.FullName)
                        }
                    } elseif ($_.Name -eq "Specials") { # Specials folder name is usually exact case
                        $targetSpecialsPath = $_.FullName
                    }
                }
                Add-LogEntry "   Found $($targetSeasonMap.Count) target season folders and $(if($targetSpecialsPath){'a'}else{'no'}) Specials folder." -ColorInput ([System.Drawing.Color]::Gray)

                # Process each image file in the source show folder
                # Use Where-Object based on extension, add -Force
                $sourceImageFiles = Get-ChildItem -Path $sourceFolder.FullName -File -Force -ErrorAction SilentlyContinue | Where-Object { $imageExtensionsToMove -contains $_.Extension.ToLower() }

                if ($null -ne $sourceImageFiles) {
                    foreach ($file in $sourceImageFiles) {
                        $targetPath = $null
                        $moveDescription = ""

                        try {
                            # A. Handle Episode Thumbnails
                            if ($file.Name -match $regexEpisodeThumb) {
                                $videoBaseName = $matches[1]
                                # Primary: use the directory map built during Rename-EpisodeImages —
                                # routes to exactly the folder the video file lives in, no folder-name scan needed
                                if ($null -ne $script:episodeTargetDirMap -and $script:episodeTargetDirMap.ContainsKey($file.Name)) {
                                    $targetSeasonDir = $script:episodeTargetDirMap[$file.Name]
                                    $moveDescription = "episode thumb (routed from video location)"
                                    if ($targetSeasonDir) { $targetPath = Join-Path $targetSeasonDir $file.Name }
                                }
                                # Fallback: derive season from filename and scan season folders
                                # (used when Move is run standalone without a prior Rename-EpisodeImages run)
                                elseif ($videoBaseName -match $regexFindEpisodeKey) {
                                    $seasonNum = $matches[1].PadLeft(2, '0')
                                    $episodeNum = $matches[2].PadLeft(2, '0')
                                    $targetSeasonDir = $null
                                    if ($seasonNum -eq "00" -and $targetSpecialsPath) {
                                        $targetSeasonDir = $targetSpecialsPath
                                        $moveDescription = "episode thumb (S${seasonNum}E${episodeNum} - Specials)"
                                    } elseif ($targetSeasonMap.ContainsKey($seasonNum)) {
                                        $targetSeasonDir = $targetSeasonMap[$seasonNum]
                                        $moveDescription = "episode thumb (S${seasonNum}E${episodeNum})"
                                    }
                                    if ($targetSeasonDir) { $targetPath = Join-Path $targetSeasonDir $file.Name }
                                    else { Add-LogEntry "      ❓ Cannot determine target season/specials folder for S${seasonNum}. Skipping move for '$($file.Name)'." -ColorInput ([System.Drawing.Color]::Orange) }
                                } else { Add-LogEntry "      ❓ Could not extract SxxExx from base name '$videoBaseName'. Skipping move for '$($file.Name)'." -ColorInput ([System.Drawing.Color]::Orange) }
                            }
                            # B. Handle seasonXX-poster.ext
                            elseif ($file.Name -match $regexSeasonPoster) {
                                $seasonNum = $matches[1]
                                $targetPath = Join-Path $targetShowPath $file.Name
                                $moveDescription = "season $seasonNum poster"
                                if (!$targetSeasonMap.ContainsKey($seasonNum)) { Add-LogEntry "      ❓ Target Season $seasonNum folder does not exist. Moving poster to main show folder anyway." -ColorInput ([System.Drawing.Color]::Orange) }
                            }
                            # C. Handle season-specials-poster.ext
                            elseif ($file.Name -match $regexSpecialsPoster) {
                                $targetPath = Join-Path $targetShowPath $file.Name
                                $moveDescription = "specials poster"
                                if (!$targetSpecialsPath) { Add-LogEntry "      ❓ Target Specials folder does not exist. Moving poster to main show folder anyway." -ColorInput ([System.Drawing.Color]::Orange) }
                            }
                            # D. Handle backdrop.ext
                            elseif ($file.Name -match $regexBackdrop) {
                                $targetPath = Join-Path $targetShowPath $file.Name
                                $moveDescription = "backdrop"
                            }
                            # E. Handle 'folder.jpg' (renamed by Button 3)
                            elseif ($file.Name -match '(?i)^folder\.(jpg|jpeg|png|webp)$') {
                                $targetPath = Join-Path $targetShowPath $file.Name
                                $moveDescription = "folder image"
                            }
                             # F. Removed check for BaseName matching folder name

                            # --- Perform the Move ---
                            if ($targetPath) {
                                # REMOVED Test-Path check to allow overwrite
                                $targetDir = Split-Path $targetPath -Parent
                                if (-not (Test-Path $targetDir -PathType Container)) {
                                    Add-LogEntry "      ❌ Target directory '$targetDir' does not exist! Cannot move '$($file.Name)'." -ColorInput ([System.Drawing.Color]::Red)
                                } else {
                                    Move-Item -Path $file.FullName -Destination $targetPath -Force -ErrorAction Stop
                                    Add-LogEntry "      ✅ Cargo delivered — $moveDescription '$($file.Name)' to '$targetPath'" -ColorInput ([System.Drawing.Color]::Green)
                                    $movedCount++
                                }
                            } elseif ($moveDescription -eq "") {
                                 Add-LogEntry "      ❓ Unrecognized cargo in hold: '$($file.Name)'. Not moved." -ColorInput ([System.Drawing.Color]::Orange)
                            }
                        } catch { Add-LogEntry "      ⚠️ Lost cargo '$($file.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red) }
                    } # End ForEach image file in subfolder
                } else { Add-LogEntry "   No image files found in the hold '$($sourceFolder.FullName)'." -ColorInput ([System.Drawing.Color]::Gray) }

                # Check for straggler files — images still in the source folder that had no matching video
                $stragglers = Get-ChildItem -Path $sourceFolder.FullName -File -Force -ErrorAction SilentlyContinue | Where-Object { $imageExtensionsToMove -contains $_.Extension.ToLower() }
                if ($null -ne $stragglers -and $stragglers.Count -gt 0) {
                    foreach ($straggler in $stragglers) {
                        Add-LogEntry "      🏝️ '$($straggler.Name)' couldn't find a matey — adrift at sea." -ColorInput ([System.Drawing.Color]::DarkCyan)
                    }
                }

                # Delete the source subfolder if it's now empty
                try {
                     if ((Get-ChildItem -Path $sourceFolder.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                        Add-LogEntry "🪣 Emptying the hold: '$($sourceFolder.FullName)'" -ColorInput ([System.Drawing.Color]::DarkGreen)
                        Remove-Item -Path $sourceFolder.FullName -Recurse -Force -ErrorAction Stop
                        $deletedFolderCount++
                     } else {
                        Add-LogEntry "🟡 Hold not empty, leaving it be: '$($sourceFolder.FullName)'" -ColorInput ([System.Drawing.Color]::Orange)
                     }
                } catch { Add-LogEntry "⚠️ Lost cargo — could not empty hold '$($sourceFolder.FullName)': $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red) }

            } else { # End if targetShowFoldersMap contains sourceFolder.Name
                 # Add-LogEntry "Debug: Source folder '$($sourceFolder.Name)' does not match any target show folder name. Skipping move." -ColorInput ([System.Drawing.Color]::DarkGray)
            }
        } # End foreach sourceFolder
    } # End else (if sourceFolders exist)

     if ($movedCount -eq 0) {
        Add-LogEntry "   No cargo was hauled from the hold." -ColorInput ([System.Drawing.Color]::Orange)
    } else {
         Add-LogEntry "   🚢 Successfully hauled $movedCount piece(s) of cargo to port." -ColorInput ([System.Drawing.Color]::Green)
    }
     if ($deletedFolderCount -gt 0) {
        Add-LogEntry "   🪣 Emptied $deletedFolderCount hold(s)." -ColorInput ([System.Drawing.Color]::DarkGreen)
     }
    Add-LogEntry "--- Finished Cargo Haul ---" -ColorInput ([System.Drawing.Color]::DarkBlue)
}


# =============================================================================
# --- WPF GUI ---
# =============================================================================

# XAML layout — Deep Ocean dark theme per the Shipwreck design guide
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Shipwreck $($script:AppVersion)"
    Width="1400" Height="700"
    MinWidth="800" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    Background="#1A1D1E"
    WindowStyle="None"
    ResizeMode="CanResizeWithGrip">

    <Window.Resources>

        <!-- ── Step Button (outline style, gold hover) ── -->
        <Style x:Key="StepBtn" TargetType="Button">
            <Setter Property="Background"       Value="Transparent"/>
            <Setter Property="Foreground"       Value="#A2BFA6"/>
            <Setter Property="BorderBrush"      Value="#A2BFA6"/>
            <Setter Property="BorderThickness"  Value="1"/>
            <Setter Property="FontFamily"       Value="Segoe UI"/>
            <Setter Property="FontSize"         Value="11"/>
            <Setter Property="Padding"          Value="10,9"/>
            <Setter Property="Cursor"           Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#C9A054"/>
                                <Setter TargetName="bd" Property="Background"  Value="#1E2C2D"/>
                                <Setter Property="Foreground" Value="#C9A054"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#162020"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── Small utility button (Browse / Add / Remove) ── -->
        <Style x:Key="SmallBtn" TargetType="Button">
            <Setter Property="Background"       Value="#1A2526"/>
            <Setter Property="Foreground"       Value="#A2BFA6"/>
            <Setter Property="BorderBrush"      Value="#3A4A4B"/>
            <Setter Property="BorderThickness"  Value="1"/>
            <Setter Property="FontFamily"       Value="Segoe UI"/>
            <Setter Property="FontSize"         Value="11"/>
            <Setter Property="Padding"          Value="10,7"/>
            <Setter Property="Cursor"           Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#C9A054"/>
                                <Setter Property="Foreground" Value="#C9A054"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── Run All button (coral green, full-width CTA) ── -->
        <Style x:Key="RunAllBtn" TargetType="Button">
            <Setter Property="Background"      Value="#4A8553"/>
            <Setter Property="Foreground"      Value="#FFFFFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="12,14"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#5A9863"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#3A6A43"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── Dark TextBox ── -->
        <Style x:Key="DarkTB" TargetType="TextBox">
            <Setter Property="Background"      Value="#1A2526"/>
            <Setter Property="Foreground"      Value="#A2BFA6"/>
            <Setter Property="BorderBrush"     Value="#3A4A4B"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="Padding"         Value="8,7"/>
            <Setter Property="CaretBrush"      Value="#C9A054"/>
            <Setter Property="SelectionBrush"  Value="#C9A054"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost"
                                          Margin="{TemplateBinding Padding}"
                                          VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#C9A054"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="BorderBrush" Value="#5A6A5E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── Dark ListBox ── -->
        <Style x:Key="DarkLB" TargetType="ListBox">
            <Setter Property="Background"      Value="#1A2526"/>
            <Setter Property="Foreground"      Value="#A2BFA6"/>
            <Setter Property="BorderBrush"     Value="#3A4A4B"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="Padding"         Value="4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ScrollViewer HorizontalScrollBarVisibility="Auto"
                                          VerticalScrollBarVisibility="Auto">
                                <ItemsPresenter/>
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Resources>
                <Style TargetType="ListBoxItem">
                    <Setter Property="Foreground" Value="#A2BFA6"/>
                    <Setter Property="Padding"    Value="6,4"/>
                    <Setter Property="Background" Value="Transparent"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="ListBoxItem">
                                <Border x:Name="ib"
                                        Background="{TemplateBinding Background}"
                                        CornerRadius="3"
                                        Padding="{TemplateBinding Padding}">
                                    <ContentPresenter/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="ib" Property="Background" Value="#2A3A3B"/>
                                        <Setter Property="Foreground" Value="#C9A054"/>
                                    </Trigger>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="ib" Property="Background" Value="#1E2C2D"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </Style.Resources>
        </Style>

        <!-- ── Gold outline button — log panel actions (Swab the Poop Deck, etc.) ── -->
        <Style x:Key="GoldBtn" TargetType="Button">
            <Setter Property="Background"      Value="Transparent"/>
            <Setter Property="Foreground"      Value="#C9A054"/>
            <Setter Property="BorderBrush"     Value="#C9A054"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="FontSize"        Value="11"/>
            <Setter Property="Padding"         Value="10,5"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background"  Value="#1E2C2D"/>
                                <Setter TargetName="bd" Property="BorderBrush" Value="#DDB064"/>
                                <Setter Property="Foreground" Value="#DDB064"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#152020"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ── ScrollBar — self-contained style, fully inlined to avoid StaticResource
             lookup failures inside ControlTemplates. RepeatButtons use Opacity=0 so they
             are invisible but fully hit-testable (enables click-to-page and thumb dragging). ── -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background"               Value="#1A2526"/>
            <Setter Property="Width"                    Value="12"/>
            <Setter Property="MinWidth"                 Value="12"/>
            <Setter Property="Stylus.IsFlicksEnabled"   Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Opacity="0" Focusable="False"
                                                  Command="ScrollBar.PageUpCommand"/>
                                </Track.DecreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border x:Name="tb" CornerRadius="3"
                                                        Background="#4A8553" Margin="1" Opacity="0.8"/>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsMouseOver" Value="True">
                                                        <Setter TargetName="tb" Property="Background" Value="#5A9863"/>
                                                        <Setter TargetName="tb" Property="Opacity"    Value="1"/>
                                                    </Trigger>
                                                    <Trigger Property="IsDragging" Value="True">
                                                        <Setter TargetName="tb" Property="Background" Value="#C9A054"/>
                                                        <Setter TargetName="tb" Property="Opacity"    Value="1"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Opacity="0" Focusable="False"
                                                  Command="ScrollBar.PageDownCommand"/>
                                </Track.IncreaseRepeatButton>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="Orientation" Value="Horizontal">
                    <Setter Property="Height"    Value="12"/>
                    <Setter Property="MinHeight" Value="12"/>
                    <Setter Property="Width"     Value="Auto"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="ScrollBar">
                                <Grid Background="{TemplateBinding Background}">
                                    <Track x:Name="PART_Track">
                                        <Track.DecreaseRepeatButton>
                                            <RepeatButton Opacity="0" Focusable="False"
                                                          Command="ScrollBar.PageLeftCommand"/>
                                        </Track.DecreaseRepeatButton>
                                        <Track.Thumb>
                                            <Thumb>
                                                <Thumb.Template>
                                                    <ControlTemplate TargetType="Thumb">
                                                        <Border x:Name="tb" CornerRadius="3"
                                                                Background="#4A8553" Margin="1" Opacity="0.8"/>
                                                        <ControlTemplate.Triggers>
                                                            <Trigger Property="IsMouseOver" Value="True">
                                                                <Setter TargetName="tb" Property="Background" Value="#5A9863"/>
                                                                <Setter TargetName="tb" Property="Opacity"    Value="1"/>
                                                            </Trigger>
                                                            <Trigger Property="IsDragging" Value="True">
                                                                <Setter TargetName="tb" Property="Background" Value="#C9A054"/>
                                                                <Setter TargetName="tb" Property="Opacity"    Value="1"/>
                                                            </Trigger>
                                                        </ControlTemplate.Triggers>
                                                    </ControlTemplate>
                                                </Thumb.Template>
                                            </Thumb>
                                        </Track.Thumb>
                                        <Track.IncreaseRepeatButton>
                                            <RepeatButton Opacity="0" Focusable="False"
                                                          Command="ScrollBar.PageRightCommand"/>
                                        </Track.IncreaseRepeatButton>
                                    </Track>
                                </Grid>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>

    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="72"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- ══════════════════════════════════════ HEADER ══════════════════════════════════════ -->
        <Border x:Name="headerBorder" Grid.Row="0" Background="#1A1D1E" BorderBrush="#2A3A3B" BorderThickness="0,0,0,1">
            <DockPanel Margin="18,0" VerticalAlignment="Center">

                <!-- Badge icon -->
                <Border DockPanel.Dock="Left" Width="52" Height="52"
                        Margin="0,0,14,0" VerticalAlignment="Center">
                    <Image x:Name="headerIcon" Stretch="Uniform"
                           RenderOptions.BitmapScalingMode="HighQuality"/>
                </Border>

                <!-- Window controls — right-docked in reverse visual order (Close first = rightmost) -->

                <!-- CLOSE -->
                <Button x:Name="closeButton" DockPanel.Dock="Right"
                        Width="44" Height="44" Cursor="Hand" VerticalAlignment="Center"
                        ToolTip="Close Shipwreck">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="Transparent" CornerRadius="4" Width="44" Height="44">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="9"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#3A1010"/>
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#5A1818"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                    <Image Width="26" Height="26" RenderOptions.BitmapScalingMode="HighQuality"/>
                </Button>

                <!-- MAXIMIZE -->
                <Button x:Name="maximizeButton" DockPanel.Dock="Right"
                        Width="44" Height="44" Cursor="Hand" VerticalAlignment="Center"
                        ToolTip="Maximize / Restore">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="Transparent" CornerRadius="4" Width="44" Height="44">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="9"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#2A3A3B"/>
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#1A2A2B"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                    <Image Width="26" Height="26" RenderOptions.BitmapScalingMode="HighQuality"/>
                </Button>

                <!-- MINIMIZE -->
                <Button x:Name="minimizeButton" DockPanel.Dock="Right"
                        Width="44" Height="44" Cursor="Hand" VerticalAlignment="Center"
                        ToolTip="Minimize">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="bd" Background="Transparent" CornerRadius="4" Width="44" Height="44">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="9"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#2A3A3B"/>
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="bd" Property="Background" Value="#1A2A2B"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                    <Image Width="26" Height="26" RenderOptions.BitmapScalingMode="HighQuality"/>
                </Button>

                <!-- Title -->
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock x:Name="titleLabel" Text="SHIPWRECK" FontFamily="Segoe UI" FontSize="26"
                               FontWeight="Bold" Foreground="#C9A054" VerticalAlignment="Center"/>
                    <TextBlock x:Name="versionLabel" FontFamily="Segoe UI" FontSize="18"
                               Foreground="#A2BFA6" VerticalAlignment="Center" Margin="4,3,0,0"/>
                </StackPanel>

            </DockPanel>
        </Border>

        <!-- ══════════════════════════════════ TWO-COLUMN BODY ══════════════════════════════════ -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2.2*"/>
                <ColumnDefinition Width="2.8*"/>
            </Grid.ColumnDefinitions>

            <!-- ─────────────────────── LEFT: CONTROL MANIFEST ─────────────────────── -->
            <Border Grid.Column="0" Background="#253233" BorderBrush="#2A3A3B" BorderThickness="0,0,1,0">
                <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                              Background="Transparent">
                    <StackPanel Margin="18,18,18,18">

                        <!-- DIRECTORIES -->
                        <TextBlock Text="📁   DIRECTORIES" FontFamily="Segoe UI" FontSize="10.5"
                                   FontWeight="Bold" Foreground="#C9A054" Margin="0,0,0,12"/>

                        <TextBlock Text="Source Folder:" FontFamily="Segoe UI" FontSize="11"
                                   Foreground="#A2BFA6" Margin="0,0,0,5"/>
                        <Grid Margin="0,0,0,16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="sourceTextBox" Grid.Column="0" Style="{StaticResource DarkTB}"/>
                            <Button  x:Name="browseSourceButton" Grid.Column="1" Content="Browse..."
                                     Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
                        </Grid>

                        <TextBlock Text="Target Folders (Media Library Roots):" FontFamily="Segoe UI"
                                   FontSize="11" Foreground="#A2BFA6" Margin="0,0,0,5"/>
                        <ListBox x:Name="targetListBox" Height="130" Margin="0,0,0,8"
                                 Style="{StaticResource DarkLB}"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Button x:Name="addTargetButton"    Grid.Column="0" Content="＋  Add Target"
                                    Style="{StaticResource SmallBtn}"/>
                            <Button x:Name="removeTargetButton" Grid.Column="1" Content="－  Remove Target"
                                    Style="{StaticResource SmallBtn}" Margin="8,0,0,0"/>
                        </Grid>

                        <!-- SEPARATOR -->
                        <Rectangle Height="1" Fill="#2A3A3B" Margin="0,22,0,22"/>

                        <!-- SALVAGE OPERATIONS -->
                        <TextBlock Text="☠   SALVAGE OPERATIONS" FontFamily="Segoe UI" FontSize="10.5"
                                   FontWeight="Bold" Foreground="#C9A054" Margin="0,0,0,12"/>

                        <Grid Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Button x:Name="extractButton"        Grid.Column="0"
                                    Content="1. Extract Archives"  Style="{StaticResource StepBtn}"/>
                            <Button x:Name="renameBackdropButton" Grid.Column="1"
                                    Content="2. Rename Backdrops"  Style="{StaticResource StepBtn}"
                                    Margin="8,0,0,0"/>
                        </Grid>

                        <Grid Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Button x:Name="renameSeasonButton"  Grid.Column="0"
                                    Content="3. Rename Media Images"   Style="{StaticResource StepBtn}"/>
                            <Button x:Name="renameEpisodeButton" Grid.Column="1"
                                    Content="4. Rename Episode Images" Style="{StaticResource StepBtn}"
                                    Margin="8,0,0,0"/>
                        </Grid>

                        <Button x:Name="moveButton" Content="5. Move Images &amp; Cleanup"
                                Style="{StaticResource StepBtn}" Margin="0,0,0,22"/>

                        <!-- RUN ALL -->
                        <Button x:Name="runAllButton" Content="🚀   RUN ALL STEPS (1-5)"
                                Style="{StaticResource RunAllBtn}"/>

                    </StackPanel>
                </ScrollViewer>
            </Border>

            <!-- ─────────────────────── RIGHT: SALVAGE LOG STREAM ─────────────────────── -->
            <Border Grid.Column="1" Background="#0F1213">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="#0F1213" BorderBrush="#2A3A3B"
                            BorderThickness="0,0,0,1" Padding="16,8,12,8">
                        <DockPanel VerticalAlignment="Center">
                            <!-- Button docked right first so label fills remaining left space -->
                            <Button x:Name="swabButton" DockPanel.Dock="Right"
                                    Content="Swab the Poop Deck"
                                    Style="{StaticResource GoldBtn}"
                                    VerticalAlignment="Center"/>
                            <TextBlock Text="📟   SALVAGE LOG STREAM" FontFamily="Segoe UI"
                                       FontSize="10.5" FontWeight="Bold" Foreground="#C9A054"
                                       VerticalAlignment="Center"/>
                        </DockPanel>
                    </Border>

                    <RichTextBox x:Name="logTextBox" Grid.Row="1"
                                 Background="#0F1213" Foreground="#A2BFA6"
                                 FontFamily="Consolas" FontSize="10.5"
                                 IsReadOnly="True" BorderThickness="0"
                                 HorizontalScrollBarVisibility="Auto"
                                 VerticalScrollBarVisibility="Auto"
                                 Padding="14,12,14,12"/>
                </Grid>
            </Border>

        </Grid>
    </Grid>
</Window>
"@

# Load XAML and get the window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Get named controls
$sourceTextBox        = $window.FindName("sourceTextBox")
$browseSourceButton   = $window.FindName("browseSourceButton")
$targetListBox        = $window.FindName("targetListBox")
$addTargetButton      = $window.FindName("addTargetButton")
$removeTargetButton   = $window.FindName("removeTargetButton")
$extractButton        = $window.FindName("extractButton")
$renameBackdropButton = $window.FindName("renameBackdropButton")
$renameSeasonButton   = $window.FindName("renameSeasonButton")
$renameEpisodeButton  = $window.FindName("renameEpisodeButton")
$moveButton           = $window.FindName("moveButton")
$runAllButton         = $window.FindName("runAllButton")
$headerIcon           = $window.FindName("headerIcon")
$headerBorder         = $window.FindName("headerBorder")
$closeButton          = $window.FindName("closeButton")
$maximizeButton       = $window.FindName("maximizeButton")
$minimizeButton       = $window.FindName("minimizeButton")
$titleLabel           = $window.FindName("titleLabel")
$versionLabel         = $window.FindName("versionLabel")
$script:logTextBox    = $window.FindName("logTextBox")
$swabButton           = $window.FindName("swabButton")

# Make the header draggable (excluding the button area — buttons handle their own clicks)
$headerBorder.Add_MouseLeftButtonDown({ $window.DragMove() })

# Window control buttons
$closeButton.Add_Click({    $window.Close() })
$minimizeButton.Add_Click({ $window.WindowState = [System.Windows.WindowState]::Minimized })
$maximizeButton.Add_Click({
    if ($window.WindowState -eq [System.Windows.WindowState]::Maximized) {
        $window.WindowState = [System.Windows.WindowState]::Normal
    } else {
        $window.WindowState = [System.Windows.WindowState]::Maximized
    }
})

# Clear the log
$swabButton.Add_Click({ $script:logTextBox.Document.Blocks.Clear() })

# Disable text wrapping in the log (PageWidth >> visible width = no wrap, horizontal scrollbar appears)
$script:logTextBox.Document.PageWidth = 8000

# Set version label
$versionLabel.Text = " $($script:AppVersion)"

# Apply WindowChrome — removes the thin OS title strip that WindowStyle=None leaves behind,
# while restoring native edge-resize (which CanResizeWithGrip alone can't fully clean up).
try {
    $chrome = New-Object System.Windows.Shell.WindowChrome
    $chrome.ResizeBorderThickness = New-Object System.Windows.Thickness(6, 6, 6, 6)
    $chrome.CaptionHeight         = 0      # entire window is client area; drag handled by header code
    $chrome.CornerRadius          = New-Object System.Windows.CornerRadius(0)
    $chrome.GlassFrameThickness   = New-Object System.Windows.Thickness(0)
    [System.Windows.Shell.WindowChrome]::SetWindowChrome($window, $chrome)
} catch { Write-Warning "WindowChrome not available: $($_.Exception.Message)" }

# Collect all buttons for bulk enable/disable during operations
$script:allButtons = @(
    $extractButton, $renameBackdropButton, $renameSeasonButton,
    $renameEpisodeButton, $moveButton, $runAllButton,
    $browseSourceButton, $addTargetButton, $removeTargetButton,
    $swabButton
)

# --- Button click handlers ---
$browseSourceButton.Add_Click({
    $selectedPath = Select-FolderDialog -Description "Select the Source Folder" -InitialDirectory $sourceTextBox.Text
    if ($selectedPath) {
        $sourceTextBox.Text = $selectedPath
        $script:sourcePath  = $selectedPath
    }
})

$addTargetButton.Add_Click({
    $selectedPath = Select-FolderDialog -Description "Add a Target Media Folder"
    if ($selectedPath) {
        $exists = $false
        foreach ($item in $targetListBox.Items) {
            if ($item -eq $selectedPath) { $exists = $true; break }
        }
        if (-not $exists) {
            $targetListBox.Items.Add($selectedPath) | Out-Null
            $script:targetPaths.Add($selectedPath)
        } else {
            [System.Windows.MessageBox]::Show(
                "This path is already in the target list.",
                "Duplicate Path",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
    }
})

$removeTargetButton.Add_Click({
    $selectedItem = $targetListBox.SelectedItem
    if ($null -ne $selectedItem) {
        $targetListBox.Items.Remove($selectedItem)
        $found = $script:targetPaths | Where-Object { $_ -eq $selectedItem } | Select-Object -First 1
        if ($found) { $script:targetPaths.Remove($found) }
    } else {
        [System.Windows.MessageBox]::Show(
            "Please select a target path from the list to remove.",
            "No Selection",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
    }
})

$extractButton.Add_Click({
    Process-Action -Action { Extract-MatchingArchives -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap }
})

$renameBackdropButton.Add_Click({
    Process-Action -Action { Rename-ExistingBackdrops -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap }
})

$renameSeasonButton.Add_Click({
    Process-Action -Action { Rename-SeasonPosters -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap }
})

$renameEpisodeButton.Add_Click({
    Process-Action -Action { Rename-EpisodeImages -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap }
})

$moveButton.Add_Click({
    Process-Action -Action { Move-ImageFilesToServer -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap }
})

$runAllButton.Add_Click({
    Process-Action -Action {
        Add-LogEntry "--- Setting Sail — All Steps ---" -ColorInput ([System.Drawing.Color]::DarkMagenta)
        Extract-MatchingArchives  -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap
        Rename-ExistingBackdrops  -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap
        Rename-SeasonPosters      -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap
        Rename-EpisodeImages      -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap
        Move-ImageFilesToServer   -SourceRoot $script:sourcePath -TargetRoots $script:targetPaths -PrebuiltMap $script:cachedTargetMap
        Add-LogEntry "⚓ All cargo delivered!" -ColorInput ([System.Drawing.Color]::DarkMagenta)
    }
})


# --- Helper Function to Wrap Button Actions ---
function Process-Action {
    param([scriptblock]$Action)

    # Sync paths from GUI
    $script:sourcePath = $sourceTextBox.Text
    $script:targetPaths.Clear()
    $targetListBox.Items | ForEach-Object { $script:targetPaths.Add($_) }

    # Input validation
    if ([string]::IsNullOrWhiteSpace($script:sourcePath) -or (-not (Test-Path $script:sourcePath -PathType Container))) {
        Add-LogEntry "Error: Source path is not set or does not exist." -ColorInput ([System.Drawing.Color]::Red)
        [System.Windows.MessageBox]::Show("Please select a valid, existing Source Folder first.", "Input Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return
    }
    if ($script:targetPaths.Count -eq 0) {
        Add-LogEntry "Error: No target paths are specified." -ColorInput ([System.Drawing.Color]::Red)
        [System.Windows.MessageBox]::Show("Please add at least one Target Folder first.", "Input Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return
    }
    $allTargetsValid = $true
    foreach ($tp in $script:targetPaths) {
        if (-not (Test-Path $tp -PathType Container)) {
            Add-LogEntry "Error: Target path not found: '$tp'" -ColorInput ([System.Drawing.Color]::Red)
            $allTargetsValid = $false
        }
    }
    if (-not $allTargetsValid) {
        [System.Windows.MessageBox]::Show("One or more Target Folders do not exist. Please check the list.", "Input Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return
    }

    # 7-Zip check for extraction
    if ($Action.ToString() -match 'Extract-MatchingArchives') {
        $script:sevenZipPath = Test-7ZipPath
        if (-not $script:sevenZipPath) {
            Add-LogEntry "Error: 7z.exe not found. Cannot salvage archives." -ColorInput ([System.Drawing.Color]::Red)
            [System.Windows.MessageBox]::Show("7-Zip (7z.exe) not found in PATH or common locations.", "Dependency Error",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        } else {
            Add-LogEntry "7-Zip found: $script:sevenZipPath" -ColorInput ([System.Drawing.Color]::DarkGray)
        }
    }

    # Disable buttons and show wait cursor
    $script:allButtons | ForEach-Object { $_.IsEnabled = $false }
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    try {
        # Build target map once
        Add-LogEntry "🤿 Diving for shows across $($script:targetPaths.Count) port(s)..." -ColorInput ([System.Drawing.Color]::Blue)
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
        $script:cachedTargetMap = Get-TargetShowFoldersMap -TargetBasePaths $script:targetPaths
        if ($script:cachedTargetMap.Count -eq 0) {
            Add-LogEntry "❌ No valid 'Show Name (YYYY)' folders found in any port. Aborting." -ColorInput ([System.Drawing.Color]::Red)
            return
        }
        Add-LogEntry "⚓ Found $($script:cachedTargetMap.Count) show(s) across $($script:targetPaths.Count) port(s)." -ColorInput ([System.Drawing.Color]::Green)

        # Fresh episode routing map
        $script:episodeTargetDirMap = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)

        # Normalize source filenames once
        Add-LogEntry "🔄 Normalizing source filenames..." -ColorInput ([System.Drawing.Color]::Blue)
        Trim-ImageFilenamesInFolder -sourcePath $script:sourcePath

        # Execute the step(s)
        Invoke-Command -ScriptBlock $Action -ErrorAction Stop

    } catch {
        Add-LogEntry "❌ An error occurred: $($_.Exception.Message)" -ColorInput ([System.Drawing.Color]::Red)
        if ($_.Exception.InnerException) {
            Add-LogEntry "   Inner: $($_.Exception.InnerException.Message)" -ColorInput ([System.Drawing.Color]::Red)
        }
        Add-LogEntry "   ScriptStackTrace: $($_.ScriptStackTrace)" -ColorInput ([System.Drawing.Color]::DarkRed)
    } finally {
        $script:allButtons | ForEach-Object { $_.IsEnabled = $true }
        $window.Cursor = $null
        Add-LogEntry "--------------------"
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }
}


# --- Window Loaded ---
$window.Add_Loaded({
    # Helper: find an icon file — checks assets subfolder first, then script root.
    # NOTE: Join-Path in PS5.1 only accepts 2 positional args, so we nest the calls.
    function Find-IconPath([string]$filename) {
        $subPath  = Join-Path (Join-Path $scriptDir "assets") $filename
        $rootPath = Join-Path $scriptDir $filename
        if (Test-Path $subPath  -PathType Leaf) { return $subPath  }
        if (Test-Path $rootPath -PathType Leaf) { return $rootPath }
        return $null
    }

    # Application badge icon (header image + taskbar)
    try {
        $iconPath = Find-IconPath "Shipwreck_Badge.ico"
        if ($iconPath) {
            # Use BitmapDecoder to read all embedded frames and pick the largest for best header quality
            $uri     = [System.Uri]::new($iconPath, [System.UriKind]::Absolute)
            $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
                $uri,
                [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                [System.Windows.Media.Imaging.BitmapCacheOption]::Default)
            $bestFrame = $decoder.Frames |
                Sort-Object { $_.PixelWidth } -Descending |
                Select-Object -First 1
            $headerIcon.Source = $bestFrame

            # Also set the taskbar/window icon
            $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create($uri)
        }
    } catch { <# Non-fatal #> }

    # Window control button icons (close, maximize, minimize)
    # Each button's Content is an Image set here — same approach as the badge icon
    foreach ($pair in @(
        @{ Button = $closeButton;    File = "close.ico"    },
        @{ Button = $maximizeButton; File = "maximize.ico" },
        @{ Button = $minimizeButton; File = "minimize.ico" }
    )) {
        try {
            $iconPath = Find-IconPath $pair.File
            if ($iconPath -and $pair.Button.Content -is [System.Windows.Controls.Image]) {
                $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                $bmp.BeginInit()
                $bmp.UriSource = [System.Uri]::new($iconPath, [System.UriKind]::Absolute)
                $bmp.EndInit()
                $pair.Button.Content.Source = $bmp
            }
        } catch { <# Non-fatal #> }
    }

    # --- Title font ---
    # Loads Porto_Buena.otf from the assets folder and applies it to the header title.
    # The #fragment must match the internal family name stored in the font file.
    # If the font renders incorrectly, open Porto_Buena.otf in Windows Font Viewer —
    # the large name shown at the top of that window is the exact string to use after #.
    try {
        $fontFile = Find-IconPath "Porto_Buena.otf"
        if ($fontFile) {
            $fontDir    = "file:///" + (Split-Path $fontFile -Parent).Replace("\", "/") + "/"
            $portoBuena = New-Object System.Windows.Media.FontFamily(
                [System.Uri]::new($fontDir), "#Porto Buena")
            $titleLabel.FontFamily   = $portoBuena
            $versionLabel.FontFamily = $portoBuena
        } else {
            Add-LogEntry "⚠️ Porto_Buena.otf not found in assets folder — using default font." -ColorInput ([System.Drawing.Color]::Orange)
        }
    } catch { <# Non-fatal — falls back to Segoe UI #> }

    # 7-Zip check
    $script:sevenZipPath = Test-7ZipPath
    if (-not $script:sevenZipPath) {
        Add-LogEntry "⚠️ 7z.exe not found. Archive salvage (.rar/.7z) will fail." -ColorInput ([System.Drawing.Color]::Orange)
    } else {
        Add-LogEntry "7-Zip found: $script:sevenZipPath" -ColorInput ([System.Drawing.Color]::Green)
    }

    # Load saved config
    Load-Configuration
    $sourceTextBox.Text = if ($script:sourcePath) { $script:sourcePath } else { "" }
    $targetListBox.Items.Clear()
    if ($script:targetPaths -ne $null -and $script:targetPaths.Count -gt 0) {
        $script:targetPaths | ForEach-Object { $targetListBox.Items.Add($_) }
    }

    Add-LogEntry "Shipwreck ready. Select your hold and ports, then dive." -ColorInput ([System.Drawing.Color]::DarkGreen)
    $window.Activate()
})


# --- Window Closing ---
$window.Add_Closing({
    $script:sourcePath = $sourceTextBox.Text
    $script:targetPaths.Clear()
    $targetListBox.Items | ForEach-Object { $script:targetPaths.Add($_) }
    Save-Configuration

    # Save log to shipwreck.log
    try {
        $textRange = New-Object System.Windows.Documents.TextRange(
            $script:logTextBox.Document.ContentStart,
            $script:logTextBox.Document.ContentEnd)
        $textRange.Text | Out-File -FilePath $LogFilePath -Encoding UTF8 -Force -ErrorAction Stop
    } catch {
        Write-Warning "Could not save log to '$LogFilePath': $($_.Exception.Message)"
    }
})


# --- Launch ---
if ($Host.Runspace.ApartmentState -ne 'STA') {
    Write-Warning "Host is not in STA mode. WPF requires STA. Relaunching..."
    if (-not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
        try {
            Start-Process powershell -ArgumentList "-Sta", "-File", "`"$($MyInvocation.MyCommand.Path)`"", "-ExecutionPolicy", "Bypass" -ErrorAction Stop
            exit
        } catch {
            Write-Error "Failed to relaunch in STA mode: $($_.Exception.Message)"
        }
    }
}

try {
    $window.ShowDialog() | Out-Null
} catch {
    Write-Error "An error occurred displaying the window: $($_.Exception.Message)"
    [System.Windows.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Launch Error",
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
}

# --- Script End ---
Write-Host "Shipwreck closed."
