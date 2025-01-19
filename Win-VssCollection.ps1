<#
##################################################################################################################
# Win-VssCollector.ps1
# v1.2
##################################################################################################################

 _       ___           _    __          ______      ____          __            
| |     / (_)___      | |  / /_________/ ____/___  / / /__  _____/ /_____  _____
| | /| / / / __ \_____| | / / ___/ ___/ /   / __ \/ / / _ \/ ___/ __/ __ \/ ___/
| |/ |/ / / / / /_____/ |/ (__  |__  ) /___/ /_/ / / /  __/ /__/ /_/ /_/ / /    
|__/|__/_/_/ /_/      |___/____/____/\____/\____/_/_/\___/\___/\__/\____/_/     

.SYNOPSIS
    This script performs a backup of system artefacts as a ZIP archive, 
    including open files using the Volume Shadow Copy Service (VSS). The
    archive is then moved to the target directory.

.NOTES
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
    PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
    HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
    OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
    SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

# Define source files and directories
$source_files = @(
    "C:\Windows\System32\config\BBI"
    "C:\Windows\System32\config\COMPONENTS"
    "C:\Windows\System32\config\DRIVERS"
    "C:\Windows\System32\config\ELAM"
    "C:\Windows\System32\config\SAM"
    "C:\Windows\System32\config\SECURITY"
    "C:\Windows\System32\config\SOFTWARE"
    "C:\Windows\System32\config\SYSTEM"
    "C:\Windows\System32\sru\SRUDB.dat"
)

$source_dirs = @(
    "C:\Windows\appcompat\Programs"
    "C:\Windows\System32\winevt\Logs"
    "C:\Windows\Prefetch"
    "C:\Windows\System32\LogFiles\SUM"
)

###############################################################################
# Collect artefacts with dynamic paths
###############################################################################

# Get the list of user accounts as an array
$user_accounts = @(Get-ChildItem C:\Users -Name -Force)

foreach ($user in $user_accounts) {
    $wildcardPath = "C:\Users\$user\NTUSER.dat*"
    $matchedFiles = Get-ChildItem -Path $wildcardPath -Force -File -ErrorAction SilentlyContinue

    foreach ($matchedFile in $matchedFiles) {
        $source_files += @{
            Path      = $matchedFile.FullName
            Renamed   = $matchedFile.Name
            Subfolder = "$user`_NTUSER"
        }
    }
}

foreach ($user in $user_accounts) {
    $wildcardPath = "C:\Users\$user\AppData\Local\Microsoft\Windows\UsrClass.dat*"
    $matchedFiles = Get-ChildItem -Path $wildcardPath -Force -File -ErrorAction SilentlyContinue

    foreach ($matchedFile in $matchedFiles) {
        $source_files += @{
            Path    = $matchedFile.FullName
            Renamed = "$user`_$($matchedFile.Name)"
            Subfolder = "$user`_UsrClass"
        }
    }
}

foreach ($user in $user_accounts) {
    $source_files += @{
        Path    = "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data\Default\History"
        Renamed = "$user`_EdgeHistory"
        Subfolder = "$user`_EdgeHistory"
    }
}

foreach ($user in $user_accounts) {
    $source_dirs += @{
        Path    = "C:\Users\$user\AppData\Roaming\Microsoft\Windows\Recent"
        Renamed = "$user`_LnkFiles"
        Subfolder = "$user`_LnkFiles"
    }
}

###############################################################################
# Function: Write-Log
###############################################################################

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
}

###############################################################################

$target_dir = "$env:systemdrive\nmc-edr-lr\vss-export"
$date = Get-Date -Format yyyy-MM-dd

# Ensure target directory exists
if (-Not (Test-Path -Path $target_dir)) {
    Write-Log "Target directory does not exist. Creating: $target_dir" -Level "Info"
    New-Item -Path $target_dir -ItemType Directory -Force | Out-Null
}

$temp_shadow_link = "C:\shadowcopy_$date"
$unix_time = Get-Date -UFormat %s -Millisecond 0
$archive_filename = "vss_$((Get-Date -Format 'yyyyMMdd-HHmmss')).zip"
$temp_archive_full_path = Join-Path -Path $env:TEMP -ChildPath $archive_filename
$temp_collected_dir = Join-Path -Path $env:TEMP -ChildPath "CollectedFiles_$date"

# Create a temporary directory for collected files
if (-Not (Test-Path -Path $temp_collected_dir)) {
    New-Item -Path $temp_collected_dir -ItemType Directory -Force | Out-Null
}

Write-Log "Creating new shadow copy snapshot." -Level "Info"
$s1 = (Get-WmiObject -List Win32_ShadowCopy).Create("C:\", "ClientAccessible")
$s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $s1.ShadowID }
$d  = $s2.DeviceObject + "\"

if (Test-Path -Path $temp_shadow_link) {
    Write-Log "Temporary shadow link exists. Deleting: $temp_shadow_link" -Level "Warning"
    Remove-Item -Path $temp_shadow_link -Recurse -Force
}
cmd /c mklink /d $temp_shadow_link $d

# -----------------------------------------------------------------------------
# Collect files into the temporary directory (with subfolders where specified)
# -----------------------------------------------------------------------------

foreach ($file in $source_files) {

    if ($file -is [Hashtable]) {

        # Convert the original "C:\..." path into the shadow path
        $shadow_file = $file.Path -replace "^C:", $temp_shadow_link

        if (Test-Path -Path $shadow_file) {
            # If a Subfolder is specified, create it
            $destinationDir = $temp_collected_dir
            if ($file.ContainsKey("Subfolder")) {
                $destinationDir = Join-Path $temp_collected_dir $file.Subfolder
                if (-not (Test-Path $destinationDir)) {
                    New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
                }
            }

            # Construct the final file path
            $destinationFile = Join-Path $destinationDir $file.Renamed

            Write-Log "Copying file: $shadow_file -> $destinationFile" -Level "Info"
            Copy-Item -Path $shadow_file -Destination $destinationFile -Force
        }
        else {
            Write-Log "File not found in shadow copy: $shadow_file" -Level "Warning"
        }

    }
    else {
        # Handle any other regular string paths in $source_files
        $shadow_file = $file -replace "^C:", $temp_shadow_link

        if (Test-Path -Path $shadow_file) {
            Write-Log "Copying file: $shadow_file" -Level "Info"
            Copy-Item -Path $shadow_file -Destination $temp_collected_dir -Force
        }
        else {
            Write-Log "File not found in shadow copy: $shadow_file" -Level "Warning"
        }
    }
}

# -----------------------------------------------------------------------------
# Collect directories into the temporary directory (with rename if specified)
# -----------------------------------------------------------------------------

foreach ($dir in $source_dirs) {

    if ($dir -is [Hashtable]) {
        # Handle directories with rename
        $shadow_dir = $dir.Path -replace "^C:", $temp_shadow_link
        $renamed_dir = Join-Path -Path $temp_collected_dir -ChildPath $dir.Renamed

        if (Test-Path -Path $shadow_dir) {
            Write-Log "Copying and renaming directory: $shadow_dir -> $renamed_dir" -Level "Info"
            Copy-Item -Path $shadow_dir -Destination $renamed_dir -Recurse -Force
        } else {
            Write-Log "Directory not found in shadow copy: $shadow_dir" -Level "Warning"
        }
    }
    else {
        # Handle regular directories
        $shadow_dir = $dir -replace "^C:", $temp_shadow_link

        if (Test-Path -Path $shadow_dir) {
            Write-Log "Copying directory: $shadow_dir" -Level "Info"
            Copy-Item -Path $shadow_dir -Destination $temp_collected_dir -Recurse -Force
        } else {
            Write-Log "Directory not found in shadow copy: $shadow_dir" -Level "Warning"
        }
    }
}

# -----------------------------------------------------------------------------
# Create the ZIP archive using System.IO.Compression.ZipFile
# -----------------------------------------------------------------------------

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Write-Log "Creating ZIP archive: $temp_archive_full_path" -Level "Info"
    [System.IO.Compression.ZipFile]::CreateFromDirectory($temp_collected_dir, $temp_archive_full_path)
}
catch {
    Write-Log "Failed to create ZIP archive. Error: $_" -Level "Error"
}

# -----------------------------------------------------------------------------
# Clean up shadow snapshot, temporary link, and collected files
# -----------------------------------------------------------------------------

Write-Log "Deleting shadow snapshot and cleaning up temporary link." -Level "Info"
$s2.Delete()
cmd /c rmdir $temp_shadow_link

Write-Log "Cleaning up temporary collected files directory." -Level "Info"
Remove-Item -Path $temp_collected_dir -Recurse -Force

# Move the archive to the target directory
Write-Log "Moving archive to target directory: $target_dir" -Level "Info"
robocopy $env:TEMP $target_dir $archive_filename /MOVE /R:3 /np

Write-Log "Backup completed successfully at $(Get-Date -Format yyyy-MM-dd_HH:mm:ss)." -Level "Info"
Write-Log "ZIP archive directory: $target_dir\$archive_filename" -Level "Info"
